"""
PDF Steel Schedule Extractor
============================
Features
--------

- Scan only the bottom-left area of the drawing.
- Export ALL detected tables in that area to Excel.
- Debug mode draws detected table boxes onto PNG files.
- Easy to tune with CONFIG.
- Supports:
    - Single PDF
    - Entire folder
"""

from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import pdfplumber
import pandas as pd
import fitz

from PIL import Image
from PIL import ImageDraw


# ==============================================================================
# CONFIG
# ==============================================================================

CONFIG = {
    # ------------------------------------------------------------------
    # Input
    # ------------------------------------------------------------------
    "input_mode": "folder",  # "single" or "folder"
    "input_pdf": r"D:\Working\Draft\input\SUB1-[0].pdf",
    "input_folder": r"D:\Working\Draft\input",
    # File name only: the file is saved inside each Run_* folder.
    "output_excel": "output.xlsx",
    # ------------------------------------------------------------------
    # Run history
    # ------------------------------------------------------------------
    "run_root": r"D:\Working\Draft",
    # ------------------------------------------------------------------
    # Debug
    # (debug images are saved in <run folder>/debug)
    # ------------------------------------------------------------------
    "debug": True,
    "export_debug_images": True,
    # ------------------------------------------------------------------
    # Search region
    #
    # Coordinates are ratios (0.00–1.00) of page width/height.
    #
    # Example:
    #
    # left   = 0.00
    # right  = 0.40
    #
    # top    = 0.60
    # bottom = 1.00
    #
    # Means:
    #
    # 40% width from left
    # lower 40% of page
    # ------------------------------------------------------------------
    "search_region": {
        "left": 0.00,
        "right": 0.4,
        "top": 0.60,
        "bottom": 1.00,
    },
    # ------------------------------------------------------------------
    # Table extraction settings
    # Passed directly to pdfplumber's find_tables().
    # Adjust if tables are missed or garbage is detected.
    # ------------------------------------------------------------------
    "table_settings": {
        # Use detected lines ("lines") or text gaps ("text") to find edges.
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        # How far a line endpoint can be from another to snap together.
        "snap_tolerance": 5,
        # How far two line segments can be apart to join into one edge.
        "join_tolerance": 5,
        # How close an intersection must be to a line to count.
        "intersection_tolerance": 8,
        # Minimum line length to consider as a table edge (pixels).
        "edge_min_length": 10,
    },
    # ------------------------------------------------------------------
    # Basic filters
    # Reject tables that do not meet these minimums.
    # Raise values if small garbage tables are still appearing.
    # ------------------------------------------------------------------
    "min_rows": 3,  # Minimum number of non-empty rows
    "min_cols": 2,  # Minimum number of columns
    "min_cells": 8,  # Minimum total cells (rows × cols)
    # ------------------------------------------------------------------
    # Output
    # ------------------------------------------------------------------
    "sheet_name": "Steel Schedule",
    "scoring": {
        # Ưu tiên bảng sát góc trái
        "left_weight": 40,
        # Ưu tiên bảng sát đáy
        "bottom_weight": 40,
        # Bảng rộng
        "aspect_weight": 30,
        # Nhiều ô
        "cell_weight": 20,
        # Nhiều text
        "text_weight": 20,
        # Phạt bảng quá cao
        "height_penalty": 50,
        # Ngưỡng
        "min_score": 80,
        # Ngưỡng phạt chiều cao (tỷ lệ chiều cao trang)
        "height_threshold": 0.20,
        # Ngưỡng phạt bảng dọc (tỷ lệ rộng/cao)
        "vertical_threshold": 1.2,
        # Điểm phạt bảng dọc
        "vertical_penalty": 100,
    },
}


# ==============================================================================
# HELPERS
# ==============================================================================


def log(*args):

    if CONFIG["debug"]:
        print(*args)


def normalize_cell(value):

    if value is None:
        return ""

    value = str(value)

    # Drop control characters; Excel export fails on them.
    value = "".join(ch for ch in value if ch.isprintable() or ch.isspace())

    value = value.replace("\n", " ")

    value = " ".join(value.split())

    return value.strip()


def clean_table(table):

    rows = []

    max_cols = 0

    for row in table:
        row = [normalize_cell(x) for x in row]

        if all(x == "" for x in row):
            continue

        rows.append(row)

        max_cols = max(max_cols, len(row))

    if len(rows) < CONFIG["min_rows"]:
        return []

    if max_cols < CONFIG["min_cols"]:
        return []

    if len(rows) * max_cols < CONFIG["min_cells"]:
        return []

    for row in rows:
        while len(row) < max_cols:
            row.append("")

    return rows


def _try(fn, *args, default=None, prefix=None, **kwargs):

    try:
        return fn(*args, **kwargs), None
    except Exception as e:
        msg = f"{prefix}: {e}" if prefix else str(e)
        return default, msg


@dataclass
class TableResult:
    bbox: Tuple[float, float, float, float]
    score: float
    data: List[List[str]]


@dataclass
class PageResult:
    page_index: int
    crop_box: Tuple[float, float, float, float]
    tables: List[TableResult] = field(default_factory=list)
    error: Optional[str] = None


@dataclass
class PdfResult:
    pdf_path: Path
    rows: List[tuple] = field(default_factory=list)
    error: Optional[str] = None


# ==============================================================================
# INPUT
# ==============================================================================


def get_pdf_files():

    mode = CONFIG["input_mode"]

    if mode == "single":
        return [Path(CONFIG["input_pdf"])]

    if mode == "folder":
        folder = Path(CONFIG["input_folder"])

        if not folder.exists() or not folder.is_dir():
            return []

        # Match .pdf and .PDF on every operating system.
        return sorted(p for p in folder.iterdir() if p.suffix.lower() == ".pdf")

    raise ValueError("Invalid input_mode")


# ==============================================================================
# SEARCH REGION
# ==============================================================================


def build_crop_box(width, height):

    region = CONFIG["search_region"]

    x0 = width * region["left"]

    x1 = width * region["right"]

    y0 = height * region["top"]

    y1 = height * region["bottom"]

    return (x0, y0, x1, y1)


def get_run_folder():

    now = datetime.now()

    root = Path(CONFIG.get("run_root", Path.cwd()))

    return root / f"Run_{now:%Y%m%d_%H%M%S}"


# ==============================================================================
# DEBUG IMAGE
# ==============================================================================


def draw_debug_image(doc, pdf_path, page_index, tables, run_folder):

    if doc is None:
        return

    debug_folder = run_folder / "debug"

    debug_folder.mkdir(parents=True, exist_ok=True)

    if page_index < 0 or page_index >= len(doc):
        print(f"Invalid page index {page_index} for debug image: {pdf_path}")
        return

    page = doc[page_index]

    zoom = 2.0

    matrix = fitz.Matrix(zoom, zoom)

    pix = page.get_pixmap(matrix=matrix)

    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    draw = ImageDraw.Draw(image)

    scale_x = image.width / page.rect.width
    scale_y = image.height / page.rect.height

    search_x0, search_y0, search_x1, search_y1 = build_crop_box(
        page.rect.width, page.rect.height
    )

    draw.rectangle(
        [
            search_x0 * scale_x,
            search_y0 * scale_y,
            search_x1 * scale_x,
            search_y1 * scale_y,
        ],
        outline="red",
        width=6,
    )

    for index, table in enumerate(tables):
        bbox = table.bbox

        x0, y0, x1, y1 = bbox

        draw.rectangle(
            [
                x0 * scale_x,
                y0 * scale_y,
                x1 * scale_x,
                y1 * scale_y,
            ],
            outline="blue",
            width=4,
        )

        draw.text((x0 * scale_x + 10, y0 * scale_y + 10), f"T{index + 1}", fill="blue")

    output = debug_folder / (f"{pdf_path.stem}_page_{page_index + 1}.png")

    image.save(output)

    log("Debug image:", output)


# ==============================================================================
# TABLE SCORE
# ==============================================================================
def compute_score(table_obj, table_data, page):

    bbox = table_obj.bbox

    x0, y0, x1, y1 = bbox

    width = x1 - x0

    height = y1 - y0

    page_width = page.width

    page_height = page.height

    rows = len(table_data)

    cols = max(len(r) for r in table_data)

    cells = rows * cols

    score = 0

    scoring = CONFIG["scoring"]

    # ------------------------------------------------------------------
    # Position score
    # ------------------------------------------------------------------

    left_ratio = x0 / page_width

    bottom_ratio = y1 / page_height

    score += (1 - left_ratio) * scoring["left_weight"]

    score += bottom_ratio * scoring["bottom_weight"]

    # ------------------------------------------------------------------
    # Aspect ratio
    # ------------------------------------------------------------------

    aspect = width / max(height, 1)

    score += min(aspect, 5) * scoring["aspect_weight"]

    # ------------------------------------------------------------------
    # Cell count
    # ------------------------------------------------------------------

    score += min(cells / 10, 10) * scoring["cell_weight"]

    # ------------------------------------------------------------------
    # Text analysis
    # ------------------------------------------------------------------

    text_count = 0

    for row in table_data:
        for cell in row:
            cell = str(cell).strip()

            if not cell:
                continue

            text_count += 1

    score += min(text_count / 10, 10) * scoring["text_weight"]

    # ------------------------------------------------------------------
    # Penalty: too tall
    # ------------------------------------------------------------------

    if height > page_height * scoring["height_threshold"]:
        score -= scoring["height_penalty"]

    # ------------------------------------------------------------------
    # Penalty: vertical drawing
    # ------------------------------------------------------------------

    if aspect < scoring["vertical_threshold"]:
        score -= scoring["vertical_penalty"]

    return score


# ==============================================================================
# TABLE EXTRACTION
# ==============================================================================


def extract_page_tables(page, page_index):

    crop_box = build_crop_box(page.width, page.height)

    cropped_page, err = _try(
        page.crop, crop_box, default=None, prefix="Failed to crop page"
    )
    if err:
        return PageResult(page_index=page_index, crop_box=crop_box, error=err)

    raw_tables, err = _try(
        cropped_page.find_tables,
        table_settings=CONFIG["table_settings"],
        default=[],
        prefix="Failed to find tables",
    )
    if err:
        return PageResult(page_index=page_index, crop_box=crop_box, error=err)

    tables = []

    min_score = CONFIG["scoring"]["min_score"]

    for table_obj in raw_tables:
        table_data, err = _try(
            table_obj.extract, default=None, prefix="Failed table extraction"
        )
        if err:
            continue

        table_data = clean_table(table_data)
        if not table_data:
            continue

        score = compute_score(table_obj, table_data, cropped_page)
        if score < min_score:
            continue

        tables.append(TableResult(bbox=table_obj.bbox, score=score, data=table_data))

    tables.sort(key=lambda t: t.score, reverse=True)

    return PageResult(page_index=page_index, crop_box=crop_box, tables=tables)


# ==============================================================================
# PDF PROCESS
# ==============================================================================


def process_pdf(pdf_path, run_folder):

    rows = []

    log()
    log("=" * 80)
    log("Processing:", pdf_path.name)

    # Open the PDF for debug images once, not once per page.
    debug_doc = None
    if CONFIG["export_debug_images"]:
        debug_doc, err = _try(
            fitz.open,
            pdf_path,
            default=None,
            prefix=f"Failed to open PDF for debug image: {pdf_path}",
        )
        if err:
            print(err)

    def _run():
        with pdfplumber.open(pdf_path) as pdf:
            for page_index, page in enumerate(pdf.pages):
                page_result = extract_page_tables(page, page_index)

                if page_result.error:
                    print(f"Page {page_index + 1}: {page_result.error}")
                else:
                    log(f"Page {page_index + 1}: {len(page_result.tables)} table(s)")

                # A failed debug image must not stop the table extraction.
                _, err = _try(
                    draw_debug_image,
                    debug_doc,
                    pdf_path,
                    page_index,
                    page_result.tables,
                    run_folder,
                    prefix=f"Page {page_index + 1}: debug image failed",
                )
                if err:
                    print(err)

                for table_index, table in enumerate(page_result.tables, start=1):
                    log(f"  Table {table_index}", f"Rows={len(table.data)}")
                    for row in table.data:
                        rows.append([pdf_path.name, page_index + 1, table_index, *row])

    try:
        _, err = _try(_run, prefix=f"Failed to process PDF {pdf_path.name}")
    finally:
        if debug_doc is not None:
            debug_doc.close()

    if err:
        return PdfResult(pdf_path=pdf_path, error=err)

    return PdfResult(pdf_path=pdf_path, rows=rows)


# ==============================================================================
# EXPORT
# ==============================================================================


def export_excel(rows, run_folder):

    if not rows:
        print("No table found.")

        return

    max_cols = max(len(r) for r in rows)

    # Tables can have different widths, so pad every row to the widest one.
    normalized = []

    for row in rows:
        row = list(row)

        while len(row) < max_cols:
            row.append("")

        normalized.append(row)

    headers = ["Source File", "Page", "Table"]

    extra = max_cols - len(headers)

    headers.extend([f"Col_{i}" for i in range(1, extra + 1)])

    df = pd.DataFrame(normalized, columns=headers)

    output = run_folder / Path(CONFIG["output_excel"]).name

    output.parent.mkdir(parents=True, exist_ok=True)

    df.to_excel(output, index=False, sheet_name=CONFIG["sheet_name"])

    print()

    print("Saved:", output)

    print("Rows:", len(df))


# ==============================================================================
# MAIN
# ==============================================================================


def main():

    pdf_files = get_pdf_files()

    if not pdf_files:
        print("No PDF found.")

        return

    run_folder = get_run_folder()

    run_folder.mkdir(parents=True, exist_ok=True)

    log("Run folder:", run_folder)

    all_rows = []

    for pdf_file in pdf_files:
        result = process_pdf(pdf_file, run_folder)
        if result.error:
            print(result.error)
        all_rows.extend(result.rows)

    export_excel(all_rows, run_folder)


if __name__ == "__main__":
    main()
