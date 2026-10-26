"""
Tests for Schedule.py (standard library only).

Run from the repo root:
    python -m unittest discover -s tests -t .
"""

import copy
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

import Schedule

SAMPLE_PDF = Path(__file__).resolve().parent.parent / "input" / "TRUSS LOT 73.pdf"


class ConfigTestCase(unittest.TestCase):
    """Gives every test its own copy of CONFIG and a temp folder."""

    def setUp(self):
        self._saved_config = copy.deepcopy(Schedule.CONFIG)
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        Schedule.CONFIG.clear()
        Schedule.CONFIG.update(self._saved_config)
        self._tmp.cleanup()


class NormalizeCellTests(unittest.TestCase):
    def test_none_becomes_empty(self):
        self.assertEqual(Schedule.normalize_cell(None), "")

    def test_newlines_and_extra_spaces_collapse(self):
        self.assertEqual(Schedule.normalize_cell("  UC250*90-S \n  300 "), "UC250*90-S 300")

    def test_numbers_become_text(self):
        self.assertEqual(Schedule.normalize_cell(3235), "3235")


class CleanTableTests(unittest.TestCase):
    def test_blank_rows_are_dropped(self):
        table = [["a", "b"], ["", None], ["c", "d"], ["e", "f"], ["g", "h"]]
        self.assertEqual(
            Schedule.clean_table(table),
            [["a", "b"], ["c", "d"], ["e", "f"], ["g", "h"]],
        )

    def test_too_few_rows_rejected(self):
        self.assertEqual(Schedule.clean_table([["a", "b"], ["c", "d"]]), [])

    def test_too_few_columns_rejected(self):
        self.assertEqual(Schedule.clean_table([["a"], ["b"], ["c"], ["d"]]), [])

    def test_too_few_cells_rejected(self):
        # 3 rows x 2 cols = 6 cells, below the minimum of 8
        self.assertEqual(Schedule.clean_table([["a", "b"], ["c", "d"], ["e", "f"]]), [])

    def test_short_rows_are_padded(self):
        table = [["a", "b", "c"], ["d"], ["e", "f"], ["g", "h", "i"]]
        result = Schedule.clean_table(table)
        self.assertEqual(result[1], ["d", "", ""])
        self.assertEqual(result[2], ["e", "f", ""])


class BuildCropBoxTests(unittest.TestCase):
    def test_default_region_is_bottom_left(self):
        page = SimpleNamespace(width=1000, height=500)
        self.assertEqual(Schedule.build_crop_box(page), (0, 300, 400, 500))


class ComputeScoreTests(unittest.TestCase):
    """Pins the current scoring numbers so later changes are deliberate."""

    def test_wide_table_near_bottom_left(self):
        page = SimpleNamespace(width=1000, height=500)
        table_obj = SimpleNamespace(bbox=(0, 400, 300, 450))
        data = [["x"] * 3 for _ in range(4)]
        # left 40 + bottom 36 + aspect 150 + cells 24 + text 24
        self.assertAlmostEqual(Schedule.compute_score(table_obj, data, page), 274)

    def test_tall_narrow_table_is_penalised(self):
        page = SimpleNamespace(width=1000, height=500)
        table_obj = SimpleNamespace(bbox=(0, 0, 100, 300))
        data = [["x"] * 3 for _ in range(3)]
        # 110 base - 50 (too tall) - 100 (vertical)
        self.assertAlmostEqual(Schedule.compute_score(table_obj, data, page), -40)


class GetPdfFilesTests(ConfigTestCase):
    def test_single_mode(self):
        Schedule.CONFIG["input_mode"] = "single"
        Schedule.CONFIG["input_pdf"] = str(self.tmp / "a.pdf")
        self.assertEqual(Schedule.get_pdf_files(), [self.tmp / "a.pdf"])

    def test_folder_mode_returns_sorted_pdfs_only(self):
        for name in ["b.pdf", "a.pdf", "notes.txt"]:
            (self.tmp / name).write_text("x")
        Schedule.CONFIG["input_mode"] = "folder"
        Schedule.CONFIG["input_folder"] = str(self.tmp)
        self.assertEqual(
            Schedule.get_pdf_files(), [self.tmp / "a.pdf", self.tmp / "b.pdf"]
        )

    def test_missing_folder_returns_empty(self):
        Schedule.CONFIG["input_mode"] = "folder"
        Schedule.CONFIG["input_folder"] = str(self.tmp / "nope")
        self.assertEqual(Schedule.get_pdf_files(), [])

    def test_invalid_mode_raises(self):
        Schedule.CONFIG["input_mode"] = "banana"
        with self.assertRaises(ValueError):
            Schedule.get_pdf_files()


class ExportExcelTests(ConfigTestCase):
    def test_no_rows_writes_nothing(self):
        Schedule.CONFIG["output_excel"] = "out.xlsx"
        Schedule.export_excel([], self.tmp)
        self.assertFalse((self.tmp / "out.xlsx").exists())

    def test_ragged_rows_are_padded_and_headers_named(self):
        Schedule.CONFIG["output_excel"] = "out.xlsx"
        rows = [["a.pdf", 1, 1, "x", "y"], ["a.pdf", 1, 1, "z"]]
        Schedule.export_excel(rows, self.tmp)
        df = pd.read_excel(self.tmp / "out.xlsx")
        self.assertEqual(
            list(df.columns), ["Source File", "Page", "Table", "Col_1", "Col_2"]
        )
        self.assertEqual(len(df), 2)
        self.assertEqual(df.loc[1, "Col_1"], "z")


@unittest.skipUnless(SAMPLE_PDF.exists(), "sample PDF not in input/ (kept out of git)")
class SamplePdfSmokeTest(ConfigTestCase):
    def test_truss_lot_73(self):
        Schedule.CONFIG["debug"] = False
        result = Schedule.process_pdf(SAMPLE_PDF, self.tmp)

        self.assertIsNone(result.error)
        self.assertEqual(len(result.rows), 115)

        tables = {(row[1], row[2]) for row in result.rows}
        self.assertEqual(
            tables,
            {(1, 1), (2, 1), (3, 1), (4, 1), (5, 1), (6, 1), (6, 2), (7, 1), (8, 1), (9, 1), (9, 2)},
        )

        flat = [cell for row in result.rows for cell in row]
        self.assertIn("MATERIAL LIST", flat)
        self.assertIn("fc380", flat)

        self.assertEqual(len(list((self.tmp / "debug").glob("*.png"))), 9)


if __name__ == "__main__":
    unittest.main()
