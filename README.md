# PDF Steel Schedule Extractor

Extracts steel schedule (material list) tables from PDF drawings and exports them to Excel.

## Features

- Scans only the bottom-left area of each drawing page
- Exports ALL tables detected in that area to one Excel file
- Debug mode draws the search area and detected tables onto PNG images
- Easy to tune through the `CONFIG` block at the top of `Schedule.py`
- Works on a single PDF or on a whole folder of PDFs

## Requirements

- Python 3.11+ (required by pandas 3.x)
- Python libraries (see `requirements.txt`):
  - pdfplumber
  - pandas
  - openpyxl (writes the Excel file)
  - PyMuPDF
  - Pillow

## Install

```bash
pip install -r requirements.txt
```

## Folder structure

You set one mother folder (`base_folder` in `CONFIG`). Everything else lives inside it.

```text
Steel-Schedule-Finder/            <- base_folder
├── Schedule.py                   # The program
├── requirements.txt              # Python libraries
├── README.md                     # This file
├── context/                      # Plans and notes (contents not pushed to git)
├── tests/                        # Automatic tests (unittest)
├── input/                        # Put your PDFs here (contents not pushed to git)
└── output/                       # Results (contents not pushed to git)
    └── Run_YYYYMMDD_HHMMSS/      # One folder per run, never overwritten
        ├── output.xlsx           # Excel result
        └── debug/                # Debug PNG images
```

## How to use

1. Put your PDF files in the `input/` folder.
2. Open `Schedule.py` and check the `CONFIG` block (see below). Most of the time you only need `base_folder` and `input_mode`.
3. Run:

   ```bash
   python Schedule.py
   ```

4. Open the newest folder in `output/`. The Excel file is `output.xlsx`; the debug images are in `debug/`.

### Folder settings

```python
CONFIG = {
    # Mother folder. input/ and output/ are created inside it.
    "base_folder": r"D:\Working\GITHUB REPO\Steel-Schedule-Finder",
    "input_dir": "input",
    "output_dir": "output",

    # "folder" = every PDF in input/, "single" = only the file named in input_pdf
    "input_mode": "folder",
    # Used only in "single" mode: a file name inside input/
    "input_pdf": "TRUSS LOT 73.pdf",

    # File name of the Excel result (saved inside each Run_* folder)
    "output_excel": "output.xlsx",

    # Debug
    "debug": True,                # Print progress messages to the console
    "export_debug_images": True,  # Save PNG images to <run folder>/debug/
    ...
}
```

To move the project to another computer or folder, change `base_folder` only.

### Search region

The region is given as ratios (0.00 to 1.00) of the page width and height:

```python
"search_region": {
    "left": 0.00,    # start at the left edge
    "right": 0.40,   # up to 40% of the page width
    "top": 0.60,     # start at 60% of the page height
    "bottom": 1.00,  # down to the bottom of the page
}
```

This means: the left 40% of the page width, and the lower 40% of the page height.

### Other settings

```python
"table_settings": {            # passed to pdfplumber's find_tables()
    "snap_tolerance": 5,
    "join_tolerance": 5,
    "intersection_tolerance": 8,
    "edge_min_length": 10,
},
"min_rows": 3,                 # minimum non-empty rows in a table
"min_cols": 2,                 # minimum columns
"min_cells": 8,                # minimum rows x columns
"sheet_name": "Steel Schedule",
"scoring": {                   # how tables are ranked (see Tuning)
    "left_weight": 40,
    "bottom_weight": 40,
    "aspect_weight": 30,
    "cell_weight": 20,
    "text_weight": 20,
    "height_penalty": 50,
    "min_score": 300,
    "height_threshold": 0.20,
    "vertical_threshold": 1.2,
    "vertical_penalty": 100,
},
```

## Debug mode

With `export_debug_images` on, each page gets a PNG in `<run folder>/debug/`:

- **Red box**: the search area (`search_region`)
- **Blue boxes**: the tables that were found
- **T1, T2, ...**: the table numbers

With `debug` on, the program also prints progress to the console. Errors are always printed, even when `debug` is off.

Use the debug images whenever you change the settings, to check what was picked up.

## Tuning

### Tables are missed or detected wrongly

Adjust `table_settings`:

- `snap_tolerance`: raise it if table lines do not join up
- `join_tolerance`: raise it if table lines are split into pieces
- `intersection_tolerance`: raise it if line crossings are not detected
- `edge_min_length`: raise it if many tiny lines are picked up

### Too many small or junk tables

- Raise `min_rows`, `min_cols` or `min_cells` (for example 4-5 rows, 3-4 columns, 10-15 cells)
- Or raise `scoring.min_score` (the default is 300; see "How tables are scored")
- Or raise `table_settings.edge_min_length`

### How tables are scored

Every table found in the search region gets a score, and only tables with a score of at least `scoring.min_score` are exported. Points are added for:

- being close to the left edge and to the bottom of the **full page**
- being wide (width / height, capped at 5)
- having many cells and a lot of text (each capped)

Points are taken away for tables taller than `height_threshold` of the page height, and for tall narrow tables (`vertical_threshold`).

The default `min_score` of 300 was tuned on the sample drawing (`TRUSS LOT 73.pdf`):

| What | Score |
|---|---|
| Real material lists | 400 to 600 |
| Junk tables made of dimension numbers | about 250 to 270 |
| Other small fragments | below 70 |

A small, fully filled list of the same width scores about 310 (3 rows) to 360 (6 rows), so it still passes. If a **real table is missing**, lower `min_score` (for example 250). If **junk tables appear**, raise it (for example 350). Check the debug images after each change. These numbers come from one drawing, so re-check them when you use drawings from a different drafter.

### The steel schedule is not ranked correctly

Adjust the weights in `scoring`:

- `left_weight`: raise it to favour tables close to the left edge
- `bottom_weight`: raise it to favour tables close to the bottom
- `aspect_weight`: raise it to favour wide tables
- `cell_weight`: raise it to favour tables with many cells
- `text_weight`: raise it to favour tables with a lot of text

## Troubleshooting

### "No PDF found."

Check that your PDFs are in `<base_folder>/input/` and that `base_folder` is correct. In `"single"` mode, check that `input_pdf` matches the file name.

### "No table found."

1. Look at the debug images to see the search area and what was detected
2. Widen `search_region`
3. Lower `min_rows`, `min_cols`, `min_cells` or `scoring.min_score`
4. Adjust `table_settings`

### Tables are cut off or data is missing

1. Widen `search_region`
2. Lower `min_rows`, `min_cols`, `min_cells`
3. Lower `scoring.min_score`

## Running the tests

```bash
python -m unittest discover -s tests -t .
```

The smoke tests use `input/TRUSS LOT 73.pdf`. That file is not in git; if it is missing, those tests are skipped.

## Notes and known limits

- The folders `input/`, `output/` and `context/` exist in git (each holds an empty `.gitkeep` file), but nothing else inside them is pushed. `Run_*/` folders and all `*.pdf` / `*.xlsx` files are never pushed either (see `.gitignore`).
- Every run creates a new `Run_YYYYMMDD_HHMMSS` folder, so old results are kept.
- If a material list has two blocks side by side, each Excel row holds two items (columns `Col_1`-`Col_7` and `Col_8`-`Col_14`). Splitting them is not done yet.
- Rotated PDF pages can make the debug boxes appear in the wrong place. Check the Excel data against the drawing.
