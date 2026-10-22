# Insect label maker

Generate print-ready insect specimen labels from an Excel workbook. The program creates one label per specimen row, with locality and collection information, a vertically oriented catalog number, a Data Matrix barcode encoding that catalog number, and a collection-event code.

The tool runs locally, reads the workbook without changing it, and exports a US Letter PDF plus a JSON layout report. Text and barcodes are vector graphics, and fonts are embedded for sharp printing.

## Files to share

- `label_maker.py` — main Python script.
- `requirements.txt` — Python dependencies.
- `label_template.json` — editable field order and punctuation.
- `examples/` — fictional example data, PDFs for both layouts and separate habitat labels, and their reports.
- `README.md` — these instructions.

Send these files together. Do not copy your `.venv` folder; each user creates their own environment. Arial font files are not included. A ready-to-share copy of these files is available in `dist/insect_label_maker.zip`.

## Installation

Use Python 3.11 or newer. Extract the files into a folder, open a terminal there, and create a virtual environment.

**macOS / Linux:**

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

**Windows PowerShell:**

```powershell
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The commands below use `python`. On Windows, substitute `.venv\Scripts\python.exe`; on macOS/Linux you can use `.venv/bin/python` without activating the environment. Activate the environment again when opening a new terminal if using the short `python` command.

### Fonts

The program requires matching **Arial Regular** and **Arial Italic** TrueType font files. It looks in `/System/Library/Fonts/Supplemental/` by default on macOS. On other systems, or if fonts are elsewhere, supply both paths:

```sh
python label_maker.py examples/example_specimens.xlsx --font "/path/to/Arial.ttf" --italic-font "/path/to/Arial Italic.ttf" --output output/example.pdf
```

Typical Windows paths are `C:\Windows\Fonts\arial.ttf` and `C:\Windows\Fonts\ariali.ttf`. Linux users must provide their own installed Arial files. The program checks that the regular and italic fonts have matching character widths.

## Prepare the Excel file

Use an `.xlsx` workbook with column headers in the first row and one specimen per subsequent row. Column order does not matter; header spelling and capitalization do. Extra columns are ignored unless referenced in the template. The active worksheet is used unless `--sheet` selects another.

The supplied template uses these columns:

| Header | Purpose / example |
| --- | --- |
| `Catalog_Number` | Required, unique, nonempty specimen identifier, e.g. `DEMO000001`. Printed vertically and encoded in the barcode. |
| `Country` | Country name or abbreviation, e.g. `USA`. |
| `StateProv` or `StateProvince` | State/province abbreviation, e.g. `UT`. |
| `County` | County display text, e.g. `Cache Co.`. |
| `LocalityLabel` | Concise locality text, e.g. `Example Meadow`. |
| `Latitude` | Decimal latitude, e.g. `41.9`. |
| `Longitude` | Decimal longitude, e.g. `-111.65`. |
| `Uncertainty` | Numeric uncertainty in meters, e.g. `100`. May be blank; the template adds `±` and `m`. |
| `Verbatim_Elevation` | Elevation display text including units, e.g. `2500m`. The alternate spelling `Verbatium_Elevation` is also accepted. |
| `Verbatim_Date` | Collection date exactly as it should appear, e.g. `12Jul2026`. |
| `Collector_Label` | Concise collector names, e.g. `A. Smith, B. Jones`. Spaces after initials are removed for display. |
| `MacroHabitat_Label` | Broad habitat, e.g. `meadow`. May be blank. |
| `Microhabitat_Label` | Collection substrate, plant, or method, e.g. `Phacelia sp.` or `pan traps`. May be blank. |
| `Field_Number` | Optional collection-event identifier, e.g. `EVT001`, printed near the barcode. This column may be omitted. |

All template-referenced headers must exist, even when their cells are blank. `Catalog_Number` is always required, including with a custom template. Store identifiers as Excel **text** if leading zeros matter. Completely empty rows are skipped; other rows with a missing catalog number or duplicate catalog numbers are rejected. No extra copies are generated from specimen-count columns.

Keep the full scientific data in other columns and use the label-specific fields for concise wording. Text is not silently truncated. Formulas use their last saved calculated values: recalculate and save in Excel before importing. Text dates are preserved; actual Excel dates are formatted as `12/Jul/2026`. Coordinates retain up to 12 significant digits.

## Generate labels

Normal layout:

```sh
python label_maker.py "/path/to/specimens.xlsx" --template label_template.json --output output/labels_lines.pdf
```

Wrapped layout:

```sh
python label_maker.py "/path/to/specimens.xlsx" --template label_template.json --layout flow --output output/labels_flow.pdf
```

Quote paths containing spaces. Use a different output filename for each batch: an existing PDF and JSON report at that path will be replaced. If a specimen cannot fit, generation stops and identifies its Excel row and catalog number; a partial batch is not exported.

### Layouts and fitting

Both layouts default to **0.68 × 0.33-inch labels**, with 0.02-inch horizontal/vertical gaps and 0.25-inch page margins. Labels fill down a column, then across the page, with automatic pagination. **Width must never exceed 0.68 inch**: labels are inserted vertically, and this limit allows four rows in a standard unit tray. Automatic widening is disabled. Height may be increased explicitly with `--height` (default `0.33` inch).

- **`lines` (default):** keeps the six template entries on separate lines. Entries 1–5 start at 4 points and individually fall back to 3.5 points if needed. Entry 6 stays at 3.5 points. The adaptive layout uses equal 3.8-point baseline spacing, independent of font size.
- **`flow`:** joins entries with commas and wraps them into a text column. It tries 4-point text first and 3.5 points when needed; physical lines containing habitat text use 3.5 points. Wrapped lines use their own flow spacing.

All main and habitat text stays behind an invisible boundary left of the barcode, including above and below it. Catalog and field numbers have separate positions. The Data Matrix is vertically centered; catalog text is 3.75 points and field-number text is 3.1 points.

Text normally uses 75% horizontal scale. In the standard normal layout, latitude/longitude beside uncertainty uses 62% scale, and the uncertainty term is always 3.5 points. The coordinate line stays on line 3; elevation/date stay on line 4.

The program first tries to retain all information, then may omit coordinate uncertainty. For the standard habitat template, if the label still cannot fit, it tries microhabitat alone. It never automatically removes microhabitat. The JSON report records omitted uncertainty/macrohabitat, rendered lines, font sizes, positions, and text runs. Review that report when completeness matters. These fallbacks can differ between the two layouts.

Within `Microhabitat_Label`, recognizable genus/species combinations are italicized: *Phacelia* sp. or *Phacelia tanacetifolia*. `sp.`, `spp.`, family names such as Asteraceae, and descriptive text remain upright. Recognition uses name structure, not a taxonomic database; review unusual or abbreviated names.

### Command-line options

| Option | Default | Effect |
| --- | --- | --- |
| `excel` | Required | Input `.xlsx` path (positional argument). |
| `-o`, `--output` | `output/labels.pdf` | PDF path; JSON report uses the same basename. |
| `--layout lines\|flow` | `lines` | Separate-line or comma-separated wrapped layout. |
| `--main-size 4\|3.5` | `4` | Preferred main-text size. `4` permits per-line fallback; `3.5` uses the smaller normal-layout profile. |
| `--template PATH` | Built-in six entries | Load a JSON template. |
| `--sheet NAME` | Active sheet | Select a worksheet. |
| `--font PATH` | macOS Arial path | Regular Arial `.ttf` file. |
| `--italic-font PATH` | macOS Arial Italic path | Matching italic `.ttf` file. |
| `--guides` | Off | Draw light cutting rectangles. |
| `--separate-habitat` | Off | Put macrohabitat and microhabitat on a smaller label below the main label. |
| `--habitat-height INCHES` | `0.20` | Height of the separate habitat label. |
| `--habitat-width INCHES` | `0.62` | Width of the smaller habitat label; must be less than the main width. |
| `--habitat-max-lines 2\|3` | `3` | Maximum physical lines on the habitat label. |
| `--habitat-gap INCHES` | `0.02` | Gap between the main and separate habitat labels. |
| `--width INCHES` | `0.68` | Label width; must be positive and never exceed 0.68 inch. |
| `--height INCHES` | `0.33` | Label height; may be increased up to 10.5 inches within the printable page. |
| `--max-width INCHES` | Not set | Deprecated compatibility option; must equal `--width`. It cannot permit widening. |
| `-h`, `--help` | — | Show usage. |

Keep the default 0.68-inch width for unit-tray compatibility. To try more vertical space without increasing width:

```sh
python label_maker.py "/path/to/specimens.xlsx" --layout flow --height 0.40 --output output/taller_labels.pdf
```

The barcode and catalog number remain vertically centered, and page row counts adapt to the chosen height. A taller label can fit more wrapped lines, but it does not solve an individual text line or word that is too wide; shorten that text or use the wrapped layout. Barcode size, code font sizes, and page gaps are constants in `label_maker.py`, not command-line options. Longer catalog identifiers may require more height or a larger barcode to scan reliably.

## Optional separate habitat label

Add `--separate-habitat` to move habitat out of the main label. Each specimen gets a smaller label directly below its main label, with a 0.02-inch cutting gap. The two labels always stay together on the same page. This works with both `--layout lines` and `--layout flow`.

```sh
python label_maker.py "/path/to/specimens.xlsx" --template label_template.json --separate-habitat --output output/label_pairs.pdf
```

The habitat label defaults to **0.62 × 0.20 inches**, narrower than the main label. It contains `MacroHabitat_Label`, then `Microhabitat_Label`, in separate left-aligned blocks. Either field may wrap, up to **three physical lines total**. The full text block is centered vertically and always uses **4-point Arial**, at the usual 75% horizontal scale. Scientific-name italics remain on the microhabitat line. There is no barcode on the smaller label; it uses the full available width.

Both habitat columns are required for this option, but values may be blank. A pair is reserved for every specimen. The program preserves both habitat values instead of applying the macrohabitat-omission fallback. The program wraps at word boundaries without shrinking or truncating text. It reports an error if a word is too wide, the text needs more than the selected maximum of two or three lines, or the chosen height cannot accommodate those lines.

Use `--habitat-height 0.22` for a taller habitat label or `--habitat-gap 0.03` for more space below the main label. `--habitat-width 0.60` narrows the label further. `--habitat-max-lines 2 --habitat-height 0.15` restores a compact two-line format when the text fits. Both options are in inches. `--guides` draws cutting rectangles around both labels. Main-label dimensions remain unchanged. The JSON report includes each habitat label’s dimensions, text, baselines, and page/position. If using a custom template, keep habitat fields in entries containing only habitat fields so they can be removed without losing other data.

See `examples/example_labels_separate_habitat.pdf` for a fictional example.

## Change displayed fields

Edit `label_template.json`; each string contains literal punctuation and `[Column_Header]` placeholders:

```json
[
  "[Country]: [StateProv], [County]",
  "[LocalityLabel]",
  "[Latitude] [Longitude]±[Uncertainty]m",
  "[Verbatim_Elevation], [Verbatim_Date]",
  "[Collector_Label]",
  "[MacroHabitat_Label]; [Microhabitat_Label]"
]
```

Use concise fields and retain the six-entry structure for the normal layout. The first five entries are main text; the sixth is habitat. Wrapped layout treats the sixth and later entries as habitat text. Unknown headers cause an error. The automatic microhabitat-only fallback recognizes the exact standard `[MacroHabitat_Label]; [Microhabitat_Label]` entry; changing that entry disables this fallback. Coordinate-specific mixed formatting recognizes the standard combined latitude/longitude/uncertainty text. Scientific-name emphasis only applies to text from `Microhabitat_Label`.

## Try the examples

The example workbook contains **six fictional specimens**, including plant names, a family name, blank uncertainty, and a long habitat description. The matching PDFs demonstrate both layouts, the habitat fallback, and the separate habitat label option. They contain no production specimen records.

```sh
python label_maker.py examples/example_specimens.xlsx --template label_template.json --layout lines --output output/example_lines.pdf
python label_maker.py examples/example_specimens.xlsx --template label_template.json --layout flow --output output/example_flow.pdf
```

Compare your output with `examples/example_labels_lines.pdf` and `examples/example_labels_flow.pdf`. Example JSON reports accompany those PDFs.

## Print and troubleshoot

Print at **100% / Actual Size**, with “Fit to page” disabled. Use normal/high quality, preferably 600 dpi or higher, with draft/toner-saving mode off. Measure a label and scan several printed barcodes on the intended paper before printing a large batch.

- **Cannot fit:** inspect the reported row, shorten label-specific wording, remove unnecessary plant-family names, or try `--layout flow`. Width is capped at 0.68 inch for curation. If appropriate, try a taller label with `--height 0.40`.
- **Unknown template field:** check capitalization and header spelling; use one of the documented aliases where appropriate.
- **Missing font:** provide the regular and italic Arial paths explicitly.
- **Dates or formulas look wrong:** check cell types and save recalculated formula values in Excel.
- **Leading zeros disappeared:** store identifiers as text in the workbook.

Each barcode is digitally decoded before export. This verifies its encoded content; physical print-and-scan testing checks the printer, paper, and scanner together.

## Double-sided printing

The tool currently exports single-sided PDFs. See [the research note](docs/double_sided_labels.md) for published guidance, practitioner discussion, and considerations for a future duplicate-on-back experiment.
