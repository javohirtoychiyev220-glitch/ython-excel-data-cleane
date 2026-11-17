# Excel Data Cleaner

A small Python command-line application that cleans `.xlsx` workbooks with
[`openpyxl`](https://openpyxl.readthedocs.io/).

## Features

- Removes duplicate data rows from every worksheet. The first row is treated
  as the header, and duplicates are compared after whitespace normalization.
- Trims leading/trailing whitespace and reduces runs of whitespace within
  text cells to a single space.
- Counts missing values (`None` and blank/whitespace-only cells) in non-empty
  data rows, both before and after duplicate removal, with per-column totals.
- Writes a cleaned workbook and a JSON summary report.
- Refuses to overwrite existing files.

Entirely blank data rows are ignored for duplicate and missing-value counts.
Headers are whitespace-normalized but are not included in missing-value counts.

## Setup

Python 3.10 or newer is recommended.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Run

From this directory:

```powershell
python excel_cleaner.py sample_data.xlsx
```

By default this creates `cleaned_sample_data.xlsx` and
`cleaned_sample_data_report.json` alongside the input. Specify output paths
when needed:

```powershell
python excel_cleaner.py input.xlsx --output output.xlsx --report report.json
```

The application returns a non-zero exit code and prints an error if the input
is missing, the paths are invalid, or either output file already exists.

## Tests

```powershell
python -m unittest -v
```
