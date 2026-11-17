# Python Excel Data Cleaner

A Python automation tool that helps clean messy Excel spreadsheets by removing duplicate rows, standardizing text, and identifying missing values.

Built with **Python** and **openpyxl**, this project demonstrates practical Excel data cleaning and automation skills.

## Features

- **Duplicate Removal:** Removes duplicate data rows across worksheets.
- **Text Cleaning:** Removes unnecessary spaces and standardizes whitespace.
- **Missing Value Detection:** Counts empty cells and generates per-column statistics.
- **Multiple Worksheets:** Processes every worksheet in an Excel workbook.
- **JSON Reports:** Generates a report summarizing the cleaning results.
- **Safe Output:** Prevents accidental overwriting of existing files.

## Technologies Used

- Python 3.10+
- openpyxl
- Git & GitHub

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/javohirtoych iyev220-glitch/ython-excel-data-cleane.git
cd ython-excel-data-cleane
python -m pip install -r requirements.txt
```

## Usage

Run the cleaner with a sample Excel file:

```bash
python excel_cleaner.py sample_data.xlsx
```

The program generates a cleaned Excel workbook and a JSON report.

You can also specify output paths:

```bash
python excel_cleaner.py input.xlsx --output output.xlsx --report report.json
```

## Testing

The project includes automated tests.

```bash
python -m pytest -v
```

Latest local test result: **4 passed (0.52 seconds)**.

## Example Use Cases

- Cleaning employee spreadsheets
- Preparing customer records
- Removing duplicate entries
- Standardizing Excel data before analysis
- Identifying missing information in datasets

## Future Improvements

- CSV file support
- Configurable cleaning rules
- Batch processing of multiple files
- A simple graphical interface

## Author

Javohir — Python Automation & Excel Data Cleaning Portfolio Project