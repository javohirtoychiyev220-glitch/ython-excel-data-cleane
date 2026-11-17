"""Clean Excel workbooks and write a JSON summary report."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook


def _normalize(value: Any) -> Any:
    if isinstance(value, str):
        return " ".join(value.split())
    return value


def _is_missing(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _check_paths(input_path: Path, output_path: Path, report_path: Path) -> None:
    resolved_input = input_path.resolve()
    resolved_output = output_path.resolve()
    resolved_report = report_path.resolve()

    if resolved_input in (resolved_output, resolved_report):
        raise ValueError("Output and report paths must not overwrite the input workbook.")
    if resolved_output == resolved_report:
        raise ValueError("The output workbook and JSON report must use different paths.")
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("The output workbook must use the .xlsx extension.")
    if report_path.suffix.lower() != ".json":
        raise ValueError("The summary report must use the .json extension.")
    if output_path.exists():
        raise FileExistsError(f"Output workbook already exists: {output_path}")
    if report_path.exists():
        raise FileExistsError(f"Summary report already exists: {report_path}")
    if not output_path.parent.is_dir():
        raise FileNotFoundError(f"Output directory does not exist: {output_path.parent}")
    if not report_path.parent.is_dir():
        raise FileNotFoundError(f"Report directory does not exist: {report_path.parent}")


def _clean_sheet(worksheet: Any) -> dict[str, Any]:
    whitespace_cells_cleaned = 0
    data_rows: list[tuple[int, tuple[Any, ...]]] = []

    for row_index in range(1, worksheet.max_row + 1):
        original_values = tuple(
            worksheet.cell(row=row_index, column=column_index).value
            for column_index in range(1, worksheet.max_column + 1)
        )
        values = tuple(_normalize(value) for value in original_values)

        for column_index, (original, normalized) in enumerate(
            zip(original_values, values), start=1
        ):
            if original != normalized:
                worksheet.cell(row=row_index, column=column_index).value = normalized
                whitespace_cells_cleaned += 1

        if row_index > 1 and any(not _is_missing(value) for value in values):
            data_rows.append((row_index, values))

    missing_before = sum(
        _is_missing(value) for _, values in data_rows for value in values
    )
    seen_rows: set[tuple[Any, ...]] = set()
    duplicate_indices: list[int] = []
    kept_rows: list[tuple[Any, ...]] = []
    for row_index, values in data_rows:
        if values in seen_rows:
            duplicate_indices.append(row_index)
        else:
            seen_rows.add(values)
            kept_rows.append(values)

    for row_index in reversed(duplicate_indices):
        worksheet.delete_rows(row_index)

    headers = [
        worksheet.cell(row=1, column=column_index).value
        for column_index in range(1, worksheet.max_column + 1)
    ]
    missing_by_column = []
    for column_index, header in enumerate(headers):
        missing_by_column.append(
            {
                "column": get_column_letter(column_index + 1),
                "header": header,
                "missing_values": sum(
                    _is_missing(row[column_index]) for row in kept_rows
                ),
            }
        )

    missing_after = sum(
        _is_missing(value) for values in kept_rows for value in values
    )
    return {
        "data_rows_before": len(data_rows),
        "data_rows_after": len(kept_rows),
        "duplicate_rows_removed": len(duplicate_indices),
        "whitespace_cells_cleaned": whitespace_cells_cleaned,
        "missing_values_before": missing_before,
        "missing_values_after": missing_after,
        "missing_by_column": missing_by_column,
    }


def clean_workbook(
    input_path: str | Path,
    output_path: str | Path | None = None,
    report_path: str | Path | None = None,
) -> dict[str, Any]:
    """Clean every worksheet and write a workbook plus a JSON summary.

    The first row of each worksheet is treated as its header. Entirely blank
    data rows are ignored, and duplicate records are removed after whitespace
    normalization. Existing output files are never overwritten.
    """
    source = Path(input_path)
    if not source.is_file():
        raise FileNotFoundError(f"Input workbook does not exist: {source}")
    if source.suffix.lower() != ".xlsx":
        raise ValueError("The input workbook must use the .xlsx extension.")

    destination = (
        Path(output_path)
        if output_path is not None
        else source.with_name(f"cleaned_{source.name}")
    )
    report = (
        Path(report_path)
        if report_path is not None
        else destination.with_name(f"{destination.stem}_report.json")
    )
    _check_paths(source, destination, report)

    workbook: Workbook = load_workbook(source)
    try:
        sheet_reports = []
        for worksheet in workbook.worksheets:
            sheet_reports.append(
                {"sheet": worksheet.title, **_clean_sheet(worksheet)}
            )
        workbook.save(destination)
    finally:
        workbook.close()

    totals = {
        key: sum(sheet[key] for sheet in sheet_reports)
        for key in (
            "data_rows_before",
            "data_rows_after",
            "duplicate_rows_removed",
            "whitespace_cells_cleaned",
            "missing_values_before",
            "missing_values_after",
        )
    }
    summary = {
        "input_file": str(source.resolve()),
        "output_file": str(destination.resolve()),
        "report_file": str(report.resolve()),
        "sheet_count": len(sheet_reports),
        "totals": totals,
        "sheets": sheet_reports,
    }
    report.write_text(
        json.dumps(summary, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Remove duplicate Excel rows, normalize spaces, and report missing values."
    )
    parser.add_argument("input", type=Path, help="Input .xlsx workbook")
    parser.add_argument(
        "--output", type=Path, help="Output workbook (default: cleaned_<input name>)"
    )
    parser.add_argument(
        "--report", type=Path, help="JSON summary report (default: next to output)"
    )
    arguments = parser.parse_args(argv)

    try:
        summary = clean_workbook(arguments.input, arguments.output, arguments.report)
    except (OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2

    totals = summary["totals"]
    print(f"Cleaned workbook: {summary['output_file']}")
    print(f"Summary report: {summary['report_file']}")
    print(
        "Rows removed: "
        f"{totals['duplicate_rows_removed']}; "
        f"whitespace cells cleaned: {totals['whitespace_cells_cleaned']}; "
        f"missing values after cleaning: {totals['missing_values_after']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
