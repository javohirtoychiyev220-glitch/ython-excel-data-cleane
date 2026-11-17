import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_cleaner import clean_workbook


class CleanWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.directory = Path(self.temp_dir.name)
        self.source = self.directory / "input.xlsx"
        self.output = self.directory / "cleaned.xlsx"
        self.report = self.directory / "summary.json"

        workbook = Workbook()
        sheet = workbook.active
        sheet.append([" Name ", "Department", "Notes"])
        sheet.append(["  Alice   Smith ", " Engineering ", None])
        sheet.append(["Alice Smith", "Engineering", None])
        sheet.append(["Bob", None, "  needs   review  "])
        sheet.append(["   ", None, None])
        workbook.save(self.source)
        workbook.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_normalizes_spaces_removes_duplicates_and_reports_missing(self):
        summary = clean_workbook(self.source, self.output, self.report)

        self.assertEqual(summary["totals"]["duplicate_rows_removed"], 1)
        self.assertEqual(summary["totals"]["whitespace_cells_cleaned"], 5)
        self.assertEqual(summary["totals"]["missing_values_before"], 3)
        self.assertEqual(summary["totals"]["missing_values_after"], 2)

        workbook = load_workbook(self.output, data_only=True)
        sheet = workbook.active
        self.assertEqual(sheet.max_row, 4)
        self.assertEqual(
            [cell.value for cell in sheet[1]], ["Name", "Department", "Notes"]
        )
        self.assertEqual(
            [cell.value for cell in sheet[2]], ["Alice Smith", "Engineering", None]
        )
        self.assertEqual(
            [cell.value for cell in sheet[3]], ["Bob", None, "needs review"]
        )
        self.assertTrue(all(cell.value is None for cell in sheet[4]))
        workbook.close()

        saved_report = json.loads(self.report.read_text(encoding="utf-8"))
        self.assertEqual(saved_report["totals"], summary["totals"])
        self.assertEqual(
            [column["missing_values"] for column in saved_report["sheets"][0]["missing_by_column"]],
            [0, 1, 1],
        )

    def test_does_not_overwrite_existing_files(self):
        self.output.write_bytes(b"keep this file")

        with self.assertRaises(FileExistsError):
            clean_workbook(self.source, self.output, self.report)

        self.assertEqual(self.output.read_bytes(), b"keep this file")
        self.assertFalse(self.report.exists())

    def test_refuses_to_overwrite_input(self):
        with self.assertRaisesRegex(ValueError, "must not overwrite"):
            clean_workbook(self.source, self.source, self.report)

        self.assertTrue(self.source.exists())

    def test_processes_each_worksheet(self):
        workbook = load_workbook(self.source)
        second_sheet = workbook.create_sheet("Other")
        second_sheet.append(["Code"])
        second_sheet.append(["x"])
        second_sheet.append(["x"])
        workbook.save(self.source)
        workbook.close()

        summary = clean_workbook(self.source, self.output, self.report)

        self.assertEqual(summary["sheet_count"], 2)
        self.assertEqual(summary["sheets"][1]["duplicate_rows_removed"], 1)


if __name__ == "__main__":
    unittest.main()
