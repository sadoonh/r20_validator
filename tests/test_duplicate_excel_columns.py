from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.filters import FilterColumn, Filters
from openpyxl.worksheet.table import Table, TableStyleInfo

from duplicate_excel_columns import duplicate_columns


class ExistingTableTests(unittest.TestCase):
    def test_rebuilds_existing_table_and_preserves_its_identity(self) -> None:
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["Name", "Amount"])
            worksheet.append(["Alpha", 10])
            worksheet.append(["Beta", 20])

            table = Table(displayName="ExistingSales", ref="A1:B3")
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9", showRowStripes=True
            )
            table._initialise_columns()
            table.tableColumns[0].name = "Name"
            table.tableColumns[1].name = "Amount"
            table.autoFilter.filterColumn.append(
                FilterColumn(colId=1, filters=Filters(filter=["10"]))
            )
            worksheet.add_table(table)
            workbook.save(input_path)

            output_path = duplicate_columns(input_path, None, header_row=1)
            result = load_workbook(output_path, data_only=False)
            result_sheet = result.active
            rebuilt = result_sheet.tables["ExistingSales"]

            self.assertEqual(rebuilt.ref, "A1:G3")
            self.assertEqual(rebuilt.tableStyleInfo.name, "TableStyleMedium9")
            self.assertEqual(
                [column.name for column in rebuilt.tableColumns],
                [
                    "Name",
                    "Name_control",
                    "Name_variance",
                    "Amount",
                    "Amount_control",
                    "Amount_variance",
                    "variance_filter",
                ],
            )
            self.assertEqual(rebuilt.autoFilter.filterColumn[0].colId, 3)
            self.assertEqual(result_sheet["C2"].value, "=A2=B2")
            self.assertEqual(result_sheet["G2"].value, "=AND(C2,F2)")

    def test_shifts_an_unrelated_table_to_the_right(self) -> None:
        with TemporaryDirectory() as directory:
            input_path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            for row in (("A", "B"), (1, 2), (3, 4)):
                worksheet.append(row)
            worksheet["J5"] = "X"
            worksheet["K5"] = "Y"
            worksheet["J6"] = 5
            worksheet["K6"] = 6
            worksheet.add_table(Table(displayName="Main", ref="A1:B3"))
            worksheet.add_table(Table(displayName="Other", ref="J5:K6"))
            workbook.save(input_path)

            output_path = duplicate_columns(input_path, None, header_row=1)
            result_sheet = load_workbook(output_path).active

            self.assertEqual(result_sheet.tables["Main"].ref, "A1:G3")
            self.assertEqual(result_sheet.tables["Other"].ref, "N5:O6")
            self.assertEqual(result_sheet["N6"].value, 5)
            self.assertEqual(result_sheet["O6"].value, 6)


if __name__ == "__main__":
    unittest.main()
