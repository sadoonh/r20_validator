from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from excel_audit import VBA_AUDIT_MODULE, install_audit_log


class ExcelAuditTests(unittest.TestCase):
    def test_requires_windows_excel(self) -> None:
        with TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "processed.xlsx"
            workbook_path.touch()
            with patch("excel_audit.sys.platform", "linux"):
                with self.assertRaisesRegex(RuntimeError, "requires Windows"):
                    install_audit_log(workbook_path)

    def test_macro_records_manual_changes_with_full_audit_context(self) -> None:
        self.assertIn("Workbook_SheetChange", VBA_AUDIT_MODULE)
        self.assertIn("Workbook_SheetSelectionChange", VBA_AUDIT_MODULE)
        self.assertIn("oldValue = selectedValues(cacheKey)", VBA_AUDIT_MODULE)
        self.assertIn("oldValue = auditValues(cacheKey)", VBA_AUDIT_MODULE)
        self.assertIn("newValue = AuditCellValue(changedCell)", VBA_AUDIT_MODULE)
        self.assertIn("changedBy = Application.UserName", VBA_AUDIT_MODULE)
        self.assertIn(
            'Array("Timestamp", "User", "Sheet", "Cell", "Old Value", "New Value")',
            VBA_AUDIT_MODULE,
        )
        self.assertNotIn("Workbook_SheetCalculate", VBA_AUDIT_MODULE)

    def test_macro_initializes_atomically_and_marks_cache_misses(self) -> None:
        self.assertIn('Set newValues = CreateObject("Scripting.Dictionary")', VBA_AUDIT_MODULE)
        self.assertIn("Set auditValues = newValues", VBA_AUDIT_MODULE)
        self.assertIn('UNKNOWN_OLD_VALUE As String = "[not captured]"', VBA_AUDIT_MODULE)
        self.assertIn("loggedOldValue = UNKNOWN_OLD_VALUE", VBA_AUDIT_MODULE)
        self.assertIn("auditValues(cacheKey) = newValue", VBA_AUDIT_MODULE)
        self.assertIn("selectedValues(cacheKey) = newValue", VBA_AUDIT_MODULE)

    def test_macro_protects_only_the_audit_sheet(self) -> None:
        self.assertIn("EnsureAuditSheet.Unprotect", VBA_AUDIT_MODULE)
        self.assertIn(
            "EnsureAuditSheet.Protect UserInterfaceOnly:=True",
            VBA_AUDIT_MODULE,
        )
        self.assertNotIn("Me.Protect", VBA_AUDIT_MODULE)

    def test_macro_autofits_audit_columns_once_per_change_batch(self) -> None:
        self.assertEqual(
            VBA_AUDIT_MODULE.count("ResizeAuditValueColumns auditSheet"), 1
        )
        self.assertIn('auditSheet.Columns("A:D").AutoFit', VBA_AUDIT_MODULE)
        self.assertIn("AUDIT_VALUE_MAX_WIDTH As Double = 60", VBA_AUDIT_MODULE)
        self.assertIn('auditSheet.Range("E:F").Columns', VBA_AUDIT_MODULE)
        self.assertIn("valueColumn.AutoFit", VBA_AUDIT_MODULE)
        self.assertIn(
            "valueColumn.ColumnWidth = AUDIT_VALUE_MAX_WIDTH",
            VBA_AUDIT_MODULE,
        )
        self.assertIn("valueColumn.WrapText = True", VBA_AUDIT_MODULE)
        self.assertIn(
            'auditSheet.Rows(firstRow & ":" & lastRow).AutoFit',
            VBA_AUDIT_MODULE,
        )

    def test_macro_skips_protected_helper_cells(self) -> None:
        self.assertIn("ShouldAuditCell(Sh, changedCell)", VBA_AUDIT_MODULE)
        self.assertIn("ShouldAuditCell(currentSheet, currentCell)", VBA_AUDIT_MODULE)
        self.assertIn("If Sh.ProtectContents Then", VBA_AUDIT_MODULE)
        self.assertIn("ShouldAuditCell = Not targetCell.Locked", VBA_AUDIT_MODULE)
        self.assertIn("formulaValue = targetCell.Formula", VBA_AUDIT_MODULE)
        self.assertIn("If Err.Number = 0 Then", VBA_AUDIT_MODULE)


if __name__ == "__main__":
    unittest.main()
