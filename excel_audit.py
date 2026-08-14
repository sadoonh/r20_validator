"""Install the Excel VBA change-audit module into a processed workbook."""

from __future__ import annotations

from pathlib import Path
import sys


VBA_AUDIT_MODULE = r'''Option Explicit

Private Const AUDIT_SHEET_NAME As String = "_AuditLog"
Private Const UNKNOWN_OLD_VALUE As String = "[not captured]"
Private Const AUDIT_VALUE_MAX_WIDTH As Double = 60
Private auditValues As Object
Private selectedValues As Object
Private auditWriting As Boolean

Private Sub Workbook_Open()
    InitializeAuditCache
End Sub

Private Sub Workbook_Activate()
    If auditValues Is Nothing Then InitializeAuditCache
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If auditWriting Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If Sh.Name = AUDIT_SHEET_NAME Then Exit Sub

    If auditValues Is Nothing Then InitializeAuditCache
    CaptureSelectedValues Sh, Target
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim changedCells As Range
    Dim changedCell As Range
    Dim auditSheet As Worksheet
    Dim oldValue As Variant
    Dim newValue As Variant
    Dim loggedOldValue As Variant
    Dim cacheKey As String
    Dim changedAt As Date
    Dim changedBy As String
    Dim oldValueKnown As Boolean
    Dim appendedRow As Long
    Dim firstAppendedRow As Long
    Dim lastAppendedRow As Long

    If auditWriting Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If Sh.Name = AUDIT_SHEET_NAME Then Exit Sub

    If auditValues Is Nothing Then InitializeAuditCache
    If auditValues Is Nothing Then Exit Sub

    Set changedCells = Intersect(Target, Sh.UsedRange)
    If changedCells Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    auditWriting = True
    Application.EnableEvents = False
    changedAt = Now
    changedBy = Application.UserName
    If Len(Trim$(changedBy)) = 0 Then changedBy = "[unknown]"

    For Each changedCell In changedCells.Cells
        If ShouldAuditCell(Sh, changedCell) Then
            cacheKey = AuditCacheKey(Sh, changedCell)
            oldValueKnown = False

            If Not selectedValues Is Nothing Then
                If selectedValues.Exists(cacheKey) Then
                    oldValue = selectedValues(cacheKey)
                    oldValueKnown = True
                End If
            End If

            If Not oldValueKnown Then
                If auditValues.Exists(cacheKey) Then
                    oldValue = auditValues(cacheKey)
                    oldValueKnown = True
                End If
            End If

            newValue = AuditCellValue(changedCell)
            If oldValueKnown Then
                loggedOldValue = oldValue
            Else
                loggedOldValue = UNKNOWN_OLD_VALUE
            End If

            If Not oldValueKnown Or Not AuditValuesEqual(oldValue, newValue) Then
                appendedRow = AppendAuditEntry(changedAt, changedBy, Sh.Name, changedCell.Address(False, False), loggedOldValue, newValue)
                If firstAppendedRow = 0 Then firstAppendedRow = appendedRow
                lastAppendedRow = appendedRow
            End If

            auditValues(cacheKey) = newValue
            If Not selectedValues Is Nothing Then
                selectedValues(cacheKey) = newValue
            End If
        End If
    Next changedCell

    If firstAppendedRow > 0 Then
        Set auditSheet = EnsureAuditSheet()
        ResizeAuditValueColumns auditSheet, firstAppendedRow, lastAppendedRow
    End If

CleanUp:
    Application.EnableEvents = True
    auditWriting = False
End Sub

Private Sub CaptureSelectedValues(ByVal Sh As Object, ByVal Target As Range)
    Dim newSelection As Object
    Dim selectedCell As Range
    Dim cacheKey As String

    On Error GoTo CaptureFailed
    Set newSelection = CreateObject("Scripting.Dictionary")

    If Target.CountLarge <= 100000 Then
        For Each selectedCell In Target.Cells
            If ShouldAuditCell(Sh, selectedCell) Then
                cacheKey = AuditCacheKey(Sh, selectedCell)
                newSelection(cacheKey) = AuditCellValue(selectedCell)
            End If
        Next selectedCell
    End If

    Set selectedValues = newSelection
    Exit Sub

CaptureFailed:
    Set selectedValues = Nothing
End Sub

Private Sub InitializeAuditCache()
    Dim auditSheet As Worksheet
    Dim currentSheet As Worksheet
    Dim currentCell As Range
    Dim cacheKey As String
    Dim newValues As Object
    Dim failureMessage As String
    Dim failureContext As String

    On Error GoTo InitializationFailed
    auditWriting = True
    Application.EnableEvents = False
    Set newValues = CreateObject("Scripting.Dictionary")
    failureContext = "creating the audit sheet"
    Set auditSheet = EnsureAuditSheet()

    For Each currentSheet In Me.Worksheets
        If currentSheet.Name <> AUDIT_SHEET_NAME Then
            For Each currentCell In currentSheet.UsedRange.Cells
                failureContext = "caching " & currentSheet.Name & "!" & currentCell.Address(False, False)
                If ShouldAuditCell(currentSheet, currentCell) Then
                    cacheKey = AuditCacheKey(currentSheet, currentCell)
                    newValues(cacheKey) = AuditCellValue(currentCell)
                End If
            Next currentCell
        End If
    Next currentSheet

    Set auditValues = newValues
    Set selectedValues = Nothing
    GoTo CleanUp

InitializationFailed:
    failureMessage = Err.Description
    Set auditValues = Nothing
    Set selectedValues = Nothing

CleanUp:
    Application.EnableEvents = True
    auditWriting = False
    If Len(failureMessage) > 0 Then
        MsgBox "The audit cache could not be initialized while " & failureContext & ": " & failureMessage, vbExclamation, "Audit log"
    End If
End Sub

Private Function EnsureAuditSheet() As Worksheet
    On Error Resume Next
    Set EnsureAuditSheet = Me.Worksheets(AUDIT_SHEET_NAME)
    On Error GoTo 0

    If EnsureAuditSheet Is Nothing Then
        Set EnsureAuditSheet = Me.Worksheets.Add(After:=Me.Worksheets(Me.Worksheets.Count))
        EnsureAuditSheet.Name = AUDIT_SHEET_NAME
        EnsureAuditSheet.Visible = xlSheetHidden
    Else
        EnsureAuditSheet.Unprotect
        If EnsureAuditSheet.Cells(1, 2).Value = "Sheet" Then
            EnsureAuditSheet.Columns(2).Insert
        End If
    End If

    EnsureAuditSheet.Range("A1:F1").Value = Array("Timestamp", "User", "Sheet", "Cell", "Old Value", "New Value")
    EnsureAuditSheet.Rows(1).Font.Bold = True
    EnsureAuditSheet.Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    EnsureAuditSheet.Columns("E:F").NumberFormat = "@"
    EnsureAuditSheet.Protect UserInterfaceOnly:=True
End Function

Private Function AppendAuditEntry(ByVal changedAt As Date, ByVal changedBy As String, ByVal sheetName As String, ByVal cellAddress As String, ByVal oldValue As Variant, ByVal newValue As Variant) As Long
    Dim auditSheet As Worksheet
    Dim nextRow As Long

    Set auditSheet = EnsureAuditSheet()
    nextRow = auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row + 1
    auditSheet.Cells(nextRow, 1).Value = changedAt
    auditSheet.Cells(nextRow, 2).Value2 = changedBy
    auditSheet.Cells(nextRow, 3).Value2 = sheetName
    auditSheet.Cells(nextRow, 4).Value2 = cellAddress
    auditSheet.Cells(nextRow, 5).Value2 = oldValue
    auditSheet.Cells(nextRow, 6).Value2 = newValue
    AppendAuditEntry = nextRow
End Function

Private Sub ResizeAuditValueColumns(ByVal auditSheet As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim valueColumn As Range

    auditSheet.Columns("A:D").AutoFit
    For Each valueColumn In auditSheet.Range("E:F").Columns
        valueColumn.WrapText = False
        valueColumn.AutoFit
        If valueColumn.ColumnWidth > AUDIT_VALUE_MAX_WIDTH Then
            valueColumn.ColumnWidth = AUDIT_VALUE_MAX_WIDTH
        End If
        valueColumn.WrapText = True
    Next valueColumn

    auditSheet.Rows(firstRow & ":" & lastRow).AutoFit
End Sub

Private Function AuditCacheKey(ByVal Sh As Object, ByVal targetCell As Range) As String
    AuditCacheKey = Sh.CodeName & "!" & targetCell.Address(False, False)
End Function

Private Function ShouldAuditCell(ByVal Sh As Object, ByVal targetCell As Range) As Boolean
    If Sh.ProtectContents Then
        ShouldAuditCell = Not targetCell.Locked
    Else
        ShouldAuditCell = True
    End If
End Function

Private Function AuditCellValue(ByVal targetCell As Range) As Variant
    Dim formulaValue As Variant

    If targetCell.HasFormula Then
        On Error Resume Next
        formulaValue = targetCell.Formula
        If Err.Number = 0 Then
            AuditCellValue = formulaValue
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    If IsError(targetCell.Value) Then
        AuditCellValue = targetCell.Text
    Else
        AuditCellValue = targetCell.Value2
    End If
End Function

Private Function AuditValuesEqual(ByVal firstValue As Variant, ByVal secondValue As Variant) As Boolean
    If IsEmpty(firstValue) And IsEmpty(secondValue) Then
        AuditValuesEqual = True
    ElseIf VarType(firstValue) <> VarType(secondValue) Then
        AuditValuesEqual = False
    Else
        AuditValuesEqual = (CStr(firstValue) = CStr(secondValue))
    End If
End Function
'''


def install_audit_log(workbook_path: Path) -> Path:
    """Inject the audit VBA into an .xlsx workbook and save it as .xlsm."""
    if sys.platform != "win32":
        raise RuntimeError(
            "Change auditing requires Windows with Microsoft Excel installed."
        )
    if workbook_path.suffix.lower() != ".xlsx":
        raise ValueError("The audit installer expects an .xlsx workbook.")

    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Change auditing requires the Windows pywin32 package."
        ) from error

    output_path = workbook_path.with_suffix(".xlsm")
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
        )
        component = workbook.VBProject.VBComponents.Item("ThisWorkbook")
        code_module = component.CodeModule
        if code_module.CountOfLines:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(VBA_AUDIT_MODULE)
        workbook.SaveAs(str(output_path.resolve()), FileFormat=52)
        workbook.Close(SaveChanges=False)
        workbook = None
    except Exception as error:
        raise RuntimeError(
            "Could not install the audit macro. In Excel Trust Center, enable "
            "'Trust access to the VBA project object model' and try again."
        ) from error
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()

    workbook_path.unlink(missing_ok=True)
    return output_path
