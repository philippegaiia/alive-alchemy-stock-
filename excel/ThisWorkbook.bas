' Paste this ENTIRE content into the ThisWorkbook code module:
'   - In the VBA editor, double-click "ThisWorkbook" in Project Explorer
'   - Delete any existing code
'   - Paste this
'   - Save (Cmd+S / Ctrl+S)
'
' Handles sheet-change events for dependent dropdowns on Procurement and
' Stock_Movements. The formulas in the workbook do not require any named
' ranges, so Available_Qty will calculate correctly even without this module.
' The dropdowns only work if the StockDropdowns standard module is also imported.

Option Explicit

Private Sub Workbook_Open()
    ' Re-protect all sheets with UserInterfaceOnly so VBA can modify cells
    ' while users cannot edit locked (formula) cells.
    ' AllowSorting + AllowFiltering preserve autofilter/sort on protected sheets.
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True
    Next ws
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' Double-click Date column (A) in Procurement or Stock_Movements → stamp today
    If (Sh.Name = "Procurement" Or Sh.Name = "Stock_Movements") Then
        If Target.Column = 1 And Target.Row >= 2 Then
            Target.Value = Date
            Cancel = True
        End If
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler

    ' Avoid recursion, multi-cell edits, header row, formulas
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.HasFormula Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    ' Force recalculation when Stock_Register filter changes (row 1 = header area)
    If Sh.Name = "Stock_Register" And Target.Row = 1 And Target.Column = 5 Then
        Application.Calculate
        Exit Sub
    End If

    If Target.Row < 2 Then Exit Sub

    If Sh.Name = "Procurement" And Target.Column = 2 Then
        Application.EnableEvents = False
        ws.Unprotect
        RefreshProcurementArticleDropdown ws, Target.Row, CStr(Target.Value)
        ws.Protect UserInterfaceOnly:=True
        Application.EnableEvents = True
    ElseIf Sh.Name = "Stock_Movements" And Target.Column = 2 Then
        Application.EnableEvents = False
        ws.Unprotect
        RefreshStockMovementBatchDropdown ws, Target.Row, CStr(Target.Value)
        ws.Protect UserInterfaceOnly:=True
        Application.EnableEvents = True
    End If
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
