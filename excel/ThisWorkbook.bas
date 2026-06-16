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

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo ErrorHandler

    ' Avoid recursion, multi-cell edits, header row, formulas
    If Application.EnableEvents = False Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.HasFormula Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh

    If Sh.Name = "Procurement" And Target.Column = 2 Then
        Application.EnableEvents = False
        RefreshProcurementArticleDropdown ws, Target.Row, CStr(Target.Value)
        Application.EnableEvents = True
    ElseIf Sh.Name = "Stock_Movements" And Target.Column = 2 Then
        Application.EnableEvents = False
        RefreshStockMovementBatchDropdown ws, Target.Row, CStr(Target.Value)
        Application.EnableEvents = True
    End If
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
End Sub
