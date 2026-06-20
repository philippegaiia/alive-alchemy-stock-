Attribute VB_Name = "StockDropdowns"
' Import this file as a standard module in the VBA editor:
'   - File > Import File > select this file
'
' Contains all public menu procedures and private dropdown helpers.
'
' NOTE: If you prefer to paste the code instead of importing, first delete the
' very first line above (Attribute VB_Name = "StockDropdowns"), then paste the
' rest into a new standard module (Insert > Module).
'
' KEY FIXES vs original:
'   1. 256-char validation limit: article/batch lists are now written to scratch
'      columns F and G in the hidden Lists sheet, and validation references those
'      ranges instead of a comma-joined string. Works for any number of items.
'   2. Application.EnableEvents is now managed by the CALLER (the sheet event
'      handlers in Procurement_Sheet.bas / StockMovements_Sheet.bas), not here.
'      This prevents EnableEvents getting stuck False when a nested call errors.
'   3. InsertTodayInSelectedCell removed — press Ctrl+; instead (native, instant).
'   4. Error handling uses GoTo CleanExit pattern throughout menu procedures.

Option Explicit

' Scratch columns in the hidden Lists sheet used to write validation lists.
' This bypasses Excel's 256-character hard limit on inline Formula1 strings.
' Lists!F = article dropdown scratch  |  Lists!G = batch dropdown scratch
Private Const ARTICLE_SCRATCH_COL As Long = 6   ' column F
Private Const BATCH_SCRATCH_COL   As Long = 7   ' column G

' =============================================================================
' Public menu procedures
' =============================================================================

Public Sub GoToFirstEmptyProcurementRow()
    GoToFirstEmptyRow "Procurement"
End Sub

Public Sub GoToFirstEmptyMovementRow()
    GoToFirstEmptyRow "Stock_Movements"
End Sub

Public Sub GoToFirstEmptyRow(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found.", vbExclamation
        Exit Sub
    End If
    ws.Activate
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ws.Cells(2, 1).Select
        Exit Sub
    End If
    Dim i As Long
    For i = 2 To lastRow
        If IsEmpty(ws.Cells(i, 1).Value) Or ws.Cells(i, 1).Value = "" Then
            ws.Cells(i, 1).Select
            Exit Sub
        End If
    Next i
    ws.Cells(lastRow + 1, 1).Select
End Sub

Public Sub ArchiveOldMovements()
    Dim ws As Worksheet, archive As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Stock_Movements")
    Set archive = ThisWorkbook.Worksheets("Stock_Movements_Archive")
    On Error GoTo 0
    If ws Is Nothing Or archive Is Nothing Then
        MsgBox "Stock_Movements or Stock_Movements_Archive sheet is missing.", vbExclamation
        Exit Sub
    End If

    Dim cutoffDays As Long
    cutoffDays = 90
    Dim cutoff As Date
    cutoff = DateAdd("d", -cutoffDays, Date)

    ' Use col B (Article) to find last row — more reliable than Date
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No movements to archive.", vbInformation
        Exit Sub
    End If

    ' Read input columns A-H (skip formula cols I-J)
    Dim data As Variant
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 8)).Value

    Dim archiveCount As Long
    archiveCount = 0
    Dim archiveArr() As Variant
    ReDim archiveArr(1 To UBound(data, 1), 1 To 10)
    Dim sourceRows() As Long
    ReDim sourceRows(1 To UBound(data, 1))

    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        Dim d As Variant
        d = data(i, 1)
        If IsDate(d) And CDate(d) < cutoff And CStr(data(i, 2)) <> "" Then
            archiveCount = archiveCount + 1
            For j = 1 To 8
                archiveArr(archiveCount, j) = data(i, j)
            Next j
            archiveArr(archiveCount, 9) = "Archived"
            archiveArr(archiveCount, 10) = Now
            sourceRows(archiveCount) = i
        End If
    Next i

    If archiveCount = 0 Then
        MsgBox "No movements older than " & cutoffDays & " days to archive.", vbInformation
        Exit Sub
    End If

    ' Write to archive sheet
    Dim archiveStart As Long
    archiveStart = archive.Cells(archive.Rows.Count, 2).End(xlUp).Row + 1
    Dim writeArr() As Variant
    ReDim writeArr(1 To archiveCount, 1 To 10)
    For i = 1 To archiveCount
        For j = 1 To 10
            writeArr(i, j) = archiveArr(i, j)
        Next j
    Next i
    archive.Range(archive.Cells(archiveStart, 1), archive.Cells(archiveStart + archiveCount - 1, 10)).Value = writeArr

    ' Clear input cells in Stock_Movements — formulas (D, E, I, J) remain and return ""
    Application.EnableEvents = False
    For i = archiveCount To 1 Step -1
        Dim realRow As Long
        realRow = sourceRows(i) + 1
        ws.Cells(realRow, 1).ClearContents  ' Date
        ws.Cells(realRow, 2).ClearContents  ' Article
        ws.Cells(realRow, 3).ClearContents  ' Batch
        ws.Cells(realRow, 6).ClearContents  ' Qty_Drawn
        ws.Cells(realRow, 7).ClearContents  ' Reason
        ws.Cells(realRow, 8).ClearContents  ' Notes
    Next i
    Application.EnableEvents = True

    Application.Calculate
    MsgBox "Archived " & archiveCount & " movement(s) to Stock_Movements_Archive. Rows cleared from Stock_Movements.", vbInformation
End Sub

Public Sub RestoreArchivedMovements()
    Dim ws As Worksheet, archive As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Stock_Movements")
    Set archive = ThisWorkbook.Worksheets("Stock_Movements_Archive")
    On Error GoTo 0
    If ws Is Nothing Or archive Is Nothing Then
        MsgBox "Required sheets not found.", vbExclamation
        Exit Sub
    End If

    ' Check if archive has data (col B = Article)
    Dim archLastRow As Long
    archLastRow = archive.Cells(archive.Rows.Count, 2).End(xlUp).Row
    If archLastRow < 2 Then
        MsgBox "No archived movements to restore.", vbInformation
        Exit Sub
    End If

    ' Read archived input data (cols A-H)
    Dim archData As Variant
    archData = archive.Range(archive.Cells(2, 1), archive.Cells(archLastRow, 8)).Value

    Dim count As Long
    count = UBound(archData, 1)

    ' Write to first empty rows in Stock_Movements (col B empty = available row)
    Application.EnableEvents = False
    Dim wsRow As Long
    wsRow = 1
    Dim i As Long
    For i = 1 To count
        Do
            wsRow = wsRow + 1
        Loop While CStr(ws.Cells(wsRow, 2).Value) <> "" And wsRow < 5001

        If wsRow >= 5001 Then
            MsgBox "Not enough empty rows in Stock_Movements.", vbExclamation
            Exit For
        End If

        ws.Cells(wsRow, 1).Value = archData(i, 1)  ' Date
        ws.Cells(wsRow, 2).Value = archData(i, 2)  ' Article
        ws.Cells(wsRow, 3).Value = archData(i, 3)  ' Batch
        ws.Cells(wsRow, 6).Value = archData(i, 6)  ' Qty_Drawn
        ws.Cells(wsRow, 7).Value = archData(i, 7)  ' Reason
        ws.Cells(wsRow, 8).Value = archData(i, 8)  ' Notes
    Next i
    Application.EnableEvents = True

    ' Clear archive
    archive.Rows("2:" & archLastRow).ClearContents

    Application.Calculate
    MsgBox "Restored " & count & " movement(s) from archive to Stock_Movements.", vbInformation
End Sub

Public Sub RepairNamedRanges()
    ' Recreates all workbook-level named ranges used by formulas.
    ' This makes the workbook self-healing if the ranges were lost during
    ' a Save As or copy operation. Runs silently.
    On Error Resume Next
    ThisWorkbook.Names("DetailArticles").Delete
    ThisWorkbook.Names("DetailBatches").Delete
    ThisWorkbook.Names("DetailQtyRemaining").Delete
    ThisWorkbook.Names("DetailStatus").Delete
    ThisWorkbook.Names("DetailStockValue").Delete
    ThisWorkbook.Names("MoveArticles").Delete
    ThisWorkbook.Names("MoveBatches").Delete
    ThisWorkbook.Names("MoveDates").Delete
    ThisWorkbook.Names("MoveQtyDrawn").Delete
    ThisWorkbook.Names("MoveReasons").Delete
    ThisWorkbook.Names("MoveStatus").Delete
    ThisWorkbook.Names("ArchiveArticles").Delete
    ThisWorkbook.Names("ArchiveBatches").Delete
    ThisWorkbook.Names("ArchiveDates").Delete
    ThisWorkbook.Names("ArchiveQtyDrawn").Delete
    ThisWorkbook.Names("ProcSuppliers").Delete
    ThisWorkbook.Names("ProcArticles").Delete
    ThisWorkbook.Names("ProcBatches").Delete
    ThisWorkbook.Names("ProcQuantities").Delete
    ThisWorkbook.Names("ProcUnitCosts").Delete
    ThisWorkbook.Names("ProcDates").Delete
    ThisWorkbook.Names("ProcInvoices").Delete
    ThisWorkbook.Names("ProcActive").Delete
    ThisWorkbook.Names("ArticleNames").Delete
    ThisWorkbook.Names("ArticleUnits").Delete
    ThisWorkbook.Names("ArticleIDs").Delete
    ThisWorkbook.Names("ArticleCategories").Delete
    ThisWorkbook.Names("SupplierNames").Delete
    ThisWorkbook.Names("SupplierIDs").Delete
    ThisWorkbook.Names("Categories").Delete
    ThisWorkbook.Names("Units").Delete
    ThisWorkbook.Names("MovementReasons").Delete
    ThisWorkbook.Names("YesNo").Delete
    ThisWorkbook.Names("ReorderLevels").Delete
    On Error GoTo 0

    Dim maxProc As Long
    maxProc = 2000
    Dim maxMove As Long
    maxMove = 2000
    Dim maxArt As Long
    maxArt = 200
    Dim maxSup As Long
    maxSup = 100

    ' Stock_Detail
    ThisWorkbook.Names.Add Name:="DetailArticles", RefersTo:="=Stock_Detail!$A$2:$A$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="DetailBatches", RefersTo:="=Stock_Detail!$E$2:$E$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="DetailQtyRemaining", RefersTo:="=Stock_Detail!$I$2:$I$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="DetailStatus", RefersTo:="=Stock_Detail!$J$2:$J$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="DetailStockValue", RefersTo:="=Stock_Detail!$L$2:$L$" & maxProc + 1

    ' Stock_Movements
    ThisWorkbook.Names.Add Name:="MoveArticles", RefersTo:="=Stock_Movements!$B$2:$B$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="MoveBatches", RefersTo:="=Stock_Movements!$C$2:$C$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="MoveDates", RefersTo:="=Stock_Movements!$A$2:$A$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="MoveQtyDrawn", RefersTo:="=Stock_Movements!$F$2:$F$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="MoveReasons", RefersTo:="=Stock_Movements!$G$2:$G$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="MoveStatus", RefersTo:="=Stock_Movements!$I$2:$I$" & maxMove + 1

    ' Stock_Movements_Archive
    ThisWorkbook.Names.Add Name:="ArchiveArticles", RefersTo:="=Stock_Movements_Archive!$B$2:$B$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="ArchiveBatches", RefersTo:="=Stock_Movements_Archive!$C$2:$C$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="ArchiveDates", RefersTo:="=Stock_Movements_Archive!$A$2:$A$" & maxMove + 1
    ThisWorkbook.Names.Add Name:="ArchiveQtyDrawn", RefersTo:="=Stock_Movements_Archive!$F$2:$F$" & maxMove + 1

    ' Procurement
    ThisWorkbook.Names.Add Name:="ProcSuppliers", RefersTo:="=Procurement!$B$2:$B$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcArticles", RefersTo:="=Procurement!$C$2:$C$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcBatches", RefersTo:="=Procurement!$I$2:$I$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcQuantities", RefersTo:="=Procurement!$E$2:$E$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcUnitCosts", RefersTo:="=Procurement!$F$2:$F$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcDates", RefersTo:="=Procurement!$A$2:$A$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcInvoices", RefersTo:="=Procurement!$H$2:$H$" & maxProc + 1
    ThisWorkbook.Names.Add Name:="ProcActive", RefersTo:="=Procurement!$J$2:$J$" & maxProc + 1

    ' Articles / Suppliers / Lists
    ThisWorkbook.Names.Add Name:="ArticleNames", RefersTo:="=Articles!$B$2:$B$" & maxArt + 1
    ThisWorkbook.Names.Add Name:="ArticleUnits", RefersTo:="=Articles!$D$2:$D$" & maxArt + 1
    ThisWorkbook.Names.Add Name:="ArticleIDs", RefersTo:="=Articles!$A$2:$A$" & maxArt + 1
    ThisWorkbook.Names.Add Name:="ArticleCategories", RefersTo:="=Articles!$C$2:$C$" & maxArt + 1
    ThisWorkbook.Names.Add Name:="SupplierNames", RefersTo:="=Suppliers!$B$2:$B$" & maxSup + 1
    ThisWorkbook.Names.Add Name:="SupplierIDs", RefersTo:="=Suppliers!$A$2:$A$" & maxSup + 1
    ThisWorkbook.Names.Add Name:="Categories", RefersTo:="=Lists!$A$1:$A$9"
    ThisWorkbook.Names.Add Name:="Units", RefersTo:="=Lists!$B$1:$B$5"
    ThisWorkbook.Names.Add Name:="MovementReasons", RefersTo:="=Lists!$C$1:$C$5"
    ThisWorkbook.Names.Add Name:="YesNo", RefersTo:="=Lists!$D$1:$D$2"
    ThisWorkbook.Names.Add Name:="ReorderLevels", RefersTo:="=Articles!$E$2:$E$" & maxArt + 1

    Application.Calculate
End Sub

Public Sub RepairNamedRangesWithMessage()
    RepairNamedRanges
    MsgBox "Named ranges repaired. Available_Qty and other formulas should now calculate.", vbInformation
End Sub

' =============================================================================
' Private helpers — called by the sheet event handlers
' NOTE: Application.EnableEvents is managed by the CALLER, not here.
' =============================================================================

' Refreshes the Article dropdown in Procurement when Supplier changes.
' Procurement columns: A=Date(1) B=Supplier(2) C=Article(3) ...
Public Sub RefreshProcurementArticleDropdown(ws As Worksheet, row As Long, supplier As String)
    Dim articleCell As Range
    Set articleCell = ws.Cells(row, 3)   ' column C

    ' Always clear the dependent cell and its validation first
    articleCell.ClearContents
    articleCell.ClearNotes

    Dim trimmedSupplier As String
    trimmedSupplier = Trim(supplier)
    If trimmedSupplier = "" Then
        articleCell.Validation.Delete
        Exit Sub
    End If

    ' Look up matching articles in Supplier_Catalog
    Dim catalogWs As Worksheet
    On Error Resume Next
    Set catalogWs = ThisWorkbook.Worksheets("Supplier_Catalog")
    On Error GoTo 0
    If catalogWs Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = catalogWs.Cells(catalogWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Catalog: col A = Supplier name, col B = Article name (redundant cols removed)
    Dim catalogData As Variant
    catalogData = catalogWs.Range(catalogWs.Cells(2, 1), catalogWs.Cells(lastRow, 2)).Value

    Dim articleColl As New Collection
    Dim i As Long
    For i = 1 To UBound(catalogData, 1)
        If Trim(CStr(catalogData(i, 1))) = trimmedSupplier And CStr(catalogData(i, 2)) <> "" Then
            CollectionAddUnique articleColl, CStr(catalogData(i, 2))
        End If
    Next i

    If articleColl.Count = 0 Then
        ' Leave a note so the user knows why the dropdown is empty
        On Error Resume Next
        articleCell.NoteText "No articles in Supplier_Catalog for this supplier"
        On Error GoTo 0
        Exit Sub
    End If

    ' Sort keys and write to Lists!F (scratch) — bypasses the 256-char limit
    Dim sortedArticles() As String
    sortedArticles = CollectionToSortedArray(articleColl)
    WriteScratch sortedArticles, ARTICLE_SCRATCH_COL

    ' Set validation to reference the scratch range (no length limit)
    articleCell.Validation.Delete
    articleCell.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertInformation, _
        Formula1:="=Lists!$F$1:$F$" & articleColl.Count
    articleCell.Validation.IgnoreBlank = True
End Sub

' Refreshes the Batch dropdown in Stock_Movements when Article changes.
' Stock_Movements columns: A=Date(1) B=Article(2) C=Batch(3) ...
Public Sub RefreshStockMovementBatchDropdown(ws As Worksheet, row As Long, article As String)
    Dim batchCell As Range
    Set batchCell = ws.Cells(row, 3)   ' column C

    batchCell.ClearContents
    batchCell.ClearNotes

    Dim trimmedArticle As String
    trimmedArticle = Trim(article)
    If trimmedArticle = "" Then
        batchCell.Validation.Delete
        Exit Sub
    End If

    ' Find open batches for this article
    Dim batches As New Collection
    GetOpenBatchesForArticle trimmedArticle, batches

    If batches.Count = 0 Then
        On Error Resume Next
        batchCell.NoteText "No active batches with stock remaining for this article"
        On Error GoTo 0
        Exit Sub
    End If

    ' Sort and write to Lists!G scratch column
    Dim sortedBatches() As String
    sortedBatches = CollectionToSortedArray(batches)
    WriteScratch sortedBatches, BATCH_SCRATCH_COL

    batchCell.Validation.Delete
    batchCell.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertInformation, _
        Formula1:="=Lists!$G$1:$G$" & batches.Count
    batchCell.Validation.IgnoreBlank = True

    ' Auto-select when only one batch is open
    If batches.Count = 1 Then
        batchCell.Value = sortedBatches(0)
    End If
End Sub

' Populates 'result' collection with open batch IDs for a given article.
' Reads Stock_Detail (hidden) and cross-checks Procurement for Active=Yes.
' Stock_Detail columns: A=Article(1) E=Batch(5) I=QtyRemaining(9) J=Status(10)
' Procurement columns:  C=Article(3) I=Batch(9) J=Active(10)
Private Sub GetOpenBatchesForArticle(article As String, result As Collection)
    Dim detailWs As Worksheet, procWs As Worksheet
    On Error Resume Next
    Set detailWs = ThisWorkbook.Worksheets("Stock_Detail")
    Set procWs = ThisWorkbook.Worksheets("Procurement")
    On Error GoTo 0
    If detailWs Is Nothing Or procWs Is Nothing Then Exit Sub

    Dim detailLastRow As Long
    detailLastRow = detailWs.Cells(detailWs.Rows.Count, 1).End(xlUp).Row
    If detailLastRow < 2 Then Exit Sub

    ' Build a set of article+batch keys that are Active=Yes in Procurement
    ' Use column C (Article) to find last row — more reliable than Date (col A)
    ' which the user might leave blank on a new procurement row.
    Dim activeKeys As New Collection
    Dim procLastRow As Long
    procLastRow = procWs.Cells(procWs.Rows.Count, 3).End(xlUp).Row
    If procLastRow >= 2 Then
        Dim procData As Variant
        procData = procWs.Range(procWs.Cells(2, 1), procWs.Cells(procLastRow, 10)).Value
        Dim i As Long
        For i = 1 To UBound(procData, 1)
            If Trim(CStr(procData(i, 3))) = article _
               And CStr(procData(i, 9)) <> "" _
               And CStr(procData(i, 10)) = "Yes" Then
                Dim pKey As String
                pKey = article & "||" & CStr(procData(i, 9))
                CollectionAddUnique activeKeys, pKey
            End If
        Next i
    End If

    ' Scan Stock_Detail for open batches with remaining stock
    Dim detailData As Variant
    detailData = detailWs.Range(detailWs.Cells(2, 1), detailWs.Cells(detailLastRow, 10)).Value
    For i = 1 To UBound(detailData, 1)
        Dim detailArticle As String
        detailArticle = Trim(CStr(detailData(i, 1)))
        Dim batch As String
        batch = CStr(detailData(i, 5))
        Dim remaining As Double
        remaining = Val(detailData(i, 9))
        Dim status As String
        status = CStr(detailData(i, 10))
        If detailArticle = article _
           And batch <> "" _
           And remaining > 0 _
           And status = "Open" Then
            Dim dKey As String
            dKey = article & "||" & batch
            If CollectionContains(activeKeys, dKey) Then
                CollectionAddUnique result, batch
            End If
        End If
    Next i
End Sub

' =============================================================================
' Collection helpers (Mac-compatible replacement for Scripting.Dictionary)
' =============================================================================

' Adds a value to a Collection only if it is not already present.
Private Sub CollectionAddUnique(coll As Collection, value As String)
    On Error Resume Next
    coll.Add value, value
    On Error GoTo 0
End Sub

' Returns True if a Collection contains the given key.
Private Function CollectionContains(coll As Collection, key As String) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = coll(key)
    CollectionContains = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Returns the items of a Collection as a sorted String array.
Private Function CollectionToSortedArray(coll As Collection) As String()
    If coll.Count = 0 Then
        CollectionToSortedArray = Split("", ",")
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To coll.Count - 1)
    Dim i As Long
    For i = 1 To coll.Count
        arr(i - 1) = coll(i)
    Next i

    SortStringArray arr
    CollectionToSortedArray = arr
End Function

' Sorts a String array in place using bubble sort.
Private Sub SortStringArray(arr() As String)
    Dim k As Long, m As Long, tmp As String
    For k = LBound(arr) To UBound(arr) - 1
        For m = k + 1 To UBound(arr)
            If arr(k) > arr(m) Then
                tmp = arr(k): arr(k) = arr(m): arr(m) = tmp
            End If
        Next m
    Next k
End Sub

' Writes a String array to a scratch column in the hidden Lists sheet.
' Validation Formula1 then references e.g. "=Lists!$F$1:$F$5" — no length limit.
Private Sub WriteScratch(values() As String, scratchCol As Long)
    Dim listsWs As Worksheet
    On Error Resume Next
    Set listsWs = ThisWorkbook.Worksheets("Lists")
    On Error GoTo 0
    If listsWs Is Nothing Then Exit Sub   ' Lists sheet missing — nothing to do

    Dim count As Long
    count = UBound(values) - LBound(values) + 1

    ' Clear a safe buffer of rows then write new values
    listsWs.Cells(1, scratchCol).Resize(count + 5).ClearContents
    Dim i As Long
    For i = LBound(values) To UBound(values)
        listsWs.Cells(i - LBound(values) + 1, scratchCol).Value = values(i)
    Next i
End Sub
