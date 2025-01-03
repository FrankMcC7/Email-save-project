Option Explicit

' Constants for error handling and configuration
Private Const ERR_NO_FILE_SELECTED As Long = 1001
Private Const ERR_NO_DATA As Long = 1002
Private Const ERR_SHEET_EXISTS As Long = 1003
Private Const ERR_INSUFFICIENT_DATA As Long = 1004
Private Const SHEET_PREFIX As String = "RandomData_"
Private Const MASTER_SHEET_NAME As String = "ApprovedData"
Private Const REVIEW_STATUS_COLUMN As String = "Review Status"
Private Const APPROVED_STATUS As String = "Approved"

' Main entry point for the macro
Public Sub ProcessDataset()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    Dim rng As Range
    Dim filePath As String
    Dim fileDialog As FileDialog
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Initialize file dialog
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select the dataset file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            RaiseError ERR_NO_FILE_SELECTED, "No file selected. Process terminated."
        End If
    End With
    
    ' Open the selected workbook
    Set wb = Workbooks.Open(filePath)
    Set wsSource = wb.Sheets(1)
    
    ' Delete first row
    wsSource.Rows(1).Delete
    
    ' Remove blank rows
    Call RemoveBlankRows(wsSource)
    
    ' Convert range to table
    lastRow = GetLastRow(wsSource)
    lastCol = GetLastColumn(wsSource)
    
    If lastRow < 2 Then
        RaiseError ERR_NO_DATA, "No data found in the worksheet."
    End If
    
    Set rng = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    Set tbl = CreateTable(rng)
    
    ' Filter for 'Approved' status
    FilterTableForApproved tbl
    
    ' Copy filtered data to master file
    Call CopyToMasterFile(wb, tbl)
    
    ' Close source workbook
    wb.Close SaveChanges:=False
    
    ' Start the continuous process
    Call ProcessContinuously(ThisWorkbook.Sheets(MASTER_SHEET_NAME))
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case ERR_NO_FILE_SELECTED
            MsgBox "No file was selected. Process terminated.", vbExclamation
        Case ERR_NO_DATA
            MsgBox "No data found in the worksheet.", vbExclamation
        Case ERR_SHEET_EXISTS
            MsgBox "Sheet already exists.", vbExclamation
        Case Else
            MsgBox "An error occurred: " & Err.Description, vbCritical
    End Select
    Resume CleanExit
End Sub

' Remove blank rows from worksheet
Private Sub RemoveBlankRows(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = GetLastRow(ws)
    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "RemoveBlankRows", "Error removing blank rows: " & Err.Description
End Sub

' Get last used row in column A
Private Function GetLastRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Then GetLastRow = 0
End Function

' Get last used column in row 1
Private Function GetLastColumn(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Err.Number <> 0 Then GetLastColumn = 0
End Function

' Create table from range
Private Function CreateTable(ByVal rng As Range) As ListObject
    On Error GoTo ErrorHandler
    
    Set CreateTable = rng.Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "CreateTable", "Error creating table: " & Err.Description
End Function

' Filter table for approved status
Private Sub FilterTableForApproved(ByVal tbl As ListObject)
    On Error GoTo ErrorHandler
    
    Dim statusColIndex As Long
    statusColIndex = GetColumnNumber(tbl, REVIEW_STATUS_COLUMN)
    
    If statusColIndex > 0 Then
        tbl.Range.AutoFilter Field:=statusColIndex, Criteria1:=APPROVED_STATUS
    Else
        Err.Raise ERR_NO_DATA, "FilterTableForApproved", "Review Status column not found"
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "FilterTableForApproved", "Error filtering table: " & Err.Description
End Sub

' Get column number by name
Private Function GetColumnNumber(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim col As ListColumn
    Dim colNum As Long
    
    colNum = 1
    For Each col In tbl.ListColumns
        If StrComp(col.Name, colName, vbTextCompare) = 0 Then
            GetColumnNumber = colNum
            Exit Function
        End If
        colNum = colNum + 1
    Next col
    
    GetColumnNumber = 0
    Exit Function
    
ErrorHandler:
    GetColumnNumber = 0
End Function

' Copy data to master file
Private Sub CopyToMasterFile(ByVal sourceWb As Workbook, ByVal sourceTbl As ListObject)
    On Error GoTo ErrorHandler
    
    Dim wsMaster As Worksheet
    
    ' Check if ApprovedData sheet exists, create if not
    If Not SheetExists(MASTER_SHEET_NAME) Then
        Set wsMaster = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaster.Name = MASTER_SHEET_NAME
    Else
        Set wsMaster = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    End If
    
    ' Clear existing data in master sheet
    wsMaster.Cells.Clear
    
    ' Copy filtered data to master sheet
    With sourceTbl.Range.SpecialCells(xlCellTypeVisible)
        .Copy
        wsMaster.Range("A1").PasteSpecial xlPasteValues
        wsMaster.Range("A1").PasteSpecial xlPasteFormats
    End With
    
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "CopyToMasterFile", "Error copying to master file: " & Err.Description
End Sub

' Check if sheet exists
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function

' Process data continuously
Private Sub ProcessContinuously(ByVal wsMaster As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim randRows() As Long
    Dim pauseTime As Long
    
    Do
        ' Create 5 sheets with random data
        lastRow = GetLastRow(wsMaster)
        If lastRow <= 1 Then Exit Sub
        
        For i = 1 To 5
            ' Create new sheet
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = SHEET_PREFIX & i
            
            ' Get 200 random rows
            ReDim randRows(1 To 200)
            Call GetRandomRows(lastRow, randRows)
            
            ' Copy headers
            wsMaster.Rows(1).Copy ws.Rows(1)
            
            ' Copy random rows
            Call CopyRandomRows(wsMaster, ws, randRows)
        Next i
        
        ' Random pause between 5 and 30 seconds
        pauseTime = Int((25 * Rnd) + 5)
        Application.Wait Now + TimeValue("00:00:" & pauseTime)
        
        ' Delete the created sheets
        Application.DisplayAlerts = False
        For i = 1 To 5
            If SheetExists(SHEET_PREFIX & i) Then
                ThisWorkbook.Sheets(SHEET_PREFIX & i).Delete
            End If
        Next i
        Application.DisplayAlerts = True
        
        DoEvents ' Allow for user interruption
    Loop
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "ProcessContinuously", "Error in continuous processing: " & Err.Description
End Sub

' Get random rows
Private Sub GetRandomRows(ByVal totalRows As Long, ByRef randRows() As Long)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim temp As Long
    Dim usedRows() As Boolean
    Dim numRowsToGet As Long
    
    ' Calculate how many rows we need
    numRowsToGet = UBound(randRows)
    
    ' Make sure we don't try to get more rows than available
    If numRowsToGet > (totalRows - 1) Then
        Err.Raise ERR_NO_DATA, "GetRandomRows", "Not enough rows in source data"
    End If
    
    ' Initialize tracking array (subtract 1 to account for header)
    ReDim usedRows(1 To totalRows - 1)
    
    ' Generate random rows
    For i = 1 To numRowsToGet
        Do
            ' Generate random row number (skip header row)
            temp = Int((totalRows - 1) * Rnd + 2)
        Loop Until Not usedRows(temp - 1)
        
        randRows(i) = temp
        usedRows(temp - 1) = True
    Next i
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "GetRandomRows", "Error generating random rows: " & Err.Description
End Sub

' Copy random rows
Private Sub CopyRandomRows(ByVal sourceWs As Worksheet, ByVal destWs As Worksheet, ByRef randRows() As Long)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim lastCol As Long
    
    lastCol = GetLastColumn(sourceWs)
    
    For i = LBound(randRows) To UBound(randRows)
        sourceWs.Rows(randRows(i)).Copy
        destWs.Rows(i + 1).PasteSpecial xlPasteValues
    Next i
    
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "CopyRandomRows", "Error copying random rows: " & Err.Description
End Sub

' Raise custom error
Private Sub RaiseError(ByVal errorNumber As Long, ByVal errorMessage As String)
    Err.Raise errorNumber, "ProcessDataset", errorMessage
End Sub
