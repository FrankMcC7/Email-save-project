Option Explicit

' Constants
Private Const ERR_NO_FILE_SELECTED As Long = 1001
Private Const ERR_NO_DATA As Long = 1002
Private Const ERR_SHEET_EXISTS As Long = 1003
Private Const ERR_INSUFFICIENT_DATA As Long = 1004
Private Const SHEET_PREFIX As String = "RandomData_"
Private Const MASTER_SHEET_NAME As String = "ApprovedData"
Private Const REVIEW_STATUS_COLUMN As String = "Review Status"
Private Const APPROVED_STATUS As String = "Approved"
Private Const CHUNK_SIZE As Long = 1000 ' Process data in chunks

' Application state management
Private Type AppState
    Calculation As XlCalculation
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
End Type

' Store original application state
Private originalState As AppState

' Main entry point for the macro
Public Sub ProcessDataset()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsMaster As Worksheet
    Dim filePath As String
    Dim fileDialog As FileDialog
    
    ' Save and set application state
    SaveAppState
    OptimizeAppState
    
    ' Initialize file dialog
    Set fileDialog = InitializeFileDialog
    If fileDialog.Show = -1 Then
        filePath = fileDialog.SelectedItems(1)
    Else
        RaiseError ERR_NO_FILE_SELECTED, "No file selected. Process terminated."
    End If
    
    ' Process source data
    Set wb = ProcessSourceWorkbook(filePath)
    
    ' Start the continuous process
    Call ProcessContinuously(ThisWorkbook.Sheets(MASTER_SHEET_NAME))
    
CleanExit:
    RestoreAppState
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
    RestoreAppState
    Resume CleanExit
End Sub

Private Function InitializeFileDialog() As FileDialog
    Set InitializeFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With InitializeFileDialog
        .Title = "Select the dataset file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        .AllowMultiSelect = False
    End With
End Function

Private Function ProcessSourceWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook
    Dim wsSource As Worksheet
    
    ' Open workbook with UpdateLinks = 0 and ReadOnly = True for better performance
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set wsSource = wb.Sheets(1)
    
    ' Process the data
    wsSource.Rows(1).Delete
    Call RemoveBlankRows(wsSource)
    Call ProcessAndCopyData(wsSource)
    
    ' Close workbook
    wb.Close SaveChanges:=False
    Set ProcessSourceWorkbook = wb
End Function

Private Sub ProcessAndCopyData(ByVal wsSource As Worksheet)
    Dim wsMaster As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim chunk As Range
    Dim targetRow As Long
    
    lastRow = GetLastRow(wsSource)
    lastCol = GetLastColumn(wsSource)
    
    If lastRow < 2 Then Exit Sub
    
    ' Create or clear master sheet
    If Not SheetExists(MASTER_SHEET_NAME) Then
        Set wsMaster = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsMaster.Name = MASTER_SHEET_NAME
    Else
        Set wsMaster = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
        wsMaster.Cells.Clear
    End If
    
    ' Copy headers
    wsSource.Rows(1).Copy wsMaster.Rows(1)
    
    ' Process data in chunks
    targetRow = 2
    For i = 2 To lastRow Step CHUNK_SIZE
        ' Define chunk range
        Set chunk = wsSource.Range( _
            wsSource.Cells(i, 1), _
            wsSource.Cells(Application.Min(i + CHUNK_SIZE - 1, lastRow), lastCol) _
        )
        
        ' Copy chunk values
        chunk.Copy
        wsMaster.Cells(targetRow, 1).PasteSpecial xlPasteValues
        
        targetRow = targetRow + chunk.Rows.Count
        DoEvents ' Allow system to breathe
    Next i
    
    Application.CutCopyMode = False
End Sub

Private Sub ProcessContinuously(ByVal wsMaster As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim controlSheet As Worksheet
    Dim dataCache() As Variant
    Dim lastRow As Long, lastCol As Long
    
    ' Setup control sheet
    Set controlSheet = SetupControlSheet
    
    ' Cache the data for faster processing
    lastRow = GetLastRow(wsMaster)
    lastCol = GetLastColumn(wsMaster)
    dataCache = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(lastRow, lastCol)).Value
    
    ' Main processing loop
    Do While controlSheet.Range("B1").Value <> "Stopped"
        ' Create sample sheets using cached data
        CreateSampleSheets dataCache
        
        ' Random pause
        Application.Wait Now + TimeValue("00:00:" & Int((25 * Rnd) + 5))
        
        ' Delete sample sheets
        DeleteSampleSheets
        
        DoEvents
    Loop
    
    ' Clean up
    Application.DisplayAlerts = False
    controlSheet.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    RestoreAppState
    Err.Raise Err.Number, "ProcessContinuously", "Error in continuous processing: " & Err.Description
End Sub

Private Sub CreateSampleSheets(ByRef dataCache As Variant)
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim randRows() As Long
    Dim targetData() As Variant
    
    ' Initialize random rows array
    ReDim randRows(1 To 200)
    
    For i = 1 To 5
        ' Get random rows
        GetRandomRows UBound(dataCache, 1), randRows
        
        ' Create new sheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_PREFIX & i
        
        ' Prepare target data array
        ReDim targetData(1 To 201, 1 To UBound(dataCache, 2))
        
        ' Copy header row
        For j = 1 To UBound(dataCache, 2)
            targetData(1, j) = dataCache(1, j)
        Next j
        
        ' Copy selected rows
        For j = 1 To 200
            If randRows(j) <= UBound(dataCache, 1) Then
                CopyArrayRow dataCache, targetData, randRows(j), j + 1
            End If
        Next j
        
        ' Write to sheet in one operation
        ws.Range(ws.Cells(1, 1), ws.Cells(201, UBound(dataCache, 2))).Value = targetData
    Next i
End Sub

Private Sub CopyArrayRow(ByRef sourceArray As Variant, ByRef targetArray As Variant, _
                        ByVal sourceRow As Long, ByVal targetRow As Long)
    Dim col As Long
    For col = 1 To UBound(sourceArray, 2)
        targetArray(targetRow, col) = sourceArray(sourceRow, col)
    Next col
End Sub

Private Function SetupControlSheet() As Worksheet
    If Not SheetExists("Control") Then
        Set SetupControlSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        SetupControlSheet.Name = "Control"
        
        ' Add stop button
        Dim btnStop As Shape
        Set btnStop = SetupControlSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 30)
        With btnStop
            .Fill.ForeColor.RGB = RGB(255, 0, 0)
            .TextFrame.Characters.Text = "Stop Process"
            .OnAction = "StopProcess"
        End With
        
        SetupControlSheet.Range("A1").Value = "Process Status:"
        SetupControlSheet.Range("B1").Value = "Running"
    Else
        Set SetupControlSheet = ThisWorkbook.Sheets("Control")
        SetupControlSheet.Range("B1").Value = "Running"
    End If
End Function

Private Sub SaveAppState()
    With originalState
        .Calculation = Application.Calculation
        .EnableEvents = Application.EnableEvents
        .ScreenUpdating = Application.ScreenUpdating
        .DisplayAlerts = Application.DisplayAlerts
    End With
End Sub

Private Sub OptimizeAppState()
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
End Sub

Private Sub RestoreAppState()
    With Application
        .Calculation = originalState.Calculation
        .EnableEvents = originalState.EnableEvents
        .ScreenUpdating = originalState.ScreenUpdating
        .DisplayAlerts = originalState.DisplayAlerts
    End With
End Sub

Private Sub DeleteSampleSheets()
    Dim i As Long
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    For i = 1 To 5
        If SheetExists(SHEET_PREFIX & i) Then
            ThisWorkbook.Sheets(SHEET_PREFIX & i).Delete
        End If
    Next i
    Application.DisplayAlerts = True
End Sub

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
        Err.Raise ERR_INSUFFICIENT_DATA, "GetRandomRows", "Not enough rows in source data"
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

Private Function GetLastRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If Err.Number <> 0 Then GetLastRow = 0
End Function

Private Function GetLastColumn(ByVal ws As Worksheet) As Long
    On Error Resume Next
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Err.Number <> 0 Then GetLastColumn = 0
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function

Private Sub RaiseError(ByVal errorNumber As Long, ByVal errorMessage As String)
    Err.Raise errorNumber, "ProcessDataset", errorMessage
End Sub

Public Sub StopProcess()
    On Error Resume Next
    If SheetExists("Control") Then
        ThisWorkbook.Sheets("Control").Range("B1").Value = "Stopped"
    End If
End Sub
