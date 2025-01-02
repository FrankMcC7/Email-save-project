Option Explicit

' Configuration Constants
Private Const SAMPLE_SIZE As Long = 100
Private Const SAMPLE_SHEETS As Long = 5
Private Const REFRESH_INTERVAL As Long = 5 ' seconds
Private Const LOG_FILE_PATH As String = "DataProcessing_Log.txt"

' API Declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Custom Error Handling
Private Enum ProcessingError
    ERR_NO_FILE_SELECTED = vbObjectError + 1
    ERR_NO_REVIEW_STATUS = vbObjectError + 2
    ERR_NO_APPROVED_DATA = vbObjectError + 3
    ERR_INVALID_HEADERS = vbObjectError + 4
    ERR_FILE_ACCESS = vbObjectError + 5
End Enum

' Type definition for processing statistics
Private Type ProcessingStats
    StartTime As Double
    TotalRecords As Long
    ApprovedRecords As Long
    SamplesCreated As Long
    ErrorCount As Long
End Type

Sub AutomatedDataProcessing()
    Dim stats As ProcessingStats
    Dim rawFilePath As String
    Dim wbRaw As Workbook
    Dim wsRaw As Worksheet
    Dim wsApproved As Worksheet
    Dim headerRange As Range
    Dim reviewStatusCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim i As Long, j As Long
    Dim sampleSheets As Collection
    Dim sampleSheet As Worksheet
    Dim approvedDataRows As Long
    Dim isRunning As Boolean
    
    ' Initialize statistics
    stats.StartTime = Timer
    
    ' Initialize application settings
    Call InitializeApplication
    
    On Error GoTo ErrorHandler
    
    ' File selection and import
    rawFilePath = SelectInputFile
    If rawFilePath = "" Then
        Err.Raise ProcessingError.ERR_NO_FILE_SELECTED, "AutomatedDataProcessing", "No file selected"
    End If
    
    ' Import data based on file type
    Set wsRaw = ImportData(rawFilePath)
    
    ' Initial data cleanup
    Call CleanupRawData(wsRaw)
    
    ' Process and filter data
    Set wsApproved = ProcessRawData(wsRaw, stats)
    
    ' Begin sampling loop
    isRunning = True
    Do While isRunning
        ' Create and process samples
        Call CreateSamples(wsApproved, stats)
        
        ' Check for exit condition
        isRunning = CheckContinueProcessing
        
        ' Update status
        Call UpdateStatus("Processing complete. Samples created: " & stats.SamplesCreated)
        
        ' Pause between iterations
        Call ProcessingDelay(REFRESH_INTERVAL)
    Loop
    
    MsgBox "Processing completed successfully." & vbNewLine & _
           "Total records processed: " & stats.TotalRecords & vbNewLine & _
           "Approved records: " & stats.ApprovedRecords & vbNewLine & _
           "Samples created: " & stats.SamplesCreated, vbInformation
    
Cleanup:
    Call CleanupApplication
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case ProcessingError.ERR_NO_FILE_SELECTED
            MsgBox "No file was selected. Operation cancelled.", vbExclamation
        Case ProcessingError.ERR_NO_REVIEW_STATUS
            MsgBox "Review Status column not found in the data.", vbCritical
        Case ProcessingError.ERR_NO_APPROVED_DATA
            MsgBox "No approved data found for processing.", vbExclamation
        Case ProcessingError.ERR_INVALID_HEADERS
            MsgBox "Invalid or missing headers in the data.", vbCritical
        Case ProcessingError.ERR_FILE_ACCESS
            MsgBox "Error accessing the file: " & Err.Description, vbCritical
        Case Else
            MsgBox "An unexpected error occurred: " & Err.Description, vbCritical
    End Select
    
    Call LogError(Err.Number, Err.Description)
    Resume Cleanup
End Sub

Private Function SelectInputFile() As String
    Dim fDialog As FileDialog
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select Data File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        
        If .Show Then
            SelectInputFile = .SelectedItems(1)
        End If
    End With
End Function

Private Function ImportData(ByVal filePath As String) As Worksheet
    Dim fileExt As String
    Dim ws As Worksheet
    
    fileExt = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))
    
    If fileExt = "csv" Then
        Set ImportData = ImportCSV(filePath)
    Else
        Set ImportData = ImportExcel(filePath)
    End If
End Function

Private Function ImportCSV(ByVal filePath As String) As Worksheet
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "RawData"
    
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1)
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    Set ImportCSV = ws
    Exit Function
    
ErrorHandler:
    Err.Raise ProcessingError.ERR_FILE_ACCESS, "ImportCSV", _
            "Error importing CSV file: " & Err.Description
End Function

Private Function ImportExcel(ByVal filePath As String) As Worksheet
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    Set wsTarget = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTarget.Name = "RawData"
    
    wbSource.Sheets(1).UsedRange.Copy Destination:=wsTarget.Range("A1")
    wbSource.Close SaveChanges:=False
    
    Set ImportExcel = wsTarget
    Exit Function
    
ErrorHandler:
    Err.Raise ProcessingError.ERR_FILE_ACCESS, "ImportExcel", _
            "Error importing Excel file: " & Err.Description
End Function

Private Sub CleanupRawData(ByRef ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    
    ' Remove header row
    ws.Rows(1).Delete
    
    ' Get data range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Remove blank rows
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Private Function ProcessRawData(ByRef wsRaw As Worksheet, ByRef stats As ProcessingStats) As Worksheet
    Dim wsApproved As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim reviewStatusCol As Long
    Dim rng As Range
    
    ' Find Review Status column
    lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    reviewStatusCol = Application.Match("Review Status", wsRaw.Rows(1), 0)
    
    If IsError(reviewStatusCol) Then
        Err.Raise ProcessingError.ERR_NO_REVIEW_STATUS
    End If
    
    ' Filter for approved records
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    stats.TotalRecords = lastRow - 1 ' Subtract header row
    
    With wsRaw
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).AutoFilter _
            Field:=reviewStatusCol, Criteria1:="Approved"
    End With
    
    ' Create approved data sheet
    Set wsApproved = CreateApprovedSheet(wsRaw, stats)
    
    ' Cleanup
    wsRaw.AutoFilterMode = False
    Application.DisplayAlerts = False
    wsRaw.Delete
    Application.DisplayAlerts = True
    
    Set ProcessRawData = wsApproved
End Function

Private Function CreateApprovedSheet(ByRef wsSource As Worksheet, ByRef stats As ProcessingStats) As Worksheet
    Dim wsApproved As Worksheet
    Dim rng As Range
    
    ' Create new sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("ApprovedData").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsApproved = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsApproved.Name = "ApprovedData"
    
    ' Copy filtered data
    Set rng = wsSource.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
    rng.Copy Destination:=wsApproved.Range("A1")
    
    ' Update statistics
    stats.ApprovedRecords = wsApproved.Cells(wsApproved.Rows.Count, 1).End(xlUp).Row - 1
    
    Set CreateApprovedSheet = wsApproved
End Function

Private Sub CreateSamples(ByRef wsSource As Worksheet, ByRef stats As ProcessingStats)
    Dim sampleSheets As Collection
    Dim sampleSheet As Worksheet
    Dim headerRange As Range
    Dim lastHeaderCol As Long
    Dim dataArray() As Variant
    Dim i As Long, j As Long
    
    ' Prepare sample sheets
    Set sampleSheets = PrepareSampleSheets
    
    ' Get header range
    lastHeaderCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set headerRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastHeaderCol))
    
    ' Create samples using arrays for better performance
    ReDim dataArray(1 To SAMPLE_SIZE, 1 To lastHeaderCol)
    
    For Each sampleSheet In sampleSheets
        ' Copy headers
        headerRange.Copy Destination:=sampleSheet.Range("A1")
        
        ' Generate random sample
        Call GenerateRandomSample(wsSource, dataArray)
        
        ' Write sample to sheet
        sampleSheet.Range("A2").Resize(SAMPLE_SIZE, lastHeaderCol) = dataArray
        
        stats.SamplesCreated = stats.SamplesCreated + 1
    Next sampleSheet
End Sub

Private Function PrepareSampleSheets() As Collection
    Dim sampleSheets As Collection
    Dim sampleSheet As Worksheet
    Dim i As Long
    
    Set sampleSheets = New Collection
    
    ' Delete existing sample sheets
    Application.DisplayAlerts = False
    For i = 1 To SAMPLE_SHEETS
        On Error Resume Next
        ThisWorkbook.Worksheets("Sample" & i).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    
    ' Create new sample sheets
    For i = 1 To SAMPLE_SHEETS
        Set sampleSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sampleSheet.Name = "Sample" & i
        sampleSheets.Add sampleSheet
    Next i
    
    Set PrepareSampleSheets = sampleSheets
End Function

Private Sub GenerateRandomSample(ByRef wsSource As Worksheet, ByRef dataArray() As Variant)
    Dim sourceData As Variant
    Dim totalRows As Long
    Dim randomIndices() As Long
    Dim i As Long
    
    ' Get source data
    totalRows = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row - 1
    sourceData = wsSource.Range("A2").Resize(totalRows, UBound(dataArray, 2)).Value
    
    ' Generate random indices
    ReDim randomIndices(1 To Application.WorksheetFunction.Min(SAMPLE_SIZE, totalRows))
    Call GenerateRandomIndices(randomIndices, totalRows)
    
    ' Fill sample array
    For i = 1 To UBound(randomIndices)
        Call CopyArrayRow(sourceData, dataArray, randomIndices(i), i)
    Next i
End Sub

Private Sub GenerateRandomIndices(ByRef indices() As Long, ByVal maxValue As Long)
    Dim i As Long
    Dim temp As Long
    Dim j As Long
    
    ' Initialize array with sequential numbers
    For i = 1 To UBound(indices)
        indices(i) = i
    Next i
    
    ' Shuffle using Fisher-Yates algorithm
    Randomize
    For i = UBound(indices) To 2 Step -1
        j = Int((i * Rnd) + 1)
        temp = indices(i)
        indices(i) = indices(j)
        indices(j) = temp
    Next i
End Sub

Private Sub CopyArrayRow(ByRef source As Variant, ByRef target() As Variant, _
                        ByVal sourceRow As Long, ByVal targetRow As Long)
    Dim col As Long
    
    For col = 1 To UBound(target, 2)
        target(targetRow, col) = source(sourceRow, col)
    Next col
End Sub

Private Function CheckContinueProcessing() As Boolean
    CheckContinueProcessing = (GetAsyncKeyState(vbKeyEscape) = 0)
End Function

Private Sub ProcessingDelay(ByVal seconds As Long)
    Dim startTime As Double
    startTime = Timer
    
    Do While Timer < startTime + seconds
        DoEvents
        If Not CheckContinueProcessing Then Exit Sub
    Loop
End Sub

Private Sub UpdateStatus(ByVal statusText As String)
    Application.StatusBar = statusText
End Sub

Private Sub LogError(ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim fileNum As Integer
    Dim logMessage As String
    
    ' Create log message
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & _
                 " - Error " & errorNumber & ": " & errorDescription
    
    ' Write to log file
    fileNum = FreeFile
    Open LOG_FILE_PATH For Append As #fileNum
    Print #fileNum, logMessage
    Close #fileNum
End Sub

Private Sub InitializeApplication()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    UpdateStatus "Initializing data processing..."
End Sub

Private Sub CleanupApplication()
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
    ' Clean up any leftover sample sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Sample" Then
            ws.Delete
        End If
    Next ws
End Sub

' Helper function to format numbers with commas
Private Function FormatNumber(ByVal number As Long) As String
    FormatNumber = Format(number, "#,##0")
End Function

' Helper function to validate headers
Private Function ValidateHeaders(ByRef ws As Worksheet) As Boolean
    Dim lastCol As Long
    Dim cell As Range
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Check for empty headers
    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If Len(Trim(cell.Value)) = 0 Then
            ValidateHeaders = False
            Exit Function
        End If
    Next cell
    
    ValidateHeaders = True
End Function

' Helper function to check if a worksheet exists
Private Function WorksheetExists(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

' Helper function to get a unique sheet name
Private Function GetUniqueSheetName(ByVal baseName As String) As String
    Dim counter As Long
    Dim proposedName As String
    
    counter = 1
    proposedName = baseName
    
    Do While WorksheetExists(proposedName)
        counter = counter + 1
        proposedName = baseName & "_" & counter
    Loop
    
    GetUniqueSheetName = proposedName
End Function
