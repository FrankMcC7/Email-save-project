Option Explicit

' Configuration Constants
Private Const SAMPLE_SIZE As Long = 100
Private Const MIN_SAMPLE_SHEETS As Long = 5
Private Const MAX_SAMPLE_SHEETS As Long = 15
Private Const MIN_REFRESH_INTERVAL As Long = 5  ' seconds
Private Const MAX_REFRESH_INTERVAL As Long = 30 ' seconds
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
    ERR_NO_FUND_GCI = vbObjectError + 6
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
    Dim wsRaw As Worksheet
    Dim wsApproved As Worksheet
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
        
        ' Random delay between iterations
        Call ProcessingDelay
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
        Case ProcessingError.ERR_NO_FUND_GCI
            MsgBox "Fund GCI column not found in the data.", vbCritical
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
    Dim fundGciCol As Long
    Dim rng As Range
    
    ' Find Fund GCI column first
    fundGciCol = Application.Match("Fund GCI", wsRaw.Rows(1), 0)
    If IsError(fundGciCol) Then
        Err.Raise ProcessingError.ERR_NO_FUND_GCI, "ProcessRawData", "Fund GCI column not found"
    End If
    
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
        
        ' Generate random sample based on Fund GCI
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
    Dim numSheets As Long
    
    Set sampleSheets = New Collection
    
    ' Randomly determine number of sample sheets
    Randomize
    numSheets = Int((MAX_SAMPLE_SHEETS - MIN_SAMPLE_SHEETS + 1) * Rnd + MIN_SAMPLE_SHEETS)
    
    ' Delete existing sample sheets
    Application.DisplayAlerts = False
    For i = 1 To MAX_SAMPLE_SHEETS
        On Error Resume Next
        ThisWorkbook.Worksheets("Sample" & i).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
    
    ' Create new sample sheets
    For i = 1 To numSheets
        Set sampleSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sampleSheet.Name = "Sample" & i
        sampleSheets.Add sampleSheet
    Next i
    
    Set PrepareSampleSheets = sampleSheets
End Function

Private Sub GenerateRandomSample(ByRef wsSource As Worksheet, ByRef dataArray() As Variant)
    Dim sourceData As Variant
    Dim totalRows As Long
    Dim fundGciCol As Long
    Dim fundGciValues() As Double
    Dim sortedIndices() As Long
    Dim i As Long, j As Long
    Dim temp As Double
    Dim tempIndex As Long
    
    ' Find Fund GCI column
    fundGciCol = Application.Match("Fund GCI", wsSource.Rows(1), 0)
    If IsError(fundGciCol) Then
        Err.Raise ProcessingError.ERR_NO_FUND_GCI, "GenerateRandomSample", "Fund GCI column not found"
    End If
    
    ' Get source data
    totalRows = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row - 1
    sourceData = wsSource.Range("A2").Resize(totalRows, UBound(dataArray, 2)).Value
    
    ' Create array of Fund GCI values with their indices
    ReDim fundGciValues(1 To totalRows)
    ReDim sortedIndices(1 To totalRows)
    
    ' Fill arrays with values and indices
    For i = 1 To totalRows
        fundGciValues(i) = CDbl(sourceData(i, fundGciCol))
        sortedIndices(i) = i
    Next i
    
    ' Sort arrays by Fund GCI value (bubble sort)
    For i = 1 To totalRows - 1
        For j = 1 To totalRows - i
            If fundGciValues(j) < fundGciValues(j + 1) Then
                ' Swap values
                temp = fundGciValues(j)
                fundGciValues(j) = fundGciValues(j + 1)
                fundGciValues(j + 1) = temp
                
                ' Swap indices
                tempIndex = sortedIndices(j)
                sortedIndices(j) = sortedIndices(j + 1)
                sortedIndices(j + 1) = tempIndex
            End If
        Next j
    Next i
    
    ' Take top 100 (or less if not enough rows) and randomize their order
    Dim sampleSize As Long
    sampleSize = Application.WorksheetFunction.Min(SAMPLE_SIZE, totalRows)
    
    ' Shuffle the top entries
    For i = 1 To sampleSize
        j = Int((sampleSize - i + 1) * Rnd + i)
        tempIndex = sortedIndices(i)
        sortedIndices(i) = sortedIndices(j)
        sortedIndices(j) = tempIndex
    Next i
    
    ' Fill sample array with randomly ordered top Fund GCI rows
    For i = 1 To sampleSize
        Call CopyArrayRow(sourceData, dataArray, sortedIndices(i), i)
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

Private Sub ProcessingDelay()
    Dim startTime As Double
    Dim delaySeconds As Long
    
    ' Generate random delay between MIN_REFRESH_INTERVAL and MAX_REFRESH_INTERVAL
    Randomize
    delaySeconds = Int((MAX_REFRESH_INTERVAL - MIN_REFRESH_INTERVAL + 1) * Rnd + MIN_REFRESH_INTERVAL)
    
    ' Update status to show wait time
    UpdateStatus "Waiting for " & delaySeconds & " seconds before next iteration..."
    
    startTime = Timer
    Do While Timer < startTime + delaySeconds
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
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Sample" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Private Function FormatNumber(ByVal number As Long) As String
    FormatNumber = Format(number, "#,##0")
End Function

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

Private Function WorksheetExists(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

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
