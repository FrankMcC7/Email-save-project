Option Explicit

' Required References:
' - Microsoft Excel Object Library
' - Microsoft Office Object Library
' - Microsoft Scripting Runtime
' - Microsoft Visual Basic For Applications Extensibility

' Configuration Constants
Private Const SAMPLE_SIZE As Long = 100
Private Const MIN_SAMPLE_SHEETS As Long = 5
Private Const MAX_SAMPLE_SHEETS As Long = 15
Private Const MIN_REFRESH_INTERVAL As Long = 5  ' seconds
Private Const MAX_REFRESH_INTERVAL As Long = 30 ' seconds
Private Const LOG_FILE_PATH As String = "DataProcessing_Log.txt"
Private Const MAX_RETRIES As Integer = 3
Private Const CHUNK_SIZE As Long = 10000 ' For large dataset processing

' API Declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Custom Error Handling
Private Enum ProcessingError
    ERR_BASE = vbObjectError + 513
    ERR_NO_FILE_SELECTED = ERR_BASE + 1
    ERR_NO_REVIEW_STATUS = ERR_BASE + 2
    ERR_NO_APPROVED_DATA = ERR_BASE + 3
    ERR_INVALID_HEADERS = ERR_BASE + 4
    ERR_FILE_ACCESS = ERR_BASE + 5
    ERR_MEMORY_EXCEEDED = ERR_BASE + 6
    ERR_INVALID_DATA = ERR_BASE + 7
    ERR_SHEET_EXISTS = ERR_BASE + 8
    ERR_INVALID_SAMPLE_SIZE = ERR_BASE + 9
    ERR_PROCESSING_CANCELLED = ERR_BASE + 10
    ERR_INITIALIZATION_FAILED = ERR_BASE + 11
    ERR_VALIDATION_FAILED = ERR_BASE + 12
End Enum

' Type definition for processing statistics
Private Type ProcessingStats
    StartTime As Double
    EndTime As Double
    TotalRecords As Long
    ApprovedRecords As Long
    SamplesCreated As Long
    ErrorCount As Long
    ProcessingTime As Double
    LastError As String
End Type

' Global Variables
Private gStats As ProcessingStats
Private gIsProcessing As Boolean
Private gCancelled As Boolean

' Main Processing Procedure
Public Sub AutomatedDataProcessing()
    If Not InitializeEnvironment Then
        MsgBox "Failed to initialize environment", vbCritical
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' Validate configuration
    If Not ValidateConfiguration Then
        Err.Raise ProcessingError.ERR_VALIDATION_FAILED, "AutomatedDataProcessing", _
                  "Configuration validation failed"
    End If
    
    ' File selection and validation
    Dim rawFilePath As String
    rawFilePath = SelectAndValidateInputFile
    If rawFilePath = "" Then
        Err.Raise ProcessingError.ERR_NO_FILE_SELECTED, "AutomatedDataProcessing", _
                  "No file was selected or file is invalid"
    End If
    
    ' Import and process data
    Dim wsRaw As Worksheet
    Set wsRaw = ImportData(rawFilePath)
    
    If Not ValidateWorksheet(wsRaw) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "AutomatedDataProcessing", _
                  "Invalid worksheet or no data"
    End If
    
    ' Process raw data
    Dim wsApproved As Worksheet
    Set wsApproved = ProcessRawDataInChunks(wsRaw)
    
    ' Begin sampling process
    gIsProcessing = True
    Do While gIsProcessing And Not gCancelled
        If Not ProcessSampleBatch(wsApproved) Then Exit Do
        If Not HandleUserInteraction Then Exit Do
        Call ProcessingDelay
    Loop
    
    ' Display results
    If Not gCancelled Then
        DisplayProcessingResults
    End If
    
Cleanup:
    Call CleanupEnvironment
    Exit Sub

ErrorHandler:
    gStats.LastError = Err.Description
    gStats.ErrorCount = gStats.ErrorCount + 1
    
    Dim errMsg As String
    errMsg = HandleError(Err.Number, Err.Description, Err.Source)
    MsgBox errMsg, vbCritical, "Processing Error"
    
    Call CleanupOnError
    Resume Cleanup
End Sub

' Initialization and Configuration
Private Function InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize global variables
    gCancelled = False
    gIsProcessing = False
    
    ' Reset statistics
    With gStats
        .StartTime = Timer
        .TotalRecords = 0
        .ApprovedRecords = 0
        .SamplesCreated = 0
        .ErrorCount = 0
        .LastError = ""
    End With
    
    ' Initialize Excel environment
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    UpdateStatus "Initializing data processing..."
    InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    InitializeEnvironment = False
End Function

Private Function ValidateConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate sample sizes
    If SAMPLE_SIZE <= 0 Then
        gStats.LastError = "Invalid SAMPLE_SIZE configuration"
        Exit Function
    End If
    
    ' Validate sheet ranges
    If MIN_SAMPLE_SHEETS > MAX_SAMPLE_SHEETS Then
        gStats.LastError = "Invalid sample sheet range configuration"
        Exit Function
    End If
    
    ' Validate refresh intervals
    If MIN_REFRESH_INTERVAL > MAX_REFRESH_INTERVAL Then
        gStats.LastError = "Invalid refresh interval configuration"
        Exit Function
    End If
    
    ValidateConfiguration = True
    Exit Function
    
ErrorHandler:
    gStats.LastError = "Configuration validation error: " & Err.Description
    ValidateConfiguration = False
End Function

' File Selection and Import
Private Function SelectAndValidateInputFile() As String
    Dim fDialog As FileDialog
    Dim selectedFile As String
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select Data File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        
        If .Show Then
            selectedFile = .SelectedItems(1)
            If ValidateFileAccess(selectedFile) Then
                SelectAndValidateInputFile = selectedFile
            End If
        End If
    End With
End Function

Private Function ValidateFileAccess(ByVal filePath As String) As Boolean
    On Error Resume Next
    Dim fileNum As Integer
    Dim fileExists As Boolean
    
    ' Check if file exists
    fileExists = Dir(filePath) <> ""
    If Not fileExists Then
        ValidateFileAccess = False
        Exit Function
    End If
    
    ' Try to open file
    fileNum = FreeFile
    Open filePath For Input Access Read As #fileNum
    If Err.Number = 0 Then
        Close #fileNum
        ValidateFileAccess = True
    Else
        ValidateFileAccess = False
    End If
    
    On Error GoTo 0
End Function

Private Function ImportData(ByVal filePath As String) As Worksheet
    Dim fileExt As String
    Dim wsImport As Worksheet
    Dim retryCount As Integer
    
    fileExt = LCase(Right(filePath, 4))
    retryCount = 0
    
    Do While retryCount < MAX_RETRIES
        On Error Resume Next
        Select Case fileExt
            Case ".csv"
                Set wsImport = ImportCSVWithProgress(filePath)
            Case Else
                Set wsImport = ImportExcelWithProgress(filePath)
        End Select
        
        If Err.Number = 0 Then
            Exit Do
        Else
            retryCount = retryCount + 1
            If retryCount = MAX_RETRIES Then
                Err.Raise ProcessingError.ERR_FILE_ACCESS, "ImportData", _
                        "Failed to import file after " & MAX_RETRIES & " attempts"
            End If
            Application.Wait Now + TimeSerial(0, 0, 2)  ' Wait 2 seconds before retry
        End If
        On Error GoTo 0
    Loop
    
    Set ImportData = wsImport
End Function

' Data Validation Functions
Private Function ValidateWorksheet(ByRef ws As Worksheet) As Boolean
    If ws Is Nothing Then
        ValidateWorksheet = False
        Exit Function
    End If
    
    ' Check if worksheet has any data
    If ws.UsedRange.Rows.Count <= 1 Then
        ValidateWorksheet = False
        Exit Function
    End If
    
    ' Validate headers
    If Not ValidateHeaders(ws) Then
        ValidateWorksheet = False
        Exit Function
    End If
    
    ValidateWorksheet = True
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

Private Function ValidateSampleSize(ByVal totalRows As Long) As Boolean
    If SAMPLE_SIZE <= 0 Or SAMPLE_SIZE > totalRows Then
        ValidateSampleSize = False
        Exit Function
    End If
    
    ValidateSampleSize = True
End Function

' Sheet Management Functions
Private Function CreateApprovedSheet() As Worksheet
    Dim wsApproved As Worksheet
    
    ' Delete existing sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("ApprovedData").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new sheet
    Set wsApproved = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsApproved.Name = "ApprovedData"
    
    ' Initialize the header row with proper formatting
    With wsApproved.Range("A1").Font
        .Bold = True
        .Size = 11
    End With
    
    ' Format as table
    With wsApproved
        .Rows(1).Interior.Color = RGB(217, 217, 217)
        .Cells.Clear
    End With
    
    Set CreateApprovedSheet = wsApproved
End Function

Private Function CreateSampleSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    ' Delete existing sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new sheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    Set CreateSampleSheet = ws
End Function

' Data Processing Functions
Private Function ProcessRawDataInChunks(ByRef wsRaw As Worksheet) As Worksheet
    If Not ValidateWorksheet(wsRaw) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "ProcessRawDataInChunks", _
                  "Invalid worksheet or no data"
    End If
    
    Dim wsApproved As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim chunkStart As Long, chunkEnd As Long
    Dim reviewStatusCol As Long
    
    ' Find Review Status column
    lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    reviewStatusCol = GetColumnByHeaderName(wsRaw, "Review Status")
    
    If reviewStatusCol = 0 Then
        Err.Raise ProcessingError.ERR_NO_REVIEW_STATUS, "ProcessRawDataInChunks", _
                  "Review Status column not found"
    End If
    
    ' Create approved data worksheet
    Set wsApproved = CreateApprovedSheet
    
    ' Process data in chunks
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    chunkStart = 2  ' Start after header row
    
    Do While chunkStart <= lastRow And Not gCancelled
        chunkEnd = Application.Min(chunkStart + CHUNK_SIZE - 1, lastRow)
        
        UpdateStatus "Processing rows " & chunkStart & " to " & chunkEnd
        
        ' Process chunk
        ProcessDataChunk wsRaw, wsApproved, chunkStart, chunkEnd, reviewStatusCol
        
        chunkStart = chunkEnd + 1
        DoEvents
        
        If Not gIsProcessing Then Exit Do
    Loop
    
    Set ProcessRawDataInChunks = wsApproved
End Function

Private Sub ProcessDataChunk(ByRef wsSource As Worksheet, ByRef wsTarget As Worksheet, _
                           ByVal startRow As Long, ByVal endRow As Long, _
                           ByVal statusCol As Long)
    Dim dataRange As Range
    Dim approvedRange As Range
    Dim cell As Range
    
    Set approvedRange = Nothing  ' Initialize to Nothing
    
    ' Get chunk range
    Set dataRange = wsSource.Range(wsSource.Cells(startRow, 1), _
                                 wsSource.Cells(endRow, wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column))
    
    ' Filter for approved records
    For Each cell In dataRange.Columns(statusCol).Cells
        If UCase(Trim(cell.Value)) = "APPROVED" Then
            If approvedRange Is Nothing Then
                Set approvedRange = wsSource.Rows(cell.Row)
            Else
                Set approvedRange = Union(approvedRange, wsSource.Rows(cell.Row))
            End If
        End If
    Next cell
    
    ' Copy approved records if any found
    If Not approvedRange Is Nothing Then
        approvedRange.Copy Destination:=wsTarget.Cells(wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1, 1)
        gStats.ApprovedRecords = gStats.ApprovedRecords + approvedRange.Rows.Count
    End If
    
    gStats.TotalRecords = gStats.TotalRecords + (endRow - startRow + 1)
End Sub

' Sample Generation Functions
Private Function ProcessSampleBatch(ByRef wsSource As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate source worksheet
    If Not ValidateWorksheet(wsSource) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "ProcessSampleBatch", _
                  "Invalid source worksheet"
    End If
    
    Dim sampleSheets As Collection
    Set sampleSheets = CreateSampleSheets
    
    If sampleSheets.Count > 0 Then
        CreateSamples wsSource, sampleSheets
        ProcessSampleBatch = True
    Else
        ProcessSampleBatch = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogError Err.Number, "ProcessSampleBatch: " & Err.Description
    ProcessSampleBatch = False
End Function

Private Function CreateSampleSheets() As Collection
    Dim sampleSheets As Collection
    Dim sampleSheet As Worksheet
    Dim numSheets As Long
    Dim i As Long
    
    Set sampleSheets = New Collection
    
    ' Randomly determine number of sample sheets
    Randomize
    numSheets = Int((MAX_SAMPLE_SHEETS - MIN_SAMPLE_SHEETS + 1) * Rnd + MIN_SAMPLE_SHEETS)
    
    ' Create new sample sheets
    For i = 1 To numSheets
        Set sampleSheet = CreateSampleSheet("Sample" & i)
        If Not sampleSheet Is Nothing Then
            sampleSheets.Add sampleSheet
        End If
    Next i
    
    Set CreateSampleSheets = sampleSheets
End Function

Private Sub CreateSamples(ByRef wsSource As Worksheet, ByRef sampleSheets As Collection)
    Dim sampleSheet As Worksheet
    Dim headerRange As Range
    Dim lastRow As Long, lastCol As Long
    
    ' Get source data range
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set headerRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol))
    
    ' Validate sample size
    If Not ValidateSampleSize(lastRow - 1) Then ' Subtract 1 for header row
        Err.Raise ProcessingError.ERR_INVALID_SAMPLE_SIZE, "CreateSamples", _
                  "Invalid sample size for the given data set"
    End If
    
    ' Process each sample sheet
    For Each sampleSheet In sampleSheets
        UpdateStatus "Creating sample in " & sampleSheet.Name
        
        ' Copy headers
        headerRange.Copy Destination:=sampleSheet.Range("A1")
        
        ' Generate and copy random sample
        GenerateRandomSample wsSource, sampleSheet, lastRow, lastCol
        
        ' Format sample sheet
        FormatSampleSheet sampleSheet
        
        gStats.SamplesCreated = gStats.SamplesCreated + 1
    Next sampleSheet
End Sub

Private Sub GenerateRandomSample(ByRef wsSource As Worksheet, ByRef wsTarget As Worksheet, _
                               ByVal sourceRows As Long, ByVal sourceCols As Long)
    Dim randomRows() As Long
    Dim i As Long, j As Long
    Dim targetRow As Long
    Dim sourceRow As Long
    Dim dataRange As Range
    
    ' Generate random row indices
    ReDim randomRows(1 To Application.Min(SAMPLE_SIZE, sourceRows - 1))
    Call GenerateUniqueRandomNumbers randomRows, 2, sourceRows ' Start from 2 to skip header
    
    ' Optimize copying by using arrays
    Dim sourceData As Variant
    Dim targetData As Variant
    
    ' Get source data
    Set dataRange = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(sourceRows, sourceCols))
    sourceData = dataRange.Value
    
    ' Prepare target array
    ReDim targetData(1 To UBound(randomRows), 1 To sourceCols)
    
    ' Fill target array with random samples
    For i = 1 To UBound(randomRows)
        For j = 1 To sourceCols
            targetData(i, j) = sourceData(randomRows(i) - 1, j)
        Next j
    Next i
    
    ' Write to target worksheet
    wsTarget.Range("A2").Resize(UBound(randomRows), sourceCols) = targetData
End Sub

Private Sub GenerateUniqueRandomNumbers(ByRef numbers() As Long, _
                                      ByVal minValue As Long, _
                                      ByVal maxValue As Long)
    Dim i As Long, j As Long
    Dim temp As Long
    Dim numCount As Long
    
    numCount = UBound(numbers) - LBound(numbers) + 1
    
    ' Validate input
    If numCount > (maxValue - minValue + 1) Then
        Err.Raise ProcessingError.ERR_INVALID_SAMPLE_SIZE, "GenerateUniqueRandomNumbers", _
                  "Sample size larger than available data range"
    End If
    
    ' Initialize array with sequential numbers
    For i = LBound(numbers) To UBound(numbers)
        numbers(i) = minValue + i - LBound(numbers)
    Next i
    
    ' Fisher-Yates shuffle
    Randomize
    For i = UBound(numbers) To LBound(numbers) + 1 Step -1
        j = Int((i - LBound(numbers) + 1) * Rnd + LBound(numbers))
        temp = numbers(i)
        numbers(i) = numbers(j)
        numbers(j) = temp
    Next i
End Sub

Private Sub FormatSampleSheet(ByRef ws As Worksheet)
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Get data range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 1 Then ' Only format if there's data
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        
        ' Format as table
        With dataRange
            ' Borders
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            ' Format headers
            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(217, 217, 217)
        End With
        
        ' Autofit columns
        ws.Cells.EntireColumn.AutoFit
    End If
End Sub

' Utility Functions
Private Function HandleUserInteraction() As Boolean
    ' Check for user cancellation (ESC key)
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then
        If MsgBox("Do you want to stop processing?", vbQuestion + vbYesNo) = vbYes Then
            gCancelled = True
            HandleUserInteraction = False
            Exit Function
        End If
    End If
    
    HandleUserInteraction = True
End Function

Private Sub ProcessingDelay()
    Dim delaySeconds As Long
    Dim startTime As Double
    
    ' Generate random delay
    Randomize
    delaySeconds = Int((MAX_REFRESH_INTERVAL - MIN_REFRESH_INTERVAL + 1) * Rnd + MIN_REFRESH_INTERVAL)
    
    UpdateStatus "Waiting " & delaySeconds & " seconds before next iteration..."
    
    startTime = Timer
    Do While Timer < startTime + delaySeconds
        DoEvents
        If Not HandleUserInteraction Then Exit Do
    Loop
End Sub

Private Sub DisplayProcessingResults()
    Dim msg As String
    
    gStats.EndTime = Timer
    gStats.ProcessingTime = gStats.EndTime - gStats.StartTime
    
    msg = "Processing completed successfully." & vbNewLine & vbNewLine & _
          "Total records processed: " & FormatNumber(gStats.TotalRecords) & vbNewLine & _
          "Approved records: " & FormatNumber(gStats.ApprovedRecords) & vbNewLine & _
          "Samples created: " & FormatNumber(gStats.SamplesCreated) & vbNewLine & _
          "Processing time: " & Format(gStats.ProcessingTime, "0.00") & " seconds" & vbNewLine & _
          "Error count: " & FormatNumber(gStats.ErrorCount)
    
    If gStats.ErrorCount > 0 Then
        msg = msg & vbNewLine & vbNewLine & "Last error: " & gStats.LastError
    End If
    
    MsgBox msg, vbInformation, "Processing Results"
End Sub

Private Function GetColumnByHeaderName(ByRef ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim col As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If StrComp(ws.Cells(1, col).Value, headerName, vbTextCompare) = 0 Then
            GetColumnByHeaderName = col
            Exit Function
        End If
    Next col
    
    GetColumnByHeaderName = 0
End Function

Private Function GetUniqueSheetName(ByVal baseName As String) As String
    Dim counter As Long
    Dim proposedName As String
    
    counter = 1
    proposedName = baseName
    
    Do While WorksheetExists(proposedName)
        counter = counter + 1
        proposedName = baseName & counter
    Loop
    
    GetUniqueSheetName = proposedName
End Function

Private Function WorksheetExists(ByVal wsName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    
    WorksheetExists = Not ws Is Nothing
End Function

Private Sub UpdateStatus(ByVal statusText As String)
    Application.StatusBar = statusText
    DoEvents
End Sub

Private Function FormatNumber(ByVal number As Double) As String
    FormatNumber = Format(number, "#,##0.00")
End Function

Private Sub CleanupOnError()
    ' Reset application states
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    
    ' Delete temporary sheets
    CleanupWorksheets
    
    ' Reset global variables
    gIsProcessing = False
    gCancelled = False
End Sub
