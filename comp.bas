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
    ERR_BASE = vbObjectError + 513  ' Explicit base for custom errors
    ERR_NO_FILE_SELECTED = ERR_BASE + 1
    ERR_NO_REVIEW_STATUS = ERR_BASE + 2
    ERR_NO_APPROVED_DATA = ERR_BASE + 3
    ERR_INVALID_HEADERS = ERR_BASE + 4
    ERR_FILE_ACCESS = ERR_BASE + 5
    ERR_MEMORY_EXCEEDED = ERR_BASE + 6
    ERR_INVALID_DATA = ERR_BASE + 7
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
End Type

' Global Variables
Private gStats As ProcessingStats
Private gIsProcessing As Boolean

Public Sub AutomatedDataProcessing()
    Dim rawFilePath As String
    Dim wsRaw As Worksheet
    Dim wsApproved As Worksheet
    
    ' Initialize statistics and application
    Call InitializeEnvironment
    
    On Error GoTo ErrorHandler
    
    ' File selection and validation
    rawFilePath = SelectAndValidateInputFile
    If rawFilePath = "" Then
        Err.Raise ProcessingError.ERR_NO_FILE_SELECTED, "AutomatedDataProcessing", _
                  "No file was selected or file is invalid"
    End If
    
    ' Import and process data
    Set wsRaw = ImportData(rawFilePath)
    If Not ValidateDataStructure(wsRaw) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "AutomatedDataProcessing", _
                  "Data structure validation failed"
    End If
    
    ' Process raw data in chunks
    Set wsApproved = ProcessRawDataInChunks(wsRaw)
    
    ' Begin sampling process
    gIsProcessing = True
    Do While gIsProcessing
        If Not ProcessSampleBatch(wsApproved) Then Exit Do
        If Not HandleUserInteraction Then Exit Do
        Call ProcessingDelay
    Loop
    
    ' Display results
    DisplayProcessingResults
    
Cleanup:
    Call CleanupEnvironment
    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = HandleError(Err.Number, Err.Description, Err.Source)
    MsgBox errMsg, vbCritical, "Processing Error"
    Resume Cleanup
End Sub

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

Private Function ImportCSVWithProgress(ByVal filePath As String) As Worksheet
    Dim ws As Worksheet
    Dim qt As QueryTable
    
    ' Create new worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = GetUniqueSheetName("RawData")
    
    ' Import CSV with progress
    UpdateStatus "Importing CSV file..."
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
    
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(2)  ' xlTextFormat
        .RefreshStyle = xlOverwriteCells
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveFormatting = True
        .Refresh BackgroundQuery:=False
    End With
    
    Set ImportCSVWithProgress = ws
End Function

Private Function ImportExcelWithProgress(ByVal filePath As String) As Worksheet
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim lastRow As Long, lastCol As Long
    
    UpdateStatus "Opening Excel file..."
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)
    
    ' Create new worksheet
    Set wsTarget = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsTarget.Name = GetUniqueSheetName("RawData")
    
    ' Copy data in chunks
    With wbSource.Sheets(1)
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        Dim startRow As Long, endRow As Long
        startRow = 1
        
        Do While startRow <= lastRow
            UpdateStatus "Copying rows " & startRow & " to " & Application.Min(startRow + CHUNK_SIZE - 1, lastRow)
            endRow = Application.Min(startRow + CHUNK_SIZE - 1, lastRow)
            
            .Range(.Cells(startRow, 1), .Cells(endRow, lastCol)).Copy _
                Destination:=wsTarget.Cells(startRow, 1)
            
            startRow = endRow + 1
            DoEvents
        Loop
    End With
    
    wbSource.Close SaveChanges:=False
    Set ImportExcelWithProgress = wsTarget
End Function

Private Function ProcessRawDataInChunks(ByRef wsRaw As Worksheet) As Worksheet
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
    
    Do While chunkStart <= lastRow
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
End Function

Private Function ProcessSampleBatch(ByRef wsSource As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
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

' ... [Additional supporting functions would go here] ...

' Helper Functions for Error Handling and Logging
Private Function HandleError(ByVal errNumber As Long, ByVal errDescription As String, _
                           Optional ByVal errSource As String = "") As String
    Dim errorMsg As String
    
    Select Case errNumber
        Case ProcessingError.ERR_NO_FILE_SELECTED
            errorMsg = "No file was selected or the file is invalid."
        Case ProcessingError.ERR_NO_REVIEW_STATUS
            errorMsg = "Review Status column not found in the data."
        Case ProcessingError.ERR_NO_APPROVED_DATA
            errorMsg = "No approved data found for processing."
        Case ProcessingError.ERR_INVALID_HEADERS
            errorMsg = "Invalid or missing headers in the data."
        Case ProcessingError.ERR_FILE_ACCESS
            errorMsg = "Error accessing the file: " & errDescription
        Case ProcessingError.ERR_MEMORY_EXCEEDED
            errorMsg = "Memory limit exceeded. Try processing a smaller dataset."
        Case ProcessingError.ERR_INVALID_DATA
            errorMsg = "Invalid data structure: " & errDescription
        Case Else
            errorMsg = "An unexpected error occurred: " & errDescription
    End Select
    
    ' Log error
    LogError errNumber, errorMsg & IIf(errSource <> "", " (Source: " & errSource & ")", "")
    
    HandleError = errorMsg
End Function

Private Sub LogError(ByVal errorNumber As Long, ByVal errorDescription As String)
    On Error Resume Next
    Dim fileNum As Integer
    Dim logMessage As String
    
    logMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & _
                 " - Error " & errorNumber & ": " & errorDescription
    
    fileNum = FreeFile
    Open LOG_FILE_PATH For Append As #fileNum
    If Err.Number = 0 Then
        Print #fileNum, logMessage
        Close #fileNum
    End If
    
    Debug.Print logMessage  ' Also output to immediate window
End Sub

Private Sub InitializeEnvironment()
    ' Reset statistics
    With gStats
        .StartTime = Timer
        .TotalRecords = 0
        .ApprovedRecords = 0
        .SamplesCreated = 0
        .ErrorCount = 0
    End With
    
    ' Initialize Excel environment
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    ' Clear status bar
    UpdateStatus "Initializing data processing..."
End Sub

Private Sub CleanupEnvironment()
    ' Update statistics
    gStats.EndTime = Timer
    gStats.ProcessingTime = gStats.EndTime - gStats.StartTime
    
    ' Restore Excel environment
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    
    ' Clean up temporary worksheets
    CleanupWorksheets
End Sub

Private Sub CleanupWorksheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Sample" Or _
           Left(ws.Name, 7) = "RawData" Or _
           ws.Name = "ApprovedData" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

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
        Set sampleSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sampleSheet.Name = GetUniqueSheetName("Sample" & i)
        sampleSheets.Add sampleSheet
    Next i
    
    Set CreateSampleSheets = sampleSheets
End Function

Private Sub CreateSamples(ByRef wsSource As Worksheet, ByRef sampleSheets As Collection)
    Dim sampleSheet As Worksheet
    Dim headerRange As Range
    Dim dataRange As Range
    Dim lastRow As Long, lastCol As Long
    
    ' Get source data range
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set headerRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol))
    
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
    
    ' Generate random row indices
    ReDim randomRows(1 To Application.Min(SAMPLE_SIZE, sourceRows - 1))
    Call GenerateUniqueRandomNumbers randomRows, 2, sourceRows
    
    ' Copy random rows
    targetRow = 2 ' Start after header
    For i = LBound(randomRows) To UBound(randomRows)
        sourceRow = randomRows(i)
        wsSource.Range(wsSource.Cells(sourceRow, 1), wsSource.Cells(sourceRow, sourceCols)).Copy _
            Destination:=wsTarget.Cells(targetRow, 1)
        targetRow = targetRow + 1
    Next i
End Sub

Private Sub GenerateUniqueRandomNumbers(ByRef numbers() As Long, _
                                      ByVal minValue As Long, _
                                      ByVal maxValue As Long)
    Dim i As Long, j As Long
    Dim temp As Long
    Dim count As Long
    
    count = UBound(numbers) - LBound(numbers) + 1
    
    ' Initialize array with sequential numbers
    For i = LBound(numbers) To UBound(numbers)
        numbers(i) = minValue + i - 1
        If numbers(i) > maxValue Then
            numbers(i) = maxValue
        End If
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
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Format as table
    With dataRange
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
End Sub

Private Function HandleUserInteraction() As Boolean
    ' Check for user cancellation (ESC key)
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then
        If MsgBox("Do you want to stop processing?", vbQuestion + vbYesNo) = vbYes Then
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

Private Function ValidateDataStructure(ByRef ws As Worksheet) As Boolean
    Dim lastCol As Long
    Dim headerRange As Range
    Dim cell As Range
    
    ' Check if worksheet has data
    If ws.Cells(1, 1).Value = "" Then
        ValidateDataStructure = False
        Exit Function
    End If
    
    ' Check headers
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    For Each cell In headerRange
        If Len(Trim(cell.Value)) = 0 Then
            ValidateDataStructure = False
            Exit Function
        End If
    Next cell
    
    ValidateDataStructure = True
End Function

Private Sub DisplayProcessingResults()
    Dim msg As String
    
    msg = "Processing completed successfully." & vbNewLine & vbNewLine & _
          "Total records processed: " & FormatNumber(gStats.TotalRecords) & vbNewLine & _
          "Approved records: " & FormatNumber(gStats.ApprovedRecords) & vbNewLine & _
          "Samples created: " & FormatNumber(gStats.SamplesCreated) & vbNewLine & _
          "Processing time: " & FormatNumber(gStats.ProcessingTime) & " seconds" & vbNewLine & _
          "Error count: " & FormatNumber(gStats.ErrorCount)
    
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
