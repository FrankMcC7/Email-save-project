Option Explicit

' Required References:
' - Microsoft Excel Object Library
' - Microsoft Office Object Library
' - Microsoft Scripting Runtime
' - Microsoft Visual Basic For Applications Extensibility

' ---------------------------------------------------
'  CONFIGURATION CONSTANTS
' ---------------------------------------------------
Private Const SAMPLE_SIZE As Long = 100              ' Default sample size
Private Const MIN_SAMPLE_SHEETS As Long = 5
Private Const MAX_SAMPLE_SHEETS As Long = 15
Private Const MIN_REFRESH_INTERVAL As Long = 5       ' seconds
Private Const MAX_REFRESH_INTERVAL As Long = 30      ' seconds
Private Const LOG_FILE_PATH As String = "DataProcessing_Log.txt"
Private Const MAX_RETRIES As Integer = 3
Private Const CHUNK_SIZE As Long = 10000             ' For large dataset processing

' ---------------------------------------------------
'  WIN32 API DECLARATIONS
' ---------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' ---------------------------------------------------
'  CUSTOM ERROR HANDLING ENUM
' ---------------------------------------------------
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

' ---------------------------------------------------
'  TYPE DEFINITION FOR PROCESSING STATISTICS
' ---------------------------------------------------
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

' ---------------------------------------------------
'  GLOBAL VARIABLES
' ---------------------------------------------------
Private gStats As ProcessingStats
Private gIsProcessing As Boolean
Private gCancelled As Boolean

' ===================================================
'                MAIN PROCESSING PROCEDURE
' ===================================================
Public Sub AutomatedDataProcessing()
    
    On Error GoTo ErrorHandler

    ' 1. Initialize environment (disable screen updating, events, etc.).
    If Not InitializeEnvironment() Then
        MsgBox "Failed to initialize environment.", vbCritical
        Exit Sub
    End If
    
    ' 2. Validate configuration constants.
    If Not ValidateConfiguration() Then
        Err.Raise ProcessingError.ERR_VALIDATION_FAILED, "AutomatedDataProcessing", _
                  "Configuration validation failed. Check constants (SAMPLE_SIZE, intervals, etc.)."
    End If
    
    ' 3. File selection and basic validation.
    Dim rawFilePath As String
    rawFilePath = SelectAndValidateInputFile()
    If rawFilePath = "" Then
        Err.Raise ProcessingError.ERR_NO_FILE_SELECTED, "AutomatedDataProcessing", _
                  "No valid file selected."
    End If
    
    ' 4. Import data (CSV or Excel).
    Dim wsRaw As Worksheet
    Set wsRaw = ImportData(rawFilePath)
    
    ' 5. Validate the imported worksheet (headers, data presence, etc.).
    If Not ValidateWorksheet(wsRaw) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "AutomatedDataProcessing", _
                  "Worksheet is invalid or contains no usable data."
    End If
    
    ' 6. Process raw data in chunks (filter "APPROVED" → "ApprovedData" sheet).
    Dim wsApproved As Worksheet
    Set wsApproved = ProcessRawDataInChunks(wsRaw)
    
    ' 7. Begin the sampling process.
    gIsProcessing = True
    Do While gIsProcessing And Not gCancelled
        If Not ProcessSampleBatch(wsApproved) Then Exit Do
        If Not HandleUserInteraction() Then Exit Do
        Call ProcessingDelay
    Loop
    
    ' 8. Display results if user did not cancel.
    If Not gCancelled Then
        DisplayProcessingResults
    End If

Cleanup:
    ' 9. Final cleanup (restore environment).
    CleanupEnvironment
    Exit Sub

ErrorHandler:
    ' Record error and increment error count
    gStats.LastError = Err.Description
    gStats.ErrorCount = gStats.ErrorCount + 1
    
    ' Log the error to file
    LogError Err.Number, Err.Description
    
    ' Show an error message to the user
    MsgBox "Error: " & Err.Description, vbCritical, "Processing Error"
    
    ' Clean up leftover/temporary objects
    CleanupOnError
    
    ' Resume the normal cleanup path
    Resume Cleanup
End Sub

' ---------------------------------------------------
'  INITIALIZATION & CONFIG
' ---------------------------------------------------
Private Function InitializeEnvironment() As Boolean
    On Error GoTo ErrorHandler
    
    ' Reset global flags and statistics
    gCancelled = False
    gIsProcessing = False
    
    With gStats
        .StartTime = Timer
        .EndTime = 0
        .TotalRecords = 0
        .ApprovedRecords = 0
        .SamplesCreated = 0
        .ErrorCount = 0
        .LastError = ""
        .ProcessingTime = 0
    End With
    
    ' Disable events, alerts, screen updating
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    UpdateStatus "Initializing data processing environment..."
    InitializeEnvironment = True
    Exit Function
    
ErrorHandler:
    InitializeEnvironment = False
    LogError Err.Number, "InitializeEnvironment: " & Err.Description
End Function

Private Sub CleanupEnvironment()
    ' Restore global flags if desired
    gIsProcessing = False
    
    ' Re-enable Excel settings
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
End Sub

Private Function ValidateConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check sample size > 0
    If SAMPLE_SIZE <= 0 Then
        gStats.LastError = "Invalid SAMPLE_SIZE."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Ensure min sample sheets <= max sample sheets
    If MIN_SAMPLE_SHEETS > MAX_SAMPLE_SHEETS Then
        gStats.LastError = "Invalid sample sheet range."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Ensure refresh intervals make sense
    If MIN_REFRESH_INTERVAL > MAX_REFRESH_INTERVAL Then
        gStats.LastError = "Invalid refresh interval configuration."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
    Exit Function
    
ErrorHandler:
    gStats.LastError = "Configuration validation error: " & Err.Description
    LogError Err.Number, "ValidateConfiguration: " & Err.Description
    ValidateConfiguration = False
End Function

' ---------------------------------------------------
'  FILE SELECTION & IMPORT
' ---------------------------------------------------
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
            Else
                ' Return empty string if file not accessible
                SelectAndValidateInputFile = ""
            End If
        Else
            SelectAndValidateInputFile = ""
        End If
    End With
End Function

Private Function ValidateFileAccess(ByVal filePath As String) As Boolean
    On Error Resume Next
    Dim fileNum As Integer
    Dim fileExists As Boolean
    
    fileExists = (Dir(filePath) <> "")
    If Not fileExists Then
        ValidateFileAccess = False
        Exit Function
    End If
    
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
        On Error GoTo ImportError
        Select Case fileExt
            Case ".csv"
                Set wsImport = ImportCSVWithProgress(filePath)
            Case Else
                Set wsImport = ImportExcelWithProgress(filePath)
        End Select
        On Error GoTo 0
        
        If Not wsImport Is Nothing Then
            Exit Do   ' Successfully imported
        End If
        
ImportError:
        retryCount = retryCount + 1
        If retryCount = MAX_RETRIES Then
            Err.Raise ProcessingError.ERR_FILE_ACCESS, "ImportData", _
                      "Failed to import file after " & MAX_RETRIES & " attempts."
        End If
        Application.Wait Now + TimeSerial(0, 0, 2) ' Wait 2 seconds before retry
        On Error GoTo 0
    Loop
    
    Set ImportData = wsImport
End Function

' ---------------------------------------------------
'  CSV/EXCEL IMPORT PLACEHOLDERS
' ---------------------------------------------------
Private Function ImportCSVWithProgress(ByVal filePath As String) As Worksheet
    ' TODO: Implement CSV-specific import logic
    ' For demonstration, we simply open the CSV and copy to a new worksheet.
    
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    
    UpdateStatus "Importing CSV: " & filePath
    
    Set wbTemp = Workbooks.Open(filePath)
    Set wsTemp = wbTemp.Sheets(1)
    
    ' Copy contents to new worksheet in ThisWorkbook
    Dim wsNew As Worksheet
    Set wsNew = ThisWorkbook.Worksheets.Add
    wsTemp.UsedRange.Copy wsNew.Range("A1")
    
    wbTemp.Close SaveChanges:=False
    
    wsNew.Name = "RawData_CSV"
    Set ImportCSVWithProgress = wsNew
End Function

Private Function ImportExcelWithProgress(ByVal filePath As String) As Worksheet
    ' TODO: Implement Excel-specific import logic
    ' For demonstration, just open the file, copy the first sheet into ThisWorkbook.

    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    
    UpdateStatus "Importing Excel: " & filePath
    
    Set wbTemp = Workbooks.Open(filePath)
    Set wsTemp = wbTemp.Sheets(1)
    
    ' Copy contents to a new worksheet in ThisWorkbook
    Dim wsNew As Worksheet
    Set wsNew = ThisWorkbook.Worksheets.Add
    wsTemp.UsedRange.Copy wsNew.Range("A1")
    
    wbTemp.Close SaveChanges:=False
    
    wsNew.Name = "RawData_Excel"
    Set ImportExcelWithProgress = wsNew
End Function

' ---------------------------------------------------
'  DATA VALIDATION
' ---------------------------------------------------
Private Function ValidateWorksheet(ByRef ws As Worksheet) As Boolean
    If ws Is Nothing Then
        ValidateWorksheet = False
        Exit Function
    End If
    
    ' Check for at least one data row besides the header
    If ws.UsedRange.Rows.Count <= 1 Then
        ValidateWorksheet = False
        Exit Function
    End If
    
    ' Validate headers (no empty column headers on row 1)
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
    
    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        If Len(Trim(cell.Value)) = 0 Then
            ValidateHeaders = False
            Exit Function
        End If
    Next cell
    
    ValidateHeaders = True
End Function

Private Function ValidateSampleSize(ByVal totalRows As Long) As Boolean
    ' SAMPLE_SIZE must be <= (number of data rows) and > 0
    If SAMPLE_SIZE <= 0 Or SAMPLE_SIZE > totalRows Then
        ValidateSampleSize = False
    Else
        ValidateSampleSize = True
    End If
End Function

' ---------------------------------------------------
'  DATA PROCESSING
' ---------------------------------------------------
Private Function ProcessRawDataInChunks(ByRef wsRaw As Worksheet) As Worksheet
    On Error GoTo ErrorHandler
    
    If Not ValidateWorksheet(wsRaw) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "ProcessRawDataInChunks", _
                  "Worksheet invalid or no data."
    End If
    
    Dim wsApproved As Worksheet
    Dim lastRow As Long, reviewStatusCol As Long
    
    ' Identify the "Review Status" column
    reviewStatusCol = GetColumnByHeaderName(wsRaw, "Review Status")
    If reviewStatusCol = 0 Then
        Err.Raise ProcessingError.ERR_NO_REVIEW_STATUS, "ProcessRawDataInChunks", _
                  "'Review Status' column not found."
    End If
    
    ' Create new "ApprovedData" sheet
    Set wsApproved = CreateApprovedSheet
    
    ' Process data in CHUNK_SIZE increments
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    Dim chunkStart As Long, chunkEnd As Long
    chunkStart = 2 ' skip header
    
    Do While chunkStart <= lastRow And Not gCancelled
        chunkEnd = Application.Min(chunkStart + CHUNK_SIZE - 1, lastRow)
        
        UpdateStatus "Processing rows " & chunkStart & " to " & chunkEnd
        ProcessDataChunk wsRaw, wsApproved, chunkStart, chunkEnd, reviewStatusCol
        
        chunkStart = chunkEnd + 1
        DoEvents
        
        If Not gIsProcessing Then Exit Do
    Loop
    
    Set ProcessRawDataInChunks = wsApproved
    Exit Function
    
ErrorHandler:
    LogError Err.Number, "ProcessRawDataInChunks: " & Err.Description
    Set ProcessRawDataInChunks = Nothing
End Function

Private Sub ProcessDataChunk(ByRef wsSource As Worksheet, ByRef wsTarget As Worksheet, _
                             ByVal startRow As Long, ByVal endRow As Long, _
                             ByVal statusCol As Long)
    Dim dataRange As Range
    Dim approvedRange As Range
    Dim cell As Range
    
    ' Define the chunk’s range
    Set dataRange = wsSource.Range(wsSource.Cells(startRow, 1), _
                                   wsSource.Cells(endRow, wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column))
    
    ' Identify rows with "APPROVED"
    For Each cell In dataRange.Columns(statusCol).Cells
        If UCase(Trim(cell.Value)) = "APPROVED" Then
            If approvedRange Is Nothing Then
                Set approvedRange = wsSource.Rows(cell.Row)
            Else
                Set approvedRange = Union(approvedRange, wsSource.Rows(cell.Row))
            End If
        End If
    Next cell
    
    ' Copy approved rows to the target sheet
    If Not approvedRange Is Nothing Then
        approvedRange.Copy Destination:=wsTarget.Cells(wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1, 1)
        gStats.ApprovedRecords = gStats.ApprovedRecords + approvedRange.Rows.Count
    End If
    
    gStats.TotalRecords = gStats.TotalRecords + (endRow - startRow + 1)
End Sub

' ---------------------------------------------------
'  SAMPLE GENERATION
' ---------------------------------------------------
Private Function ProcessSampleBatch(ByRef wsSource As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate the source data sheet
    If Not ValidateWorksheet(wsSource) Then
        Err.Raise ProcessingError.ERR_INVALID_DATA, "ProcessSampleBatch", _
                  "Source worksheet is invalid."
    End If
    
    ' Create some number of new sample sheets
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
    Dim sampleSheets As New Collection
    Dim sampleSheet As Worksheet
    Dim numSheets As Long, i As Long
    
    ' Randomly determine how many sample sheets to create
    Randomize
    numSheets = Int((MAX_SAMPLE_SHEETS - MIN_SAMPLE_SHEETS + 1) * Rnd + MIN_SAMPLE_SHEETS)
    
    ' Create each sample sheet
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
    
    ' Determine source range size
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set headerRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol))
    
    ' Validate sample size
    If Not ValidateSampleSize(lastRow - 1) Then
        Err.Raise ProcessingError.ERR_INVALID_SAMPLE_SIZE, "CreateSamples", _
                  "Sample size exceeds available data rows."
    End If
    
    ' Populate each sample sheet
    For Each sampleSheet In sampleSheets
        UpdateStatus "Generating sample on " & sampleSheet.Name
        
        ' Copy the header row
        headerRange.Copy Destination:=sampleSheet.Range("A1")
        
        ' Generate and copy a random sample
        GenerateRandomSample wsSource, sampleSheet, lastRow, lastCol
        
        ' Format the newly created sample sheet
        FormatSampleSheet sampleSheet
        
        gStats.SamplesCreated = gStats.SamplesCreated + 1
    Next sampleSheet
End Sub

Private Sub GenerateRandomSample(ByRef wsSource As Worksheet, ByRef wsTarget As Worksheet, _
                                 ByVal sourceRows As Long, ByVal sourceCols As Long)
    Dim randomRows() As Long
    Dim i As Long, j As Long
    Dim dataRange As Range
    Dim sourceData As Variant
    Dim targetData As Variant
    
    ' Decide how many random rows to pick
    ReDim randomRows(1 To Application.Min(SAMPLE_SIZE, sourceRows - 1))
    
    ' Fill array with unique row numbers (2..sourceRows) to skip headers
    GenerateUniqueRandomNumbers randomRows, 2, sourceRows
    
    ' Retrieve source data into a variant array
    Set dataRange = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(sourceRows, sourceCols))
    sourceData = dataRange.Value
    
    ' Prepare target array
    ReDim targetData(1 To UBound(randomRows), 1 To sourceCols)
    
    ' Fill target array
    For i = 1 To UBound(randomRows)
        For j = 1 To sourceCols
            ' randomRows(i) - 1 because dataRange starts at row 2 in the sheet
            targetData(i, j) = sourceData(randomRows(i) - 1, j)
        Next j
    Next i
    
    ' Write sample rows to target
    wsTarget.Range("A2").Resize(UBound(randomRows), sourceCols).Value = targetData
End Sub

Private Sub GenerateUniqueRandomNumbers(ByRef numbers() As Long, _
                                        ByVal minValue As Long, _
                                        ByVal maxValue As Long)
    Dim i As Long, j As Long, temp As Long
    Dim numCount As Long
    
    numCount = UBound(numbers) - LBound(numbers) + 1
    
    ' Ensure we have enough range to accommodate sample size
    If numCount > (maxValue - minValue + 1) Then
        Err.Raise ProcessingError.ERR_INVALID_SAMPLE_SIZE, "GenerateUniqueRandomNumbers", _
                  "Sample size is larger than available data range."
    End If
    
    ' Initialize array with sequential numbers
    For i = LBound(numbers) To UBound(numbers)
        numbers(i) = minValue + (i - LBound(numbers))
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
    Dim lastRow As Long, lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 1 Then
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        
        ' Apply borders
        With dataRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            ' Format header row
            .Rows(1).Font.Bold = True
            .Rows(1).Interior.Color = RGB(217, 217, 217)
        End With
        
        ws.Cells.EntireColumn.AutoFit
    End If
End Sub

' ---------------------------------------------------
'  SHEET MANAGEMENT
' ---------------------------------------------------
Private Function CreateApprovedSheet() As Worksheet
    Dim wsApproved As Worksheet
    
    ' Delete existing "ApprovedData" if present
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("ApprovedData").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new
    Set wsApproved = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsApproved.Name = "ApprovedData"
    
    ' Format the header row
    With wsApproved.Range("A1").Font
        .Bold = True
        .Size = 11
    End With
    
    With wsApproved
        .Rows(1).Interior.Color = RGB(217, 217, 217)
        .Cells.Clear
    End With
    
    Set CreateApprovedSheet = wsApproved
End Function

Private Function CreateSampleSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    ' Delete existing sheet with same name
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create new
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    
    Set CreateSampleSheet = ws
End Function

Private Sub CleanupWorksheets()
    ' Example: remove any leftover "SampleX" sheets
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "Sample") = 1 Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
End Sub

' ---------------------------------------------------
'  UTILITY FUNCTIONS
' ---------------------------------------------------
Private Function HandleUserInteraction() As Boolean
    ' Check if user pressed ESC
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
    
    ' Pause for random interval between MIN_REFRESH_INTERVAL and MAX_REFRESH_INTERVAL
    Randomize
    delaySeconds = Int((MAX_REFRESH_INTERVAL - MIN_REFRESH_INTERVAL + 1) * Rnd + MIN_REFRESH_INTERVAL)
    
    UpdateStatus "Waiting " & delaySeconds & " second(s) before next iteration..."
    
    startTime = Timer
    Do While Timer < startTime + delaySeconds
        DoEvents
        If Not HandleUserInteraction() Then Exit Do
    Loop
End Sub

Private Sub DisplayProcessingResults()
    gStats.EndTime = Timer
    gStats.ProcessingTime = gStats.EndTime - gStats.StartTime
    
    Dim msg As String
    msg = "Processing completed successfully." & vbNewLine & vbNewLine & _
          "Total records processed: " & FormatNumberCustom(gStats.TotalRecords) & vbNewLine & _
          "Approved records: " & FormatNumberCustom(gStats.ApprovedRecords) & vbNewLine & _
          "Samples created: " & FormatNumberCustom(gStats.SamplesCreated) & vbNewLine & _
          "Processing time: " & Format(gStats.ProcessingTime, "0.00") & " seconds" & vbNewLine & _
          "Error count: " & FormatNumberCustom(gStats.ErrorCount)
    
    If gStats.ErrorCount > 0 Then
        msg = msg & vbNewLine & vbNewLine & "Last error: " & gStats.LastError
    End If
    
    MsgBox msg, vbInformation, "Processing Results"
End Sub

Private Function GetColumnByHeaderName(ByRef ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, col As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If StrComp(ws.Cells(1, col).Value, headerName, vbTextCompare) = 0 Then
            GetColumnByHeaderName = col
            Exit Function
        End If
    Next col
    
    GetColumnByHeaderName = 0
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

Private Function FormatNumberCustom(ByVal number As Double) As String
    ' Simple numeric format (with thousand separators, no decimals)
    FormatNumberCustom = Format(number, "#,##0")
End Function

' ---------------------------------------------------
'  ERROR & CLEANUP HANDLERS
' ---------------------------------------------------
Private Sub LogError(errNumber As Long, errDescription As String)
    ' Logs error details to a text file.
    On Error Resume Next
    Dim fso As Object, ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(LOG_FILE_PATH, 8, True)  ' 8 = ForAppending
    
    ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & _
                 " | Error " & errNumber & " | " & errDescription
    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
End Sub

Private Sub CleanupOnError()
    ' Restore application state if an error occurred mid-processing
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    
    ' Optionally delete partially created sample sheets or anything else
    CleanupWorksheets
    
    ' Reset flags
    gIsProcessing = False
    gCancelled = False
End Sub
