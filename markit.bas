Option Explicit

Sub ProcessMarkitAndApprovedFunds()
    ' Performance-optimized macro for processing Approved funds and Markit files
    ' Handles large datasets (35,000+ rows) efficiently
    
    ' ===== Variable declarations =====
    ' Workbooks and worksheets
    Dim wbMaster As Workbook
    Dim wbApproved As Workbook
    Dim wbMarkit As Workbook
    Dim wsApproved As Worksheet
    Dim wsMarkit As Worksheet
    Dim wsRawData As Worksheet
    Dim wsUpload As Worksheet
    
    ' File paths
    Dim approvedFilePath As String
    Dim markitFilePath As String
    
    ' Tables
    Dim tblApproved As ListObject
    Dim tblMarkit As ListObject
    Dim tblRaw As ListObject
    Dim tblUpload As ListObject
    
    ' Range variables
    Dim lastRowApproved As Long
    Dim lastColApproved As Long
    Dim lastRowMarkit As Long
    Dim lastColMarkit As Long
    Dim approvedData As Range
    Dim markitData As Range
    
    ' Arrays for data processing
    Dim approvedArray() As Variant
    Dim markitArray() As Variant
    Dim rawArray() As Variant
    Dim uploadArray() As Variant
    Dim rawDataHeaders(14) As String
    
    ' Dictionary objects for faster lookups
    Dim approvedColMap As Object
    Dim markitColMap As Object
    Dim fundCodeMap As Object
    Dim fundLEIMap As Object
    
    ' Loop counters and flags
    Dim i As Long, j As Long, k As Long 
    Dim foundMatch As Boolean
    Dim matchFundCode As Boolean
    Dim matchFundLEI As Boolean
    
    ' Timing and status variables
    Dim startTime As Double
    Dim endTime As Double
    Dim executionTime As String
    
    ' Record start time
    startTime = Timer
    
    ' ===== Setup and optimization settings =====
    ' Turn off screen updating, events, and calculations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Initializing..."
    
    ' Initialize dictionary objects
    Set approvedColMap = CreateObject("Scripting.Dictionary")
    Set markitColMap = CreateObject("Scripting.Dictionary")
    Set fundCodeMap = CreateObject("Scripting.Dictionary")
    Set fundLEIMap = CreateObject("Scripting.Dictionary")
    
    ' Set reference to the master workbook (current workbook)
    Set wbMaster = ThisWorkbook
    
    ' ===== File selection =====
    ' Ask user to locate the Approved funds file
    Application.StatusBar = "Please select the Approved funds file..."
    MsgBox "Please select the Approved funds file.", vbInformation, "Select Approved Funds File"
    approvedFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Approved Funds File")
    
    ' Check if user canceled the file selection
    If approvedFilePath = "False" Then
        MsgBox "Operation canceled by user.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' Ask user to locate the Markit file
    Application.StatusBar = "Please select the Markit file..."
    MsgBox "Please select the Markit file.", vbInformation, "Select Markit File"
    markitFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Markit File")
    
    ' Check if user canceled the file selection
    If markitFilePath = "False" Then
        MsgBox "Operation canceled by user.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' ===== Open and process files =====
    On Error Resume Next
    Application.StatusBar = "Opening files..."
    
    ' Open the selected files with optimized settings (read-only, no links)
    Set wbApproved = Workbooks.Open(approvedFilePath, UpdateLinks:=False, ReadOnly:=True)
    If wbApproved Is Nothing Then
        MsgBox "Failed to open Approved funds file. Please check the file.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    Set wbMarkit = Workbooks.Open(markitFilePath, UpdateLinks:=False, ReadOnly:=True)
    If wbMarkit Is Nothing Then
        MsgBox "Failed to open Markit file. Please check the file.", vbExclamation
        wbApproved.Close SaveChanges:=False
        GoTo CleanupAndExit
    End If
    On Error GoTo 0
    
    ' ===== Process Approved funds file =====
    Application.StatusBar = "Processing Approved funds file..."
    Set wsApproved = wbApproved.Sheets(1)
    
    ' Delete the first row
    wsApproved.Rows(1).Delete
    
    ' Determine data range
    lastRowApproved = wsApproved.Cells(wsApproved.Rows.Count, "A").End(xlUp).Row
    lastColApproved = wsApproved.Cells(1, wsApproved.Columns.Count).End(xlToLeft).Column
    
    If lastRowApproved < 1 Or lastColApproved < 1 Then
        MsgBox "No data found in Approved funds file.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    Set approvedData = wsApproved.Range(wsApproved.Cells(1, 1), wsApproved.Cells(lastRowApproved, lastColApproved))
    
    ' Convert to table if needed
    On Error Resume Next
    If wsApproved.ListObjects.Count > 0 Then
        Set tblApproved = wsApproved.ListObjects(1)
        tblApproved.Name = "Approved"
    Else
        Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes)
        If Err.Number <> 0 Then
            ' Try to clear any existing tables first
            Err.Clear
            For i = wsApproved.ListObjects.Count To 1 Step -1
                wsApproved.ListObjects(i).Unlist
            Next i
            
            Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes)
            If Err.Number <> 0 Then
                MsgBox "Error creating Approved table: " & Err.Description, vbExclamation
                GoTo CleanupAndExit
            End If
        End If
        tblApproved.Name = "Approved"
    End If
    On Error GoTo 0
    
    ' Filter to keep only 'FI-EMEA', 'FI-US', and 'FI-GMC-ASIA' in 'Business Unit' column
    Dim buColIndex As Long
    buColIndex = 0
    
    For i = 1 To tblApproved.ListColumns.Count
        If tblApproved.ListColumns(i).Name = "Business Unit" Then
            buColIndex = i
            Exit For
        End If
    Next i
    
    If buColIndex > 0 Then
        Application.StatusBar = "Filtering Approved funds by Business Unit..."
        tblApproved.Range.AutoFilter Field:=buColIndex, Criteria1:=Array("FI-EMEA", "FI-US", "FI-GMC-ASIA"), Operator:=xlFilterValues
    Else
        MsgBox "Warning: 'Business Unit' column not found in Approved funds file.", vbExclamation
    End If
    
    ' Load data into memory
    Application.StatusBar = "Loading Approved data into memory..."
    
    ' Create column mapping
    For i = 1 To tblApproved.ListColumns.Count
        If Not approvedColMap.Exists(tblApproved.ListColumns(i).Name) Then
            approvedColMap.Add tblApproved.ListColumns(i).Name, i
        End If
    Next i
    
    ' If filtered, get only visible cells
    If tblApproved.ShowAutoFilter Then
        ' Count visible rows for array allocation
        Dim visibleRowCount As Long
        Dim visibleCells As Range
        On Error Resume Next
        Set visibleCells = tblApproved.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not visibleCells Is Nothing Then
            ' Determine visible row count
            visibleRowCount = WorksheetFunction.CountA(tblApproved.ListColumns(1).Range.SpecialCells(xlCellTypeVisible)) - 1
            If visibleRowCount <= 0 Then visibleRowCount = 1 ' Ensure at least one row
            
            ' Create array with header row
            ReDim approvedArray(1 To visibleRowCount + 1, 1 To tblApproved.ListColumns.Count)
            
            ' Add header row
            For i = 1 To tblApproved.ListColumns.Count
                approvedArray(1, i) = tblApproved.ListColumns(i).Name
            Next i
            
            ' Copy visible rows data
            If visibleRowCount > 0 Then
                Dim dataRow As Long
                dataRow = 2 ' Start after header
                
                Dim visibleRow As Range
                Dim visibleArea As Range
                
                ' Loop through each visible cell area (might be non-contiguous due to filter)
                For Each visibleArea In visibleCells.Areas
                    For Each visibleRow In visibleArea.Rows
                        If dataRow <= UBound(approvedArray, 1) Then
                            For i = 1 To tblApproved.ListColumns.Count
                                approvedArray(dataRow, i) = visibleRow.Cells(1, i).Value
                            Next i
                            dataRow = dataRow + 1
                        End If
                    Next visibleRow
                Next visibleArea
            End If
        Else
            ' If no visible cells (all filtered out), create minimal array
            ReDim approvedArray(1 To 1, 1 To tblApproved.ListColumns.Count)
            For i = 1 To tblApproved.ListColumns.Count
                approvedArray(1, i) = tblApproved.ListColumns(i).Name
            Next i
        End If
    Else
        ' No filter, get all data
        approvedArray = tblApproved.Range.Value
    End If
    
    ' ===== Process Markit file =====
    Application.StatusBar = "Processing Markit file..."
    Set wsMarkit = wbMarkit.Sheets(1)
    
    ' Determine data range
    lastRowMarkit = wsMarkit.Cells(wsMarkit.Rows.Count, "A").End(xlUp).Row
    lastColMarkit = wsMarkit.Cells(1, wsMarkit.Columns.Count).End(xlToLeft).Column
    
    If lastRowMarkit < 1 Or lastColMarkit < 1 Then
        MsgBox "No data found in Markit file.", vbExclamation
        GoTo CleanupAndExit
    End If
    
    Set markitData = wsMarkit.Range(wsMarkit.Cells(1, 1), wsMarkit.Cells(lastRowMarkit, lastColMarkit))
    
    ' Convert to table if needed
    On Error Resume Next
    If wsMarkit.ListObjects.Count > 0 Then
        Set tblMarkit = wsMarkit.ListObjects(1)
        tblMarkit.Name = "Markit"
    Else
        Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes)
        If Err.Number <> 0 Then
            ' Try to clear any existing tables first
            Err.Clear
            For i = wsMarkit.ListObjects.Count To 1 Step -1
                wsMarkit.ListObjects(i).Unlist
            Next i
            
            Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes)
            If Err.Number <> 0 Then
                MsgBox "Error creating Markit table: " & Err.Description, vbExclamation
                GoTo CleanupAndExit
            End If
        End If
        tblMarkit.Name = "Markit"
    End If
    On Error GoTo 0
    
    ' Load data into memory
    Application.StatusBar = "Loading Markit data into memory..."
    markitArray = tblMarkit.Range.Value
    
    ' Create column mapping
    For i = 1 To tblMarkit.ListColumns.Count
        If Not markitColMap.Exists(tblMarkit.ListColumns(i).Name) Then
            markitColMap.Add tblMarkit.ListColumns(i).Name, i
        End If
    Next i
    
    ' Build lookup dictionaries for faster matching
    Application.StatusBar = "Building lookup tables for faster matching..."
    Dim codeColIndex As Long, leiColIndex As Long
    
    If markitColMap.Exists("Client Identifier") Then
        codeColIndex = markitColMap("Client Identifier")
        
        ' Build Fund Code lookup map
        For i = 2 To UBound(markitArray, 1) ' Skip header row
            If Not IsEmpty(markitArray(i, codeColIndex)) Then
                ' If fund code already exists, keep the first occurrence
                If Not fundCodeMap.Exists(CStr(markitArray(i, codeColIndex))) Then
                    fundCodeMap.Add CStr(markitArray(i, codeColIndex)), i
                End If
            End If
        Next i
    End If
    
    If markitColMap.Exists("LEI") Then
        leiColIndex = markitColMap("LEI")
        
        ' Build Fund LEI lookup map
        For i = 2 To UBound(markitArray, 1) ' Skip header row
            If Not IsEmpty(markitArray(i, leiColIndex)) Then
                ' If LEI already exists, keep the first occurrence
                If Not fundLEIMap.Exists(CStr(markitArray(i, leiColIndex))) Then
                    fundLEIMap.Add CStr(markitArray(i, leiColIndex)), i
                End If
            End If
        Next i
    End If
    
    ' ===== Prepare master workbook =====
    Application.StatusBar = "Preparing master workbook..."
    
    ' Delete existing sheets if they exist
    On Error Resume Next
    Application.DisplayAlerts = False
    wbMaster.Sheets("Approved").Delete
    wbMaster.Sheets("Markit").Delete
    wbMaster.Sheets("Raw_data").Delete
    wbMaster.Sheets("Markit NAV today date").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheets
    Set wsApproved = wbMaster.Sheets.Add(After:=wbMaster.Sheets(wbMaster.Sheets.Count))
    wsApproved.Name = "Approved"
    
    Set wsMarkit = wbMaster.Sheets.Add(After:=wsApproved)
    wsMarkit.Name = "Markit"
    
    ' ===== Copy data to master workbook =====
    Application.StatusBar = "Copying Approved data to master workbook..."
    If Not IsEmpty(approvedArray) Then
        ' Find dimensions of the approved array
        If IsArray(approvedArray) Then
            lastRowApproved = UBound(approvedArray, 1)
            lastColApproved = UBound(approvedArray, 2)
            wsApproved.Range("A1").Resize(lastRowApproved, lastColApproved).Value = approvedArray
        End If
    End If
    
    Application.StatusBar = "Copying Markit data to master workbook..."
    If Not IsEmpty(markitArray) Then
        ' Find dimensions of the markit array
        If IsArray(markitArray) Then
            lastRowMarkit = UBound(markitArray, 1)
            lastColMarkit = UBound(markitArray, 2)
            wsMarkit.Range("A1").Resize(lastRowMarkit, lastColMarkit).Value = markitArray
        End If
    End If
    
    ' Convert copied data to tables in master workbook
    Application.StatusBar = "Creating tables in master workbook..."
    
    ' For Approved table
    lastRowApproved = wsApproved.Cells(wsApproved.Rows.Count, "A").End(xlUp).Row
    lastColApproved = wsApproved.Cells(1, wsApproved.Columns.Count).End(xlToLeft).Column
    
    If lastRowApproved > 0 And lastColApproved > 0 Then
        On Error Resume Next
        Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, wsApproved.Range("A1").Resize(lastRowApproved, lastColApproved), , xlYes)
        tblApproved.Name = "Approved"
        On Error GoTo 0
    End If
    
    ' For Markit table
    lastRowMarkit = wsMarkit.Cells(wsMarkit.Rows.Count, "A").End(xlUp).Row
    lastColMarkit = wsMarkit.Cells(1, wsMarkit.Columns.Count).End(xlToLeft).Column
    
    If lastRowMarkit > 0 And lastColMarkit > 0 Then
        On Error Resume Next
        Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, wsMarkit.Range("A1").Resize(lastRowMarkit, lastColMarkit), , xlYes)
        tblMarkit.Name = "Markit"
        On Error GoTo 0
    End If
    
    ' Close source files to free up memory
    Application.StatusBar = "Closing source files..."
    wbApproved.Close SaveChanges:=False
    wbMarkit.Close SaveChanges:=False
    
    ' ===== Create Raw_data sheet =====
    Application.StatusBar = "Creating Raw_data sheet..."
    Set wsRawData = wbMaster.Sheets.Add(After:=wsMarkit)
    wsRawData.Name = "Raw_data"
    
    ' Define columns for Raw table
    rawDataHeaders(0) = "Business Unit"
    rawDataHeaders(1) = "IA GCI"
    rawDataHeaders(2) = "RFAD Investment Manager"
    rawDataHeaders(3) = "Markit Investment Manager"
    rawDataHeaders(4) = "Fund GCI"
    rawDataHeaders(5) = "RFAD Fund Name"
    rawDataHeaders(6) = "Markit Fund Name"
    rawDataHeaders(7) = "Fund LEI"
    rawDataHeaders(8) = "Fund Code"
    rawDataHeaders(9) = "RFAD Currency"
    rawDataHeaders(10) = "Markit Currency"
    rawDataHeaders(11) = "RFAD Latest NAV Date"
    rawDataHeaders(12) = "RFAD Latest NAV"
    rawDataHeaders(13) = "Markit Latest NAV Date"
    rawDataHeaders(14) = "Markit Latest NAV"
    
    ' Create headers for Raw table
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        wsRawData.Cells(1, i + 1).Value = rawDataHeaders(i)
    Next i
    
    ' ===== Process data for Raw table =====
    Application.StatusBar = "Processing data for Raw table..."
    
    ' Determine dimensions for approved data (excluding header)
    Dim approvedRowCount As Long
    approvedRowCount = UBound(approvedArray, 1) - 1 ' Subtract header row
    
    ' Pre-allocate Raw array (max size = approved count + header)
    Dim maxRawRows As Long
    maxRawRows = approvedRowCount + 1 ' +1 for header
    ReDim rawArray(1 To maxRawRows, 1 To UBound(rawDataHeaders) + 1)
    
    ' Add header row to rawArray
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        rawArray(1, i + 1) = rawDataHeaders(i)
    Next i
    
    ' Row counter for rawArray
    Dim rawRowCount As Long
    rawRowCount = 1 ' Start at 1 (header row)
    
    ' Process matching
    Application.StatusBar = "Matching Approved and Markit data (this may take time)..."
    
    ' Loop through Approved data (skip header row)
    For i = 2 To UBound(approvedArray, 1)
        ' Update status every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & i & " of " & UBound(approvedArray, 1) & "..."
            DoEvents ' Allow UI to update
        End If
        
        ' Get fund code and LEI from Approved array
        Dim approvedFundCode As String, approvedFundLEI As String, fundGCI As String
        
        ' Initialize variables
        approvedFundCode = ""
        approvedFundLEI = ""
        fundGCI = ""
        
        ' Get values safely
        On Error Resume Next
        If approvedColMap.Exists("Fund Code") Then
            approvedFundCode = CStr(approvedArray(i, approvedColMap("Fund Code")))
        End If
        
        If approvedColMap.Exists("Fund LEI") Then
            approvedFundLEI = CStr(approvedArray(i, approvedColMap("Fund LEI")))
        End If
        
        If approvedColMap.Exists("Fund GCI") Then
            fundGCI = CStr(approvedArray(i, approvedColMap("Fund GCI")))
        Else
            ' Skip this row if Fund GCI doesn't exist
            GoTo NextApprovedRow
        End If
        On Error GoTo 0
        
        ' Skip if both Fund Code and Fund LEI are empty
        If Len(Trim(approvedFundCode)) = 0 And Len(Trim(approvedFundLEI)) = 0 Then
            GoTo NextApprovedRow
        End If
        
        ' Initialize match variables
        foundMatch = False
        matchFundCode = False
        matchFundLEI = False
        Dim matchedMarkitRow As Long
        matchedMarkitRow = 0
        
        ' Try to match by Fund Code first (using dictionary lookup)
        If Len(Trim(approvedFundCode)) > 0 Then
            If fundCodeMap.Exists(approvedFundCode) Then
                matchFundCode = True
                matchedMarkitRow = fundCodeMap(approvedFundCode)
            End If
        End If
        
        ' If Fund Code didn't match, try Fund LEI
        If Not matchFundCode And Len(Trim(approvedFundLEI)) > 0 Then
            If fundLEIMap.Exists(approvedFundLEI) Then
                matchFundLEI = True
                matchedMarkitRow = fundLEIMap(approvedFundLEI)
            End If
        End If
        
        ' If either Fund Code or Fund LEI matches, populate Raw array
        If matchFundCode Or matchFundLEI Then
            foundMatch = True
            rawRowCount = rawRowCount + 1
            
            ' Add data to rawArray
            ' 1. Business Unit from Approved
            If approvedColMap.Exists("Business Unit") Then
                rawArray(rawRowCount, 1) = approvedArray(i, approvedColMap("Business Unit"))
            End If
            
            ' 2. IA GCI from Approved
            If approvedColMap.Exists("IA GCI") Then
                rawArray(rawRowCount, 2) = approvedArray(i, approvedColMap("IA GCI"))
            End If
            
            ' 3. RFAD Investment Manager from Approved
            If approvedColMap.Exists("Investment Manager") Then
                rawArray(rawRowCount, 3) = approvedArray(i, approvedColMap("Investment Manager"))
            End If
            
            ' 4. Markit Investment Manager from Markit
            If markitColMap.Exists("Investment Manager") Then
                rawArray(rawRowCount, 4) = markitArray(matchedMarkitRow, markitColMap("Investment Manager"))
            End If
            
            ' 5. Fund GCI from Approved
            rawArray(rawRowCount, 5) = fundGCI
            
            ' 6. RFAD Fund Name from Approved
            If approvedColMap.Exists("Fund Name") Then
                rawArray(rawRowCount, 6) = approvedArray(i, approvedColMap("Fund Name"))
            End If
            
            ' 7. Markit Fund Name from Markit
            If markitColMap.Exists("Full Legal Fund Name") Then
                rawArray(rawRowCount, 7) = markitArray(matchedMarkitRow, markitColMap("Full Legal Fund Name"))
            End If
            
            ' 8. Fund LEI from Approved
            If approvedColMap.Exists("Fund LEI") Then
                rawArray(rawRowCount, 8) = approvedArray(i, approvedColMap("Fund LEI"))
            End If
            
            ' 9. Fund Code from Approved
            If approvedColMap.Exists("Fund Code") Then
                rawArray(rawRowCount, 9) = approvedArray(i, approvedColMap("Fund Code"))
            End If
            
            ' 10. RFAD Currency from Approved
            If approvedColMap.Exists("Currency") Then
                rawArray(rawRowCount, 10) = approvedArray(i, approvedColMap("Currency"))
            End If
            
            ' 11. Markit Currency from Markit
            If markitColMap.Exists("Currency") Then
                rawArray(rawRowCount, 11) = markitArray(matchedMarkitRow, markitColMap("Currency"))
            End If
            
            ' 12. RFAD Latest NAV Date from Approved
            If approvedColMap.Exists("Latest NAV Date") Then
                rawArray(rawRowCount, 12) = approvedArray(i, approvedColMap("Latest NAV Date"))
            End If
            
            ' 13. RFAD Latest NAV from Approved
            If approvedColMap.Exists("Latest NAV") Then
                rawArray(rawRowCount, 13) = approvedArray(i, approvedColMap("Latest NAV"))
            End If
            
            ' 14. Markit Latest NAV Date from Markit
            If markitColMap.Exists("As of Date") Then
                rawArray(rawRowCount, 14) = markitArray(matchedMarkitRow, markitColMap("As of Date"))
            End If
            
            ' 15. Markit Latest NAV from Markit
            If markitColMap.Exists("AUM Value") Then
                rawArray(rawRowCount, 15) = markitArray(matchedMarkitRow, markitColMap("AUM Value"))
            End If
        End If
NextApprovedRow:
    Next i
    
    ' Resize rawArray to actual data size (can't use ReDim Preserve on first dimension of 2D array)
    ' Instead, we'll create a new array with the exact size we need
    If rawRowCount > 1 Then
        Dim finalRawArray() As Variant
        Dim j As Long, k As Long
        ReDim finalRawArray(1 To rawRowCount, 1 To UBound(rawDataHeaders) + 1)
        
        ' Copy data from original array to the final array
        For j = 1 To rawRowCount
            For k = 1 To UBound(rawDataHeaders) + 1
                finalRawArray(j, k) = rawArray(j, k)
            Next k
        Next j
        
        ' Write the correctly sized array to the worksheet
        wsRawData.Range("A1").Resize(rawRowCount, UBound(rawDataHeaders) + 1).Value = finalRawArray
    Else
        ' If no data, just write the header row
        wsRawData.Range("A1").Resize(1, UBound(rawDataHeaders) + 1).Value = Application.WorksheetFunction.Index(rawArray, 1, 0)
    End If
    
    ' Create Raw table
    On Error Resume Next
    Set tblRaw = wsRawData.ListObjects.Add(xlSrcRange, wsRawData.Range("A1").Resize(rawRowCount, UBound(rawDataHeaders) + 1), , xlYes)
    tblRaw.Name = "Raw"
    On Error GoTo 0
    
    ' ===== Create Upload sheet =====
    Application.StatusBar = "Creating Upload sheet..."
    Set wsUpload = wbMaster.Sheets.Add(After:=wsRawData)
    wsUpload.Name = "Markit NAV today date"
    
    ' Create Upload table headers (same as Raw + Delta)
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        wsUpload.Cells(1, i + 1).Value = rawDataHeaders(i)
    Next i
    wsUpload.Cells(1, UBound(rawDataHeaders) + 2).Value = "Delta"
    
    ' Create Upload array with headers
    ReDim uploadArray(1 To 1, 1 To UBound(rawDataHeaders) + 2)
    
    ' Add header row
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        uploadArray(1, i + 1) = rawDataHeaders(i)
    Next i
    uploadArray(1, UBound(rawDataHeaders) + 2) = "Delta"
    
    ' Prepare for Upload data processing
    Dim uploadRowCount As Long
    uploadRowCount = 1 ' Start at 1 (header row)
    
    ' Process data for Upload table
    Application.StatusBar = "Analyzing NAV dates for Upload table..."
    
    ' Loop through Raw data (skip header row)
    For i = 2 To UBound(rawArray, 1)
        ' Get dates
        Dim rfadNAVDate As Variant, markitNAVDate As Variant
        Dim isMoreRecent As Boolean
        
        rfadNAVDate = rawArray(i, 12)
        markitNAVDate = rawArray(i, 14)
        
        ' Check if Markit date is more recent
        isMoreRecent = False
        
        ' Compare dates if both exist
        If Not IsEmpty(rfadNAVDate) And Not IsEmpty(markitNAVDate) Then
            ' Convert to dates if needed
            If Not IsDate(rfadNAVDate) Then
                On Error Resume Next
                rfadNAVDate = CDate(rfadNAVDate)
                On Error GoTo 0
            End If
            
            If Not IsDate(markitNAVDate) Then
                On Error Resume Next
                markitNAVDate = CDate(markitNAVDate)
                On Error GoTo 0
            End If
            
            ' Compare dates if both are valid
            If IsDate(rfadNAVDate) And IsDate(markitNAVDate) Then
                isMoreRecent = (CDate(markitNAVDate) > CDate(rfadNAVDate))
            End If
        End If
        
        ' If Markit date is more recent, add to Upload
        If isMoreRecent Then
            uploadRowCount = uploadRowCount + 1
            
            ' If we need to resize the array
            If uploadRowCount > UBound(uploadArray, 1) Then
                ' Create a new larger array
                Dim tempUploadArray() As Variant
                Dim k As Long
                ReDim tempUploadArray(1 To uploadRowCount * 2, 1 To UBound(rawDataHeaders) + 2)
                
                ' Copy existing data
                For j = 1 To UBound(uploadArray, 1)
                    For k = 1 To UBound(uploadArray, 2)
                        tempUploadArray(j, k) = uploadArray(j, k)
                    Next k
                Next j
                
                ' Replace the old array with the new one
                uploadArray = tempUploadArray
            End If
            
            ' Copy all columns from Raw to Upload
            For j = 1 To UBound(rawDataHeaders) + 1
                uploadArray(uploadRowCount, j) = rawArray(i, j)
            Next j
            
            ' Calculate Delta
            Dim rfadNAV As Variant, markitNAV As Variant
            Dim deltaValue As Double
            
            rfadNAV = rawArray(i, 13)
            markitNAV = rawArray(i, 15)
            
            ' Calculate Delta if both NAVs are valid numbers
            If IsNumeric(rfadNAV) And IsNumeric(markitNAV) And CDbl(rfadNAV) <> 0 Then
                deltaValue = (CDbl(markitNAV) / CDbl(rfadNAV)) - 1
                uploadArray(uploadRowCount, UBound(rawDataHeaders) + 2) = deltaValue
            End If
        End If
    Next i
    
    ' Resize uploadArray to actual data size (can't use ReDim Preserve on first dimension of 2D array)
    If uploadRowCount > 1 Then
        Dim finalUploadArray() As Variant
        Dim j As Long, k As Long
        ReDim finalUploadArray(1 To uploadRowCount, 1 To UBound(rawDataHeaders) + 2)
        
        ' Copy data from original array to the final array
        For j = 1 To uploadRowCount
            For k = 1 To UBound(rawDataHeaders) + 2
                finalUploadArray(j, k) = uploadArray(j, k)
            Next k
        Next j
        
        ' Write the correctly sized array to the worksheet
        wsUpload.Range("A1").Resize(uploadRowCount, UBound(rawDataHeaders) + 2).Value = finalUploadArray
    Else
        ' If no data, just write the header row
        wsUpload.Range("A1").Resize(1, UBound(rawDataHeaders) + 2).Value = uploadArray
    End If
    
    ' Create Upload table
    On Error Resume Next
    Set tblUpload = wsUpload.ListObjects.Add(xlSrcRange, wsUpload.Range("A1").Resize(uploadRowCount, UBound(uploadArray, 2)), , xlYes)
    tblUpload.Name = "Upload"
    On Error GoTo 0
    
    ' ===== Apply formatting =====
    Application.StatusBar = "Applying formatting..."
    
    ' Format date columns
    On Error Resume Next
    If Not tblRaw Is Nothing Then
        With tblRaw
            If .ListColumns.Count >= 12 Then .ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' RFAD Latest NAV Date
            If .ListColumns.Count >= 14 Then .ListColumns(14).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' Markit Latest NAV Date
            If .ListColumns.Count >= 13 Then .ListColumns(13).DataBodyRange.NumberFormat = "#,##0.00" ' RFAD Latest NAV
            If .ListColumns.Count >= 15 Then .ListColumns(15).DataBodyRange.NumberFormat = "#,##0.00" ' Markit Latest NAV
        End With
    End If
    
    If Not tblUpload Is Nothing Then
        With tblUpload
            If .ListColumns.Count >= 12 Then .ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' RFAD Latest NAV Date
            If .ListColumns.Count >= 14 Then .ListColumns(14).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' Markit Latest NAV Date
            If .ListColumns.Count >= 13 Then .ListColumns(13).DataBodyRange.NumberFormat = "#,##0.00" ' RFAD Latest NAV
            If .ListColumns.Count >= 15 Then .ListColumns(15).DataBodyRange.NumberFormat = "#,##0.00" ' Markit Latest NAV
            If .ListColumns.Count >= 16 Then .ListColumns(16).DataBodyRange.NumberFormat = "0.00%" ' Delta
        End With
        
        ' Highlight cells in Delta column if value >= 100% or <= -50%
        If Not tblUpload.DataBodyRange Is Nothing Then
            For i = 1 To tblUpload.ListRows.Count
                Dim cellValue As Variant
                cellValue = tblUpload.ListRows(i).Range.Cells(1, tblUpload.ListColumns.Count).Value
                
                If IsNumeric(cellValue) Then
                    If cellValue >= 1 Or cellValue <= -0.5 Then
                        tblUpload.ListRows(i).Range.Cells(1, tblUpload.ListColumns.Count).Interior.Color = RGB(255, 0, 0)
                    End If
                End If
            Next i
        End If
    End If
    On Error GoTo 0
    
    ' Auto-fit columns
    wsRawData.Columns.AutoFit
    wsUpload.Columns.AutoFit
    
    ' ===== Finish up =====
    ' Activate first sheet
    wbMaster.Sheets(1).Activate
    
    ' Calculate execution time
    endTime = Timer
    executionTime = Format((endTime - startTime) / 86400, "hh:mm:ss")
    
    ' Success message
    Application.StatusBar = False
    MsgBox "Processing completed successfully!" & vbCrLf & _
           "Total rows processed: " & (UBound(approvedArray, 1) - 1) & vbCrLf & _
           "Total matches found: " & (rawRowCount - 1) & vbCrLf & _
           "NAV updates found: " & (uploadRowCount - 1) & vbCrLf & _
           "Execution time: " & executionTime, vbInformation
    
    GoTo NormalExit

CleanupAndExit:
    ' Error handling cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' Close files if they're still open
    On Error Resume Next
    If Not wbApproved Is Nothing Then
        If wbApproved.Name <> "" Then wbApproved.Close SaveChanges:=False
    End If

    If Not wbMarkit Is Nothing Then
        If wbMarkit.Name <> "" Then wbMarkit.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    MsgBox "An error occurred during processing. Please check your files and try again.", vbExclamation
    Exit Sub
    
NormalExit:
    ' Normal exit cleanup
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
