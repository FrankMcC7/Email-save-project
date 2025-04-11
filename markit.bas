Option Explicit

Sub ProcessMarkitAndApprovedFunds()
    ' Performance-optimized macro for processing Approved funds and Markit files
    ' Handles large datasets (35,000+ rows) efficiently
    ' Modified to create three separate matching tables and update upload criteria
    
    ' ===== Variable declarations =====
    ' Workbooks and worksheets
    Dim wbMaster As Workbook
    Dim wbApproved As Workbook
    Dim wbMarkit As Workbook
    Dim wsApproved As Worksheet
    Dim wsMarkit As Worksheet
    Dim wsRawDataCode As Worksheet
    Dim wsRawDataLEI As Worksheet
    Dim wsRawDataName As Worksheet
    Dim wsUpload As Worksheet
    
    ' File paths
    Dim approvedFilePath As String
    Dim markitFilePath As String
    
    ' Tables
    Dim tblApproved As ListObject
    Dim tblMarkit As ListObject
    Dim tblRawCode As ListObject
    Dim tblRawLEI As ListObject
    Dim tblRawName As ListObject
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
    Dim rawCodeArray() As Variant
    Dim rawLEIArray() As Variant
    Dim rawNameArray() As Variant
    Dim uploadArray() As Variant
    Dim finalRawCodeArray() As Variant
    Dim finalRawLEIArray() As Variant
    Dim finalRawNameArray() As Variant
    Dim finalUploadArray() As Variant
    Dim tempUploadArray() As Variant
    Dim rawDataHeaders(14) As String
    
    ' Dictionary objects for faster lookups
    Dim approvedColMap As Object
    Dim markitColMap As Object
    Dim fundCodeMap As Object
    Dim fundLEIMap As Object
    Dim fundNameMap As Object
    Dim matchedFundMap As Object   ' To track which funds have been matched already
    
    ' Loop counters and flags
    Dim i As Long, j As Long, k As Long
    Dim foundMatch As Boolean
    Dim matchFundCode As Boolean
    Dim matchFundLEI As Boolean
    Dim matchFundName As Boolean
    
    ' Column indexes and values
    Dim buColIndex As Long
    Dim codeColIndex As Long
    Dim leiColIndex As Long
    Dim nameColIndex As Long
    Dim markitNameColIndex As Long
    Dim approvedFundCode As String
    Dim approvedFundLEI As String
    Dim approvedFundName As String
    Dim markitFundName As String
    Dim fundGCI As String
    Dim matchedMarkitRow As Long
    
    ' Date and NAV variables
    Dim rfadNAVDate As Variant
    Dim markitNAVDate As Variant
    Dim rfadNAV As Variant
    Dim markitNAV As Variant
    Dim isMoreRecentBy15Days As Boolean
    Dim daysDifference As Long
    Dim deltaValue As Double
    
    ' Row counters
    Dim approvedRowCount As Long
    Dim rawCodeRowCount As Long
    Dim rawLEIRowCount As Long
    Dim rawNameRowCount As Long
    Dim uploadRowCount As Long
    Dim maxRawRows As Long
    Dim visibleRowCount As Long
    Dim dataRow As Long
    
    ' Timing and status variables
    Dim startTime As Double
    Dim endTime As Double
    Dim executionTime As String
    
    ' Range objects for visible cells processing
    Dim visibleCells As Range
    Dim visibleArea As Range
    Dim visibleRow As Range
    
    ' For objects
    Dim existingObj As Object
    
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
    Set fundNameMap = CreateObject("Scripting.Dictionary")
    Set matchedFundMap = CreateObject("Scripting.Dictionary")  ' New: track already matched funds
    
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
                dataRow = 2 ' Start after header
                
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
    codeColIndex = 0
    leiColIndex = 0
    nameColIndex = 0
    markitNameColIndex = 0
    
    ' Find column indexes
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
    
    ' Add Fund Name lookup for name matching
    If markitColMap.Exists("Full Legal Fund Name") Then
        markitNameColIndex = markitColMap("Full Legal Fund Name")
        
        ' Build Fund Name lookup map
        For i = 2 To UBound(markitArray, 1) ' Skip header row
            If Not IsEmpty(markitArray(i, markitNameColIndex)) Then
                ' If fund name already exists, keep the first occurrence
                If Not fundNameMap.Exists(UCase(CStr(markitArray(i, markitNameColIndex)))) Then
                    fundNameMap.Add UCase(CStr(markitArray(i, markitNameColIndex))), i
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
    wbMaster.Sheets("Raw_data_Code").Delete
    wbMaster.Sheets("Raw_data_LEI").Delete
    wbMaster.Sheets("Raw_data_Name").Delete
    wbMaster.Sheets("Markit NAV Upload").Delete
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
    
    ' ===== Create three Raw_data sheets =====
    ' 1. First sheet for Fund Code matches
    Application.StatusBar = "Creating Raw_data_Code sheet..."
    Set wsRawDataCode = wbMaster.Sheets.Add(After:=wsMarkit)
    wsRawDataCode.Name = "Raw_data_Code"
    
    ' 2. Second sheet for Fund LEI matches
    Application.StatusBar = "Creating Raw_data_LEI sheet..."
    Set wsRawDataLEI = wbMaster.Sheets.Add(After:=wsRawDataCode)
    wsRawDataLEI.Name = "Raw_data_LEI"
    
    ' 3. Third sheet for Fund Name matches
    Application.StatusBar = "Creating Raw_data_Name sheet..."
    Set wsRawDataName = wbMaster.Sheets.Add(After:=wsRawDataLEI)
    wsRawDataName.Name = "Raw_data_Name"
    
    ' Define columns for Raw tables (same for all three)
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
    
    ' Create headers for all Raw tables
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        wsRawDataCode.Cells(1, i + 1).Value = rawDataHeaders(i)
        wsRawDataLEI.Cells(1, i + 1).Value = rawDataHeaders(i)
        wsRawDataName.Cells(1, i + 1).Value = rawDataHeaders(i)
    Next i
    
    ' ===== Process data for Raw tables =====
    Application.StatusBar = "Processing data for Raw tables..."
    
    ' Determine dimensions for approved data (excluding header)
    approvedRowCount = UBound(approvedArray, 1) - 1 ' Subtract header row
    
    ' Pre-allocate Raw arrays (max size = approved count + header)
    maxRawRows = approvedRowCount + 1 ' +1 for header
    ReDim rawCodeArray(1 To maxRawRows, 1 To UBound(rawDataHeaders) + 1)
    ReDim rawLEIArray(1 To maxRawRows, 1 To UBound(rawDataHeaders) + 1)
    ReDim rawNameArray(1 To maxRawRows, 1 To UBound(rawDataHeaders) + 1)
    
    ' Add header rows to all rawArrays
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        rawCodeArray(1, i + 1) = rawDataHeaders(i)
        rawLEIArray(1, i + 1) = rawDataHeaders(i)
        rawNameArray(1, i + 1) = rawDataHeaders(i)
    Next i
    
    ' Row counters for rawArrays
    rawCodeRowCount = 1 ' Start at 1 (header row)
    rawLEIRowCount = 1 ' Start at 1 (header row)
    rawNameRowCount = 1 ' Start at 1 (header row)
    
    ' Process matching - First by Fund Code
    Application.StatusBar = "Matching by Fund Code..."
    
    ' Get Fund Name column index in Approved data
    If approvedColMap.Exists("Fund Name") Then
        nameColIndex = approvedColMap("Fund Name")
    End If
    
    ' Loop through Approved data (skip header row)
    For i = 2 To UBound(approvedArray, 1)
        ' Update status every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing Fund Code matching: row " & i & " of " & UBound(approvedArray, 1) & "..."
            DoEvents ' Allow UI to update
        End If
        
        ' Initialize variables for this row
        approvedFundCode = ""
        approvedFundLEI = ""
        approvedFundName = ""
        fundGCI = ""
        foundMatch = False
        matchFundCode = False
        matchedMarkitRow = 0
        
        ' Get values safely
        On Error Resume Next
        If approvedColMap.Exists("Fund Code") Then
            approvedFundCode = CStr(approvedArray(i, approvedColMap("Fund Code")))
        End If
        
        If approvedColMap.Exists("Fund GCI") Then
            fundGCI = CStr(approvedArray(i, approvedColMap("Fund GCI")))
        Else
            ' Skip this row if Fund GCI doesn't exist
            GoTo NextApprovedRowCode
        End If
        
        If nameColIndex > 0 Then
            approvedFundName = CStr(approvedArray(i, nameColIndex))
        End If
        On Error GoTo 0
        
        ' Skip if Fund Code is empty
        If Len(Trim(approvedFundCode)) = 0 Then
            GoTo NextApprovedRowCode
        End If
        
        ' Try to match by Fund Code (using dictionary lookup)
        If fundCodeMap.Exists(approvedFundCode) Then
            matchFundCode = True
            matchedMarkitRow = fundCodeMap(approvedFundCode)
            
            ' Add to matched fund map to avoid duplicate matches
            If Not matchedFundMap.Exists(fundGCI) Then
                matchedFundMap.Add fundGCI, "Code"
                
                ' Populate Raw Code array
                rawCodeRowCount = rawCodeRowCount + 1
                PopulateRawArray rawCodeArray, rawCodeRowCount, approvedArray, markitArray, i, matchedMarkitRow, approvedColMap, markitColMap
            End If
        End If
NextApprovedRowCode:
    Next i
    
    ' Process matching - Next by Fund LEI
    Application.StatusBar = "Matching by Fund LEI..."
    
    ' Loop through Approved data again (skip header row)
    For i = 2 To UBound(approvedArray, 1)
        ' Update status every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing Fund LEI matching: row " & i & " of " & UBound(approvedArray, 1) & "..."
            DoEvents ' Allow UI to update
        End If
        
        ' Initialize variables for this row
        approvedFundCode = ""
        approvedFundLEI = ""
        fundGCI = ""
        foundMatch = False
        matchFundLEI = False
        matchedMarkitRow = 0
        
        ' Get values safely
        On Error Resume Next
        If approvedColMap.Exists("Fund LEI") Then
            approvedFundLEI = CStr(approvedArray(i, approvedColMap("Fund LEI")))
        End If
        
        If approvedColMap.Exists("Fund GCI") Then
            fundGCI = CStr(approvedArray(i, approvedColMap("Fund GCI")))
        Else
            ' Skip this row if Fund GCI doesn't exist
            GoTo NextApprovedRowLEI
        End If
        On Error GoTo 0
        
        ' Skip if Fund LEI is empty
        If Len(Trim(approvedFundLEI)) = 0 Then
            GoTo NextApprovedRowLEI
        End If
        
        ' Skip if already matched by Fund Code
        If matchedFundMap.Exists(fundGCI) Then
            GoTo NextApprovedRowLEI
        End If
        
        ' Try to match by Fund LEI
        If fundLEIMap.Exists(approvedFundLEI) Then
            matchFundLEI = True
            matchedMarkitRow = fundLEIMap(approvedFundLEI)
            
            ' Add to matched fund map to avoid duplicate matches
            If Not matchedFundMap.Exists(fundGCI) Then
                matchedFundMap.Add fundGCI, "LEI"
                
                ' Populate Raw LEI array
                rawLEIRowCount = rawLEIRowCount + 1
                PopulateRawArray rawLEIArray, rawLEIRowCount, approvedArray, markitArray, i, matchedMarkitRow, approvedColMap, markitColMap
            End If
        End If
NextApprovedRowLEI:
    Next i
    
    ' Process matching - Finally by Fund Name
    Application.StatusBar = "Matching by Fund Name..."
    
    ' Loop through Approved data one more time (skip header row)
    For i = 2 To UBound(approvedArray, 1)
        ' Update status every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing Fund Name matching: row " & i & " of " & UBound(approvedArray, 1) & "..."
            DoEvents ' Allow UI to update
        End If
        
        ' Initialize variables for this row
        approvedFundName = ""
        fundGCI = ""
        matchFundName = False
        matchedMarkitRow = 0
        
        ' Get values safely
        On Error Resume Next
        If nameColIndex > 0 Then
            approvedFundName = CStr(approvedArray(i, nameColIndex))
        End If
        
        If approvedColMap.Exists("Fund GCI") Then
            fundGCI = CStr(approvedArray(i, approvedColMap("Fund GCI")))
        Else
            ' Skip this row if Fund GCI doesn't exist
            GoTo NextApprovedRowName
        End If
        On Error GoTo 0
        
        ' Skip if Fund Name is empty
        If Len(Trim(approvedFundName)) = 0 Then
            GoTo NextApprovedRowName
        End If
        
        ' Skip if already matched by Fund Code or LEI
        If matchedFundMap.Exists(fundGCI) Then
            GoTo NextApprovedRowName
        End If
        
        ' Try to match by Fund Name (using exact match)
        If fundNameMap.Exists(UCase(approvedFundName)) Then
            matchFundName = True
            matchedMarkitRow = fundNameMap(UCase(approvedFundName))
            
            ' Add to matched fund map to avoid duplicate matches
            If Not matchedFundMap.Exists(fundGCI) Then
                matchedFundMap.Add fundGCI, "Name"
                
                ' Populate Raw Name array
                rawNameRowCount = rawNameRowCount + 1
                PopulateRawArray rawNameArray, rawNameRowCount, approvedArray, markitArray, i, matchedMarkitRow, approvedColMap, markitColMap
            End If
        End If
NextApprovedRowName:
    Next i
    
    ' Resize and write rawCodeArray
    If rawCodeRowCount > 1 Then
        ReDim finalRawCodeArray(1 To rawCodeRowCount, 1 To UBound(rawDataHeaders) + 1)
        
        ' Copy data from original array to the final array
        For i = 1 To rawCodeRowCount
            For k = 1 To UBound(rawDataHeaders) + 1
                finalRawCodeArray(i, k) = rawCodeArray(i, k)
            Next k
        Next i
        
        ' Write the correctly sized array to the worksheet
        Application.StatusBar = "Writing Raw_data_Code to worksheet..."
        wsRawDataCode.Range("A1").Resize(rawCodeRowCount, UBound(rawDataHeaders) + 1).Value = finalRawCodeArray
    Else
        ' If no data, just write the header row
        wsRawDataCode.Range("A1").Resize(1, UBound(rawDataHeaders) + 1).Value = Application.WorksheetFunction.Index(rawCodeArray, 1, 0)
    End If
    
    ' Resize and write rawLEIArray
    If rawLEIRowCount > 1 Then
        ReDim finalRawLEIArray(1 To rawLEIRowCount, 1 To UBound(rawDataHeaders) + 1)
        
        ' Copy data from original array to the final array
        For i = 1 To rawLEIRowCount
            For k = 1 To UBound(rawDataHeaders) + 1
                finalRawLEIArray(i, k) = rawLEIArray(i, k)
            Next k
        Next i
        
        ' Write the correctly sized array to the worksheet
        Application.StatusBar = "Writing Raw_data_LEI to worksheet..."
        wsRawDataLEI.Range("A1").Resize(rawLEIRowCount, UBound(rawDataHeaders) + 1).Value = finalRawLEIArray
    Else
        ' If no data, just write the header row
        wsRawDataLEI.Range("A1").Resize(1, UBound(rawDataHeaders) + 1).Value = Application.WorksheetFunction.Index(rawLEIArray, 1, 0)
    End If
    
    ' Resize and write rawNameArray
    If rawNameRowCount > 1 Then
        ReDim finalRawNameArray(1 To rawNameRowCount, 1 To UBound(rawDataHeaders) + 1)
        
        ' Copy data from original array to the final array
        For i = 1 To rawNameRowCount
            For k = 1 To UBound(rawDataHeaders) + 1
                finalRawNameArray(i, k) = rawNameArray(i, k)
            Next k
        Next i
        
        ' Write the correctly sized array to the worksheet
        Application.StatusBar = "Writing Raw_data_Name to worksheet..."
        wsRawDataName.Range("A1").Resize(rawNameRowCount, UBound(rawDataHeaders) + 1).Value = finalRawNameArray
    Else
        ' If no data, just write the header row
        wsRawDataName.Range("A1").Resize(1, UBound(rawDataHeaders) + 1).Value = Application.WorksheetFunction.Index(rawNameArray, 1, 0)
    End If
    
    ' Create Raw tables
    On Error Resume Next
    Set tblRawCode = wsRawDataCode.ListObjects.Add(xlSrcRange, wsRawDataCode.Range("A1").Resize(rawCodeRowCount, UBound(rawDataHeaders) + 1), , xlYes)
    tblRawCode.Name = "RawCode"
    
    Set tblRawLEI = wsRawDataLEI.ListObjects.Add(xlSrcRange, wsRawDataLEI.Range("A1").Resize(rawLEIRowCount, UBound(rawDataHeaders) + 1), , xlYes)
    tblRawLEI.Name = "RawLEI"
    
    Set tblRawName = wsRawDataName.ListObjects.Add(xlSrcRange, wsRawDataName.Range("A1").Resize(rawNameRowCount, UBound(rawDataHeaders) + 1), , xlYes)
    tblRawName.Name = "RawName"
    On Error GoTo 0
    
    ' ===== Create Upload sheet =====
    Application.StatusBar = "Creating Markit NAV Upload sheet..."
    Set wsUpload = wbMaster.Sheets.Add(After:=wsRawDataName)
    wsUpload.Name = "Markit NAV Upload"
    
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
    uploadRowCount = 1 ' Start at 1 (header row)
    
    ' Process data for Upload table from all three Raw tables
    Application.StatusBar = "Analyzing NAV dates for Upload table..."
    
    ' Process Raw Code data
    ProcessRawDataForUpload rawCodeArray, rawCodeRowCount, uploadArray, uploadRowCount
    
    ' Process Raw LEI data
    ProcessRawDataForUpload rawLEIArray, rawLEIRowCount, uploadArray, uploadRowCount
    
    ' Process Raw Name data
    ProcessRawDataForUpload rawNameArray, rawNameRowCount, uploadArray, uploadRowCount
    
    ' Resize uploadArray to actual data size
    If uploadRowCount > 1 Then
        ReDim finalUploadArray(1 To uploadRowCount, 1 To UBound(rawDataHeaders) + 2)
        
        ' Copy data from original array to the final array
        For i = 1 To uploadRowCount
            For k = 1 To UBound(rawDataHeaders) + 2
                finalUploadArray(i, k) = uploadArray(i, k)
            Next k
        Next i
        
        ' Write the correctly sized array to the worksheet
        Application.StatusBar = "Writing Upload data to worksheet..."
        wsUpload.Range("A1").Resize(uploadRowCount, UBound(rawDataHeaders) + 2).Value = finalUploadArray
    Else
        ' If no data, just write the header row
        wsUpload.Range("A1").Resize(1, UBound(rawDataHeaders) + 2).Value = uploadArray
    End If
    
    ' Create Upload table
    On Error Resume Next
    Set tblUpload = wsUpload.ListObjects.Add(xlSrcRange, wsUpload.Range("A1").Resize(uploadRowCount, UBound(rawDataHeaders) + 2), , xlYes)
    tblUpload.Name = "Upload"
    On Error GoTo 0
    
    ' ===== Apply formatting =====
    Application.StatusBar = "Applying formatting..."
    
    ' Format Raw Code table
    FormatRawTable tblRawCode
    
    ' Format Raw LEI table
    FormatRawTable tblRawLEI
    
    ' Format Raw Name table
    FormatRawTable tblRawName
    
    ' Format Upload table
    FormatUploadTable tblUpload
    
    ' Auto-fit columns
    wsRawDataCode.Columns.AutoFit
    wsRawDataLEI.Columns.AutoFit
    wsRawDataName.Columns.AutoFit
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
           "Fund Code matches found: " & (rawCodeRowCount - 1) & vbCrLf & _
           "Fund LEI matches found: " & (rawLEIRowCount - 1) & vbCrLf & _
           "Fund Name matches found: " & (rawNameRowCount - 1) & vbCrLf & _
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

' Helper function to populate a raw array with data from approved and markit arrays
Sub PopulateRawArray(rawArray As Variant, rawRowCount As Long, approvedArray As Variant, markitArray As Variant, _
                    approvedRow As Long, markitRow As Long, approvedColMap As Object, markitColMap As Object)
    
    ' 1. Business Unit from Approved
    If approvedColMap.Exists("Business Unit") Then
        rawArray(rawRowCount, 1) = approvedArray(approvedRow, approvedColMap("Business Unit"))
    End If
    
    ' 2. IA GCI from Approved
    If approvedColMap.Exists("IA GCI") Then
        rawArray(rawRowCount, 2) = approvedArray(approvedRow, approvedColMap("IA GCI"))
    End If
    
    ' 3. RFAD Investment Manager from Approved
    If approvedColMap.Exists("Investment Manager") Then
        rawArray(rawRowCount, 3) = approvedArray(approvedRow, approvedColMap("Investment Manager"))
    End If
    
    ' 4. Markit Investment Manager from Markit
    If markitColMap.Exists("Investment Manager") Then
        rawArray(rawRowCount, 4) = markitArray(markitRow, markitColMap("Investment Manager"))
    End If
    
    ' 5. Fund GCI from Approved
    If approvedColMap.Exists("Fund GCI") Then
        rawArray(rawRowCount, 5) = approvedArray(approvedRow, approvedColMap("Fund GCI"))
    End If
    
    ' 6. RFAD Fund Name from Approved
    If approvedColMap.Exists("Fund Name") Then
        rawArray(rawRowCount, 6) = approvedArray(approvedRow, approvedColMap("Fund Name"))
    End If
    
    ' 7. Markit Fund Name from Markit
    If markitColMap.Exists("Full Legal Fund Name") Then
        rawArray(rawRowCount, 7) = markitArray(markitRow, markitColMap("Full Legal Fund Name"))
    End If
    
    ' 8. Fund LEI from Approved
    If approvedColMap.Exists("Fund LEI") Then
        rawArray(rawRowCount, 8) = approvedArray(approvedRow, approvedColMap("Fund LEI"))
    End If
    
    ' 9. Fund Code from Approved
    If approvedColMap.Exists("Fund Code") Then
        rawArray(rawRowCount, 9) = approvedArray(approvedRow, approvedColMap("Fund Code"))
    End If
    
    ' 10. RFAD Currency from Approved
    If approvedColMap.Exists("Currency") Then
        rawArray(rawRowCount, 10) = approvedArray(approvedRow, approvedColMap("Currency"))
    End If
    
    ' 11. Markit Currency from Markit
    If markitColMap.Exists("Currency") Then
        rawArray(rawRowCount, 11) = markitArray(markitRow, markitColMap("Currency"))
    End If
    
    ' 12. RFAD Latest NAV Date from Approved
    If approvedColMap.Exists("Latest NAV Date") Then
        rawArray(rawRowCount, 12) = approvedArray(approvedRow, approvedColMap("Latest NAV Date"))
    End If
    
    ' 13. RFAD Latest NAV from Approved
    If approvedColMap.Exists("Latest NAV") Then
        rawArray(rawRowCount, 13) = approvedArray(approvedRow, approvedColMap("Latest NAV"))
    End If
    
    ' 14. Markit Latest NAV Date from Markit
    If markitColMap.Exists("As of Date") Then
        rawArray(rawRowCount, 14) = markitArray(markitRow, markitColMap("As of Date"))
    End If
    
    ' 15. Markit Latest NAV from Markit
    If markitColMap.Exists("AUM Value") Then
        rawArray(rawRowCount, 15) = markitArray(markitRow, markitColMap("AUM Value"))
    End If
End Sub

' Helper function to process raw data arrays for upload table
Sub ProcessRawDataForUpload(rawArray As Variant, rawRowCount As Long, uploadArray As Variant, ByRef uploadRowCount As Long)
    Dim i As Long, j As Long, k As Long
    Dim tempUploadArray() As Variant
    Dim rfadNAVDate As Variant
    Dim markitNAVDate As Variant
    Dim isMoreRecentBy15Days As Boolean
    Dim daysDifference As Long
    Dim rfadNAV As Variant
    Dim markitNAV As Variant
    Dim deltaValue As Double
    
    ' Loop through Raw data (skip header row)
    For i = 2 To rawRowCount
        ' Get dates
        rfadNAVDate = rawArray(i, 12)
        markitNAVDate = rawArray(i, 14)
        
        ' Check if Markit date is at least 15 days more recent
        isMoreRecentBy15Days = False
        
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
                ' Calculate days difference
                daysDifference = DateDiff("d", CDate(rfadNAVDate), CDate(markitNAVDate))
                isMoreRecentBy15Days = (daysDifference >= 15)
            End If
        End If
        
        ' If Markit date is at least 15 days more recent, add to Upload
        If isMoreRecentBy15Days Then
            uploadRowCount = uploadRowCount + 1
            
            ' If we need to resize the array
            If uploadRowCount > UBound(uploadArray, 1) Then
                ' Create a new larger array
                ReDim tempUploadArray(1 To uploadRowCount * 2, 1 To UBound(uploadArray, 2))
                
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
            For j = 1 To UBound(rawArray, 2)
                uploadArray(uploadRowCount, j) = rawArray(i, j)
            Next j
            
            ' Calculate Delta with Markit NAV in millions
            rfadNAV = rawArray(i, 13)
            markitNAV = rawArray(i, 15)
            
            ' Calculate Delta if both NAVs are valid numbers
            If IsNumeric(rfadNAV) And IsNumeric(markitNAV) And CDbl(rfadNAV) <> 0 Then
                ' Convert Markit NAV to millions - RFAD NAV is already in millions
                markitNAV = CDbl(markitNAV) / 1000000
                deltaValue = (markitNAV / CDbl(rfadNAV)) - 1
                uploadArray(uploadRowCount, UBound(uploadArray, 2)) = deltaValue
            End If
        End If
    Next i
End Sub

' Helper function to format Raw tables
Sub FormatRawTable(tbl As ListObject)
    Dim i As Long
    
    On Error Resume Next
    If Not tbl Is Nothing Then
        With tbl
            If .ListColumns.Count >= 12 Then .ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' RFAD Latest NAV Date
            If .ListColumns.Count >= 14 Then .ListColumns(14).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' Markit Latest NAV Date
            If .ListColumns.Count >= 13 Then .ListColumns(13).DataBodyRange.NumberFormat = "#,##0.00" ' RFAD Latest NAV
            If .ListColumns.Count >= 15 Then .ListColumns(15).DataBodyRange.NumberFormat = "#,##0.00" ' Markit Latest NAV
            
            ' Add currency mismatch highlighting
            If .ListColumns.Count >= 10 And .ListColumns.Count >= 11 Then
                ' Highlight Markit Currency if it doesn't match RFAD Currency
                For i = 1 To .ListRows.Count
                    Dim cellValue As Variant
                    Dim rfadCurrency As String
                    Dim markitCurrency As String
                    Dim rfadCurrency As String
                    Dim markitCurrency As String
                    
                    rfadCurrency = Trim(CStr(.ListRows(i).Range.Cells(1, 10).Value))
                    markitCurrency = Trim(CStr(.ListRows(i).Range.Cells(1, 11).Value))
                    
                    If rfadCurrency <> "" And markitCurrency <> "" And rfadCurrency <> markitCurrency Then
                        .ListRows(i).Range.Cells(1, 11).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                    End If
                Next i
            End If
        End With
    End If
    On Error GoTo 0
End Sub

' Helper function to format Upload table
Sub FormatUploadTable(tbl As ListObject)
    Dim i As Long
    
    On Error Resume Next
    If Not tbl Is Nothing Then
        With tbl
            If .ListColumns.Count >= 12 Then .ListColumns(12).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' RFAD Latest NAV Date
            If .ListColumns.Count >= 14 Then .ListColumns(14).DataBodyRange.NumberFormat = "yyyy-mm-dd" ' Markit Latest NAV Date
            If .ListColumns.Count >= 13 Then .ListColumns(13).DataBodyRange.NumberFormat = "#,##0.00" ' RFAD Latest NAV
            If .ListColumns.Count >= 15 Then .ListColumns(15).DataBodyRange.NumberFormat = "#,##0.00" ' Markit Latest NAV
            If .ListColumns.Count >= 16 Then .ListColumns(16).DataBodyRange.NumberFormat = "0.00%" ' Delta
            
            ' Add currency mismatch highlighting
            If .ListColumns.Count >= 10 And .ListColumns.Count >= 11 Then
                ' Highlight Markit Currency if it doesn't match RFAD Currency
                For i = 1 To .ListRows.Count
                    Dim rfadCurrency As String
                    Dim markitCurrency As String
                    
                    rfadCurrency = Trim(CStr(.ListRows(i).Range.Cells(1, 10).Value))
                    markitCurrency = Trim(CStr(.ListRows(i).Range.Cells(1, 11).Value))
                    
                    If rfadCurrency <> "" And markitCurrency <> "" And rfadCurrency <> markitCurrency Then
                        .ListRows(i).Range.Cells(1, 11).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                    End If
                Next i
            End If
            
            ' Highlight cells in Delta column if value >= 100% or <= -50%
            If Not .DataBodyRange Is Nothing Then
                For i = 1 To .ListRows.Count
                    cellValue = .ListRows(i).Range.Cells(1, .ListColumns.Count).Value
                    
                    If IsNumeric(cellValue) Then
                        If cellValue >= 1 Or cellValue <= -0.5 Then
                            .ListRows(i).Range.Cells(1, .ListColumns.Count).Interior.Color = RGB(255, 0, 0)
                        End If
                    End If
                Next i
            End If
        End With
    End If
    On Error GoTo 0
End Sub
