Sub ProcessMarkitAndApprovedFunds()
    ' Declare variables
    Dim wbMaster As Workbook
    Dim wbApproved As Workbook
    Dim wbMarkit As Workbook
    Dim wsApproved As Worksheet
    Dim wsMarkit As Worksheet
    Dim wsRawData As Worksheet
    Dim wsUpload As Worksheet
    Dim approvedFilePath As String
    Dim markitFilePath As String
    Dim tblApproved As ListObject
    Dim tblMarkit As ListObject
    Dim tblRaw As ListObject
    Dim tblUpload As ListObject
    Dim lastRowApproved As Long
    Dim lastColApproved As Long
    Dim lastRowMarkit As Long
    Dim lastColMarkit As Long
    Dim approvedData As Range
    Dim markitData As Range
    Dim rawDataHeaders() As String
    Dim foundMatch As Boolean
    Dim i As Long, j As Long
    Dim matchFundCode As Boolean
    Dim matchFundLEI As Boolean
    Dim existingListObj As ListObject
    
    ' Turn off screen updating and calculations to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set reference to the master workbook (current workbook)
    Set wbMaster = ThisWorkbook
    
    ' Ask user to locate the Approved funds file
    MsgBox "Please select the Approved funds file.", vbInformation, "Select Approved Funds File"
    approvedFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Approved Funds File")
    
    ' Check if user canceled the file selection
    If approvedFilePath = "False" Then
        MsgBox "Operation canceled by user.", vbExclamation
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Ask user to locate the Markit file
    MsgBox "Please select the Markit file.", vbInformation, "Select Markit File"
    markitFilePath = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select Markit File")
    
    ' Check if user canceled the file selection
    If markitFilePath = "False" Then
        MsgBox "Operation canceled by user.", vbExclamation
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Open the selected files
    Set wbApproved = Workbooks.Open(approvedFilePath)
    Set wbMarkit = Workbooks.Open(markitFilePath)
    
    ' Process Approved funds file
    ' Assume data is in the first worksheet
    Set wsApproved = wbApproved.Sheets(1)
    
    ' Delete the first row
    wsApproved.Rows(1).Delete
    
    ' Check if the file already contains tables (ListObjects)
    If wsApproved.ListObjects.Count > 0 Then
        ' The file already has tables - we'll use them instead of creating new ones
        Set tblApproved = wsApproved.ListObjects(1)  ' Use the first ListObject in the sheet
        
        ' If the table already has a name, capture it for reference
        If tblApproved.Name <> "" Then
            Debug.Print "Using existing table: " & tblApproved.Name
        Else
            ' Try to rename the table to our standard name
            On Error Resume Next
            tblApproved.Name = "Approved"
            On Error GoTo 0
        End If
    Else
        ' No existing tables, so create one
        lastRowApproved = wsApproved.Cells(wsApproved.Rows.Count, "A").End(xlUp).Row
        lastColApproved = wsApproved.Cells(1, wsApproved.Columns.Count).End(xlToLeft).Column
        Set approvedData = wsApproved.Range(wsApproved.Cells(1, 1), wsApproved.Cells(lastRowApproved, lastColApproved))
        
        ' Create a table with error handling - use specific error handling for table overlap
        On Error Resume Next
        Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes)
        
        If Err.Number = 1004 Then  ' Table overlap error
            ' Try to clear any existing tables that might not be properly detected
            Err.Clear
            
            ' Convert any existing tables to ranges first
            For i = wsApproved.ListObjects.Count To 1 Step -1
                wsApproved.ListObjects(i).Unlist
            Next i
            
            ' Try again with the clean sheet
            Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes)
            
            If Err.Number <> 0 Then
                MsgBox "Error: The Approved file already contains tables that couldn't be processed. " & _
                       "Please manually convert tables to ranges in the file before running this macro.", vbExclamation
                On Error GoTo 0
                GoTo CleanupAndExit
            End If
        ElseIf Err.Number <> 0 Then
            MsgBox "Error creating Approved table: " & Err.Description, vbExclamation
            On Error GoTo 0
            GoTo CleanupAndExit
        End If
        On Error GoTo 0
        
        ' Ensure the table is named correctly
        tblApproved.Name = "Approved"
    End If
    
    ' Filter to keep only 'FI-EMEA', 'FI-US', and 'FI-GMC-ASIA' in 'Business Unit' column
    ' Find the index of "Business Unit" column
    Dim buColIndex As Long
    buColIndex = 0
    
    For i = 1 To tblApproved.ListColumns.Count
        If tblApproved.ListColumns(i).Name = "Business Unit" Then
            buColIndex = i
            Exit For
        End If
    Next i
    
    ' Apply filter if Business Unit column exists
    If buColIndex > 0 Then
        tblApproved.Range.AutoFilter Field:=buColIndex, Criteria1:=Array("FI-EMEA", "FI-US", "FI-GMC-ASIA"), Operator:=xlFilterValues
    Else
        MsgBox "Warning: 'Business Unit' column not found in Approved funds file.", vbExclamation
    End If
    
    ' Process Markit file
    ' Assume data is in the first worksheet
    Set wsMarkit = wbMarkit.Sheets(1)
    
    ' Check if the file already contains tables (ListObjects)
    If wsMarkit.ListObjects.Count > 0 Then
        ' The file already has tables - we'll use them instead of creating new ones
        Set tblMarkit = wsMarkit.ListObjects(1)  ' Use the first ListObject in the sheet
        
        ' If the table already has a name, capture it for reference
        If tblMarkit.Name <> "" Then
            Debug.Print "Using existing table: " & tblMarkit.Name
        Else
            ' Try to rename the table to our standard name
            On Error Resume Next
            tblMarkit.Name = "Markit"
            On Error GoTo 0
        End If
    Else
        ' No existing tables, so create one
        lastRowMarkit = wsMarkit.Cells(wsMarkit.Rows.Count, "A").End(xlUp).Row
        lastColMarkit = wsMarkit.Cells(1, wsMarkit.Columns.Count).End(xlToLeft).Column
        
        ' Make sure range is valid
        If lastRowMarkit < 1 Or lastColMarkit < 1 Then
            MsgBox "Warning: Could not determine data range in Markit file", vbExclamation
            GoTo CleanupAndExit
        End If
        
        Set markitData = wsMarkit.Range(wsMarkit.Cells(1, 1), wsMarkit.Cells(lastRowMarkit, lastColMarkit))
        
        ' Create a table with error handling - use specific error handling for table overlap
        On Error Resume Next
        Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes)
        
        If Err.Number = 1004 Then  ' Table overlap error
            ' Try to clear any existing tables that might not be properly detected
            Err.Clear
            
            ' Convert any existing tables to ranges first
            For i = wsMarkit.ListObjects.Count To 1 Step -1
                wsMarkit.ListObjects(i).Unlist
            Next i
            
            ' Try again with the clean sheet
            Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes)
            
            If Err.Number <> 0 Then
                MsgBox "Error: The Markit file already contains tables that couldn't be processed. " & _
                       "Please manually convert tables to ranges in the file before running this macro.", vbExclamation
                On Error GoTo 0
                GoTo CleanupAndExit
            End If
        ElseIf Err.Number <> 0 Then
            MsgBox "Error creating Markit table: " & Err.Description, vbExclamation
            On Error GoTo 0
            GoTo CleanupAndExit
        End If
        On Error GoTo 0
        
        ' Ensure the table is named correctly
        tblMarkit.Name = "Markit"
    End If
    tblMarkit.Name = "Markit"
    
    ' Create new sheets in master workbook for the tables
    On Error Resume Next
    Application.DisplayAlerts = False
    wbMaster.Sheets("Approved").Delete
    wbMaster.Sheets("Markit").Delete
    wbMaster.Sheets("Raw_data").Delete
    wbMaster.Sheets("Markit NAV today date").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add new clean sheets
    Set wsApproved = wbMaster.Sheets.Add(After:=wbMaster.Sheets(wbMaster.Sheets.Count))
    wsApproved.Name = "Approved"
    
    Set wsMarkit = wbMaster.Sheets.Add(After:=wsApproved)
    wsMarkit.Name = "Markit"
    
    ' Copy the data from the source workbooks to the master workbook
    ' Important: Copy only the visible cells from the filtered Approved table
    If tblApproved.ShowAutoFilter Then
        tblApproved.Range.SpecialCells(xlCellTypeVisible).Copy wsApproved.Range("A1")
    Else
        tblApproved.Range.Copy wsApproved.Range("A1")
    End If
    
    tblMarkit.Range.Copy wsMarkit.Range("A1")
    
    ' Get the dimensions of the copied data
    lastRowApproved = wsApproved.Cells(wsApproved.Rows.Count, "A").End(xlUp).Row
    lastColApproved = wsApproved.Cells(1, wsApproved.Columns.Count).End(xlToLeft).Column
    
    lastRowMarkit = wsMarkit.Cells(wsMarkit.Rows.Count, "A").End(xlUp).Row
    lastColMarkit = wsMarkit.Cells(1, wsMarkit.Columns.Count).End(xlToLeft).Column
    
    ' Create ranges for the copied data
    If lastRowApproved > 0 And lastColApproved > 0 Then
        Set approvedData = wsApproved.Range(wsApproved.Cells(1, 1), wsApproved.Cells(lastRowApproved, lastColApproved))
    Else
        MsgBox "No data found in Approved sheet", vbExclamation
        GoTo CleanupAndExit
    End If
    
    If lastRowMarkit > 0 And lastColMarkit > 0 Then
        Set markitData = wsMarkit.Range(wsMarkit.Cells(1, 1), wsMarkit.Cells(lastRowMarkit, lastColMarkit))
    Else
        MsgBox "No data found in Markit sheet", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' Make sure there are no existing tables before creating new ones
    ' This is critical for preventing the "table can't overlap" error
    For Each existingListObj In wsApproved.ListObjects
        existingListObj.Unlist
    Next existingListObj
    
    For Each existingListObj In wsMarkit.ListObjects
        existingListObj.Unlist
    Next existingListObj
    
    ' Now create the tables safely
    On Error Resume Next
    Set tblApproved = wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes)
    If Err.Number <> 0 Then
        Err.Clear
        ' Alternative approach - create table directly with worksheet data
        wsApproved.ListObjects.Add(xlSrcRange, approvedData, , xlYes).Name = "Approved"
        Set tblApproved = wsApproved.ListObjects("Approved")
    Else
        tblApproved.Name = "Approved"
    End If
    
    Set tblMarkit = wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes)
    If Err.Number <> 0 Then
        Err.Clear
        ' Alternative approach - create table directly with worksheet data
        wsMarkit.ListObjects.Add(xlSrcRange, markitData, , xlYes).Name = "Markit"
        Set tblMarkit = wsMarkit.ListObjects("Markit")
    Else
        tblMarkit.Name = "Markit"
    End If
    On Error GoTo 0
    
    ' Close the original files
    wbApproved.Close SaveChanges:=False
    wbMarkit.Close SaveChanges:=False
    
    ' Create Raw_data sheet with Raw table
    Set wsRawData = wbMaster.Sheets.Add(After:=wsMarkit)
    wsRawData.Name = "Raw_data"
    
    ' Define columns for Raw table
    rawDataHeaders = Array("Business Unit", "IA GCI", "RFAD Investment Manager", "Markit Investment Manager", _
                          "Fund GCI", "RFAD Fund Name", "Markit Fund Name", "Fund LEI", "Fund Code", _
                          "RFAD Currency", "Markit Currency", "RFAD Latest NAV Date", "RFAD Latest NAV", _
                          "Markit Latest NAV Date", "Markit Latest NAV")
    
    ' Create headers for Raw table
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        wsRawData.Cells(1, i + 1).Value = rawDataHeaders(i)
    Next i
    
    ' Create Raw table (initially empty except for headers)
    Set tblRaw = wsRawData.ListObjects.Add(xlSrcRange, wsRawData.Range("A1").Resize(1, UBound(rawDataHeaders) - LBound(rawDataHeaders) + 1), , xlYes)
    tblRaw.Name = "Raw"
    
    ' Create mapping of column names to indexes for both tables
    Dim approvedColMap As Object, markitColMap As Object
    Set approvedColMap = CreateObject("Scripting.Dictionary")
    Set markitColMap = CreateObject("Scripting.Dictionary")
    
    ' Populate dictionaries with column names and indexes
    For i = 1 To tblApproved.ListColumns.Count
        approvedColMap.Add tblApproved.ListColumns(i).Name, i
    Next i
    
    For i = 1 To tblMarkit.ListColumns.Count
        markitColMap.Add tblMarkit.ListColumns(i).Name, i
    Next i
    
    ' Loop through each row in the Approved table
    For i = 1 To tblApproved.ListRows.Count
        foundMatch = False
        
        ' Get fund code and LEI from Approved table
        Dim approvedFundCode As String, approvedFundLEI As String, fundGCI As String
        
        ' Safely get values with error handling
        On Error Resume Next
        If approvedColMap.Exists("Fund Code") Then
            approvedFundCode = tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund Code")).Value
        End If
        
        If approvedColMap.Exists("Fund LEI") Then
            approvedFundLEI = tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund LEI")).Value
        End If
        
        If approvedColMap.Exists("Fund GCI") Then
            fundGCI = tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund GCI")).Value
        Else
            ' Skip this row if Fund GCI doesn't exist
            GoTo NextApprovedRow
        End If
        On Error GoTo 0
        
        ' Skip if both Fund Code and Fund LEI are empty
        If Len(Trim(CStr(approvedFundCode))) = 0 And Len(Trim(CStr(approvedFundLEI))) = 0 Then
            GoTo NextApprovedRow
        End If
        
        ' Look for matching row in Markit table
        For j = 1 To tblMarkit.ListRows.Count
            matchFundCode = False
            matchFundLEI = False
            
            ' Check if Fund Code matches
            If Len(Trim(CStr(approvedFundCode))) > 0 And markitColMap.Exists("Client Identifier") Then
                If CStr(tblMarkit.ListRows(j).Range.Cells(1, markitColMap("Client Identifier")).Value) = CStr(approvedFundCode) Then
                    matchFundCode = True
                End If
            End If
            
            ' If Fund Code doesn't match, check Fund LEI
            If Not matchFundCode And Len(Trim(CStr(approvedFundLEI))) > 0 And markitColMap.Exists("LEI") Then
                If CStr(tblMarkit.ListRows(j).Range.Cells(1, markitColMap("LEI")).Value) = CStr(approvedFundLEI) Then
                    matchFundLEI = True
                End If
            End If
            
            ' If either Fund Code or Fund LEI matches, populate Raw table
            If matchFundCode Or matchFundLEI Then
                foundMatch = True
                
                ' Add new row to Raw table
                tblRaw.ListRows.Add
                
                ' Populate all columns in Raw table based on mapping
                ' Business Unit from Approved
                If approvedColMap.Exists("Business Unit") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 1).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Business Unit")).Value
                End If
                
                ' IA GCI from Approved
                If approvedColMap.Exists("IA GCI") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 2).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("IA GCI")).Value
                End If
                
                ' RFAD Investment Manager from Approved (Investment Manager column)
                If approvedColMap.Exists("Investment Manager") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 3).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Investment Manager")).Value
                End If
                
                ' Markit Investment Manager from Markit
                If markitColMap.Exists("Investment Manager") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 4).Value = _
                        tblMarkit.ListRows(j).Range.Cells(1, markitColMap("Investment Manager")).Value
                End If
                
                ' Fund GCI from Approved
                tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 5).Value = fundGCI
                
                ' RFAD Fund Name from Approved (Fund Name column)
                If approvedColMap.Exists("Fund Name") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 6).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund Name")).Value
                End If
                
                ' Markit Fund Name from Markit (Full Legal Fund Name column)
                If markitColMap.Exists("Full Legal Fund Name") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 7).Value = _
                        tblMarkit.ListRows(j).Range.Cells(1, markitColMap("Full Legal Fund Name")).Value
                End If
                
                ' Fund LEI from Approved
                If approvedColMap.Exists("Fund LEI") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 8).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund LEI")).Value
                End If
                
                ' Fund Code from Approved
                If approvedColMap.Exists("Fund Code") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 9).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Fund Code")).Value
                End If
                
                ' RFAD Currency from Approved (Currency column)
                If approvedColMap.Exists("Currency") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 10).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Currency")).Value
                End If
                
                ' Markit Currency from Markit
                If markitColMap.Exists("Currency") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 11).Value = _
                        tblMarkit.ListRows(j).Range.Cells(1, markitColMap("Currency")).Value
                End If
                
                ' RFAD Latest NAV Date from Approved
                If approvedColMap.Exists("Latest NAV Date") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 12).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Latest NAV Date")).Value
                End If
                
                ' RFAD Latest NAV from Approved
                If approvedColMap.Exists("Latest NAV") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 13).Value = _
                        tblApproved.ListRows(i).Range.Cells(1, approvedColMap("Latest NAV")).Value
                End If
                
                ' Markit Latest NAV Date from Markit (As of Date column)
                If markitColMap.Exists("As of Date") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 14).Value = _
                        tblMarkit.ListRows(j).Range.Cells(1, markitColMap("As of Date")).Value
                End If
                
                ' Markit Latest NAV from Markit (AUM Value column)
                If markitColMap.Exists("AUM Value") Then
                    tblRaw.ListRows(tblRaw.ListRows.Count).Range.Cells(1, 15).Value = _
                        tblMarkit.ListRows(j).Range.Cells(1, markitColMap("AUM Value")).Value
                End If
                
                ' Only use the first match (either by Fund Code or Fund LEI)
                Exit For
            End If
        Next j
NextApprovedRow:
    Next i
    
    ' Create "Markit NAV today date" sheet with upload table
    Set wsUpload = wbMaster.Sheets.Add(After:=wsRawData)
    wsUpload.Name = "Markit NAV today date"
    
    ' Create upload table with same columns as Raw table + Delta column
    For i = LBound(rawDataHeaders) To UBound(rawDataHeaders)
        wsUpload.Cells(1, i + 1).Value = rawDataHeaders(i)
    Next i
    wsUpload.Cells(1, UBound(rawDataHeaders) + 2).Value = "Delta"
    
    Set tblUpload = wsUpload.ListObjects.Add(xlSrcRange, wsUpload.Range("A1").Resize(1, UBound(rawDataHeaders) - LBound(rawDataHeaders) + 2), , xlYes)
    tblUpload.Name = "Upload"
    
    ' Populate upload table with rows from Raw where Markit Latest NAV Date > RFAD Latest NAV Date
    Dim uploadRow As Long
    uploadRow = 2 ' Start populating from row 2 (after headers)
    
    ' Loop through each row in Raw table
    For i = 1 To tblRaw.ListRows.Count
        Dim rfadNAVDate As Variant, markitNAVDate As Variant
        Dim isMoreRecent As Boolean
        
        ' Get dates from Raw table
        rfadNAVDate = tblRaw.ListRows(i).Range.Cells(1, 12).Value
        markitNAVDate = tblRaw.ListRows(i).Range.Cells(1, 14).Value
        
        ' Check if both dates are valid
        isMoreRecent = False
        
        ' Only compare if both dates are not empty
        If Not IsEmpty(rfadNAVDate) And Not IsEmpty(markitNAVDate) Then
            ' Try to convert to dates if they're not already
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
            
            ' Compare dates if both are valid dates
            If IsDate(rfadNAVDate) And IsDate(markitNAVDate) Then
                isMoreRecent = (CDate(markitNAVDate) > CDate(rfadNAVDate))
            End If
        End If
        
        ' If Markit date is more recent, add to Upload table
        If isMoreRecent Then
            ' Add new row to Upload table
            tblUpload.ListRows.Add
            
            ' Copy all columns from Raw to Upload
            For j = 1 To UBound(rawDataHeaders) + 1
                tblUpload.ListRows(tblUpload.ListRows.Count).Range.Cells(1, j).Value = _
                    tblRaw.ListRows(i).Range.Cells(1, j).Value
            Next j
            
            ' Add Delta column calculation
            Dim rfadNAV As Variant, markitNAV As Variant
            Dim deltaValue As Double
            
            rfadNAV = tblRaw.ListRows(i).Range.Cells(1, 13).Value
            markitNAV = tblRaw.ListRows(i).Range.Cells(1, 15).Value
            
            ' Calculate Delta only if both NAVs are valid numbers
            If IsNumeric(rfadNAV) And IsNumeric(markitNAV) And CDbl(rfadNAV) <> 0 Then
                ' Calculate Delta and format as percentage
                deltaValue = (CDbl(markitNAV) / CDbl(rfadNAV)) - 1
                
                tblUpload.ListRows(tblUpload.ListRows.Count).Range.Cells(1, UBound(rawDataHeaders) + 2).Value = deltaValue
                tblUpload.ListRows(tblUpload.ListRows.Count).Range.Cells(1, UBound(rawDataHeaders) + 2).NumberFormat = "0.00%"
                
                ' Highlight cell if value >= 100% or <= -50%
                If deltaValue >= 1 Or deltaValue <= -0.5 Then
                    tblUpload.ListRows(tblUpload.ListRows.Count).Range.Cells(1, UBound(rawDataHeaders) + 2).Interior.Color = RGB(255, 0, 0) ' Red for highlighting
                End If
            End If
        End If
    Next i
    
    ' Apply proper formatting to date columns
    On Error Resume Next
    tblRaw.ListColumns("RFAD Latest NAV Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    tblRaw.ListColumns("Markit Latest NAV Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    tblUpload.ListColumns("RFAD Latest NAV Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    tblUpload.ListColumns("Markit Latest NAV Date").DataBodyRange.NumberFormat = "yyyy-mm-dd"
    
    ' Format NAV columns as currency
    tblRaw.ListColumns("RFAD Latest NAV").DataBodyRange.NumberFormat = "#,##0.00"
    tblRaw.ListColumns("Markit Latest NAV").DataBodyRange.NumberFormat = "#,##0.00"
    tblUpload.ListColumns("RFAD Latest NAV").DataBodyRange.NumberFormat = "#,##0.00"
    tblUpload.ListColumns("Markit Latest NAV").DataBodyRange.NumberFormat = "#,##0.00"
    On Error GoTo 0
    
    ' Auto-fit columns for better readability
    wsRawData.Columns.AutoFit
    wsUpload.Columns.AutoFit
    
    ' Activate the first sheet for a clean finish
    wbMaster.Sheets(1).Activate
    
    ' Turn screen updating and calculations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Processing completed successfully!", vbInformation
    Exit Sub
    
CleanupAndExit:
    ' Turn screen updating and calculations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

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
    
End Sub
