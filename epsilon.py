Option Explicit

Sub MacroEpsilon()

    '---------------------------------------------------------------------------------
    ' 1. DECLARE & INITIALIZE
    '---------------------------------------------------------------------------------
    Dim wsPortfolio As Worksheet, wsIA As Worksheet
    Dim portfolioTable As ListObject, iaTable As ListObject
    
    Dim wbPrevious As Workbook, wsIAPrev As Worksheet
    Dim iaTablePrev As ListObject
    
    Dim uniqueGCI As Object
    Dim regionData As Object, managerData As Object, triggerData As Object
    Dim navSourceData As Object, clientContactData As Object
    Dim triggerCountData As Object, nonTriggerCountData As Object
    Dim missingTriggerData As Object, missingNonTriggerData As Object
    Dim manualColumns As Object
    
    Dim portfolioData As Variant, manualData As Variant, iaData() As Variant
    Dim userChoice As VbMsgBoxResult
    Dim wantManualData As Boolean
    
    Dim previousFile As Variant  ' for GetOpenFilename (can return False)
    Dim numRowsPortfolio As Long, numRowsIAPrev As Long
    Dim numUniqueGCI As Long
    
    Dim i As Long, j As Long
    Dim gci As Variant
    
    '---------------------------------------------------------------------------------
    ' 2. OPTIMIZE APPLICATION SETTINGS
    '---------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '---------------------------------------------------------------------------------
    ' 3. SET REFERENCES TO WORKSHEETS
    '---------------------------------------------------------------------------------
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsIA = ThisWorkbook.Sheets("IA_Level")
    
    '---------------------------------------------------------------------------------
    ' 4. GET PORTFOLIOTABLE & DO GUARD CHECKS
    '---------------------------------------------------------------------------------
    Dim msg As String
    
    Set portfolioTable = Nothing
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    
    If portfolioTable Is Nothing Then
        MsgBox "Table 'PortfolioTable' not found on the 'Portfolio' sheet." & vbCrLf & _
               "Please verify the table name/spelling.", vbCritical
        GoTo CleanUp
    End If
    
    ' Ensure PortfolioTable has data rows
    If portfolioTable.DataBodyRange Is Nothing Then
        MsgBox "No data found in 'PortfolioTable'. Please ensure it has at least one row.", vbExclamation
        GoTo CleanUp
    End If
    
    '----- Check that each required column exists in PortfolioTable -----
    Call CheckColumnExists(portfolioTable, "Fund Manager GCI")
    Call CheckColumnExists(portfolioTable, "Region")
    Call CheckColumnExists(portfolioTable, "Fund Manager")
    Call CheckColumnExists(portfolioTable, "Trigger/Non-Trigger")
    Call CheckColumnExists(portfolioTable, "NAV Source")
    Call CheckColumnExists(portfolioTable, "Primary Client Contact")
    Call CheckColumnExists(portfolioTable, "Secondary Client Contact")
    Call CheckColumnExists(portfolioTable, "Wks Missing")
    '----- If the code never stopped, it means all columns exist. -----
    
    '---------------------------------------------------------------------------------
    ' 5. CHECK IF IA_TABLE EXISTS; IF NOT, CREATE IT
    '---------------------------------------------------------------------------------
    Set iaTable = Nothing
    On Error Resume Next
    Set iaTable = wsIA.ListObjects("IA_Table")
    On Error GoTo 0
    
    If iaTable Is Nothing Then
        ' Create required header row in row 1 (columns A:T = 20 columns)
        Dim requiredCols As Variant
        requiredCols = Array( _
            "Fund Manager GCI", "Region", "Fund Manager", "Trigger/Non-Trigger", "NAV Source", _
            "Client Contact(s)", "Trigger", "Non-Trigger", "Total Funds", "Missing Trigger", _
            "Missing Non-Trigger", "Total Missing", "Days to Report", "1st Client Outreach Date", _
            "2nd Client Outreach Date", "OA Escalation Date", "NOA Escalation Date", _
            "Escalation Name", "Final Status", "Comments")
        
        Dim col As Long
        For col = LBound(requiredCols) To UBound(requiredCols)
            wsIA.Cells(1, col + 1).Value = requiredCols(col)
        Next col
        
        ' Convert A1:T1 into a table
        Set iaTable = wsIA.ListObjects.Add( _
                            SourceType:=xlSrcRange, _
                            Source:=wsIA.Range("A1:T1"), _
                            XlListObjectHasHeaders:=xlYes)
        iaTable.Name = "IA_Table"
    End If
    
    ' Clear IA_Table except headers
    If Not iaTable.DataBodyRange Is Nothing Then
        iaTable.DataBodyRange.Delete
    End If
    
    '---------------------------------------------------------------------------------
    ' 6. PROMPT USER FOR MANUAL DATA IMPORT
    '---------------------------------------------------------------------------------
    userChoice = MsgBox( _
        "Would you like to load manual data from a previous version of IA_Table?", _
        vbYesNo + vbQuestion, _
        "Load Manual Data?")
    wantManualData = (userChoice = vbYes)
    
    '---------------------------------------------------------------------------------
    ' 7. INITIALIZE DICTIONARIES
    '---------------------------------------------------------------------------------
    Set uniqueGCI = CreateObject("Scripting.Dictionary")
    Set regionData = CreateObject("Scripting.Dictionary")
    Set managerData = CreateObject("Scripting.Dictionary")
    Set triggerData = CreateObject("Scripting.Dictionary")
    Set navSourceData = CreateObject("Scripting.Dictionary")
    Set clientContactData = CreateObject("Scripting.Dictionary")
    Set triggerCountData = CreateObject("Scripting.Dictionary")
    Set nonTriggerCountData = CreateObject("Scripting.Dictionary")
    Set missingTriggerData = CreateObject("Scripting.Dictionary")
    Set missingNonTriggerData = CreateObject("Scripting.Dictionary")
    Set manualColumns = CreateObject("Scripting.Dictionary")
    
    '---------------------------------------------------------------------------------
    ' 8. IF USER WANTS MANUAL DATA, OPEN PREVIOUS WORKBOOK
    '---------------------------------------------------------------------------------
    If wantManualData Then
        
        previousFile = Application.GetOpenFilename( _
            "Excel Files (*.xls*), *.xls*", , "Select the previous version of IA_Table")
        
        If VarType(previousFile) = vbBoolean And previousFile = False Then
            MsgBox "No file selected. Manual columns will remain blank.", vbExclamation
            wantManualData = False
        Else
            ' Open the previous workbook
            Set wbPrevious = Workbooks.Open(CStr(previousFile))
            
            ' Attempt to set references
            Set wsIAPrev = Nothing
            On Error Resume Next
            Set wsIAPrev = wbPrevious.Sheets("IA_Level")
            On Error GoTo 0
            
            If wsIAPrev Is Nothing Then
                MsgBox "Sheet 'IA_Level' not found in the previous file. Manual columns will remain blank.", vbCritical
                wbPrevious.Close SaveChanges:=False
                wantManualData = False
            Else
                Set iaTablePrev = Nothing
                On Error Resume Next
                Set iaTablePrev = wsIAPrev.ListObjects("IA_Table")
                On Error GoTo 0
                
                If iaTablePrev Is Nothing Then
                    MsgBox "IA_Table not found in 'IA_Level' of the previous file." & _
                           vbCrLf & "Manual columns will remain blank.", vbCritical
                    wbPrevious.Close SaveChanges:=False
                    wantManualData = False
                Else
                    If Not iaTablePrev.DataBodyRange Is Nothing Then
                        numRowsIAPrev = iaTablePrev.DataBodyRange.Rows.Count
                        If numRowsIAPrev > 0 Then
                            manualData = iaTablePrev.DataBodyRange.Value
                            
                            ' Check columns in the old IA_Table
                            Call CheckColumnExists(iaTablePrev, "Fund Manager GCI")
                            Call CheckColumnExists(iaTablePrev, "Days to Report")
                            Call CheckColumnExists(iaTablePrev, "1st Client Outreach Date")
                            Call CheckColumnExists(iaTablePrev, "2nd Client Outreach Date")
                            Call CheckColumnExists(iaTablePrev, "OA Escalation Date")
                            Call CheckColumnExists(iaTablePrev, "NOA Escalation Date")
                            Call CheckColumnExists(iaTablePrev, "Escalation Name")
                            Call CheckColumnExists(iaTablePrev, "Final Status")
                            Call CheckColumnExists(iaTablePrev, "Comments")
                            ' If no message popped up, columns exist
                            
                            ' Loop old IA_Table rows
                            For i = 1 To numRowsIAPrev
                                Dim prevGCI As String
                                prevGCI = manualData(i, iaTablePrev.ListColumns("Fund Manager GCI").Index)
                                
                                If prevGCI <> "" Then
                                    manualColumns.Add prevGCI, Array( _
                                        manualData(i, iaTablePrev.ListColumns("Days to Report").Index), _
                                        manualData(i, iaTablePrev.ListColumns("1st Client Outreach Date").Index), _
                                        manualData(i, iaTablePrev.ListColumns("2nd Client Outreach Date").Index), _
                                        manualData(i, iaTablePrev.ListColumns("OA Escalation Date").Index), _
                                        manualData(i, iaTablePrev.ListColumns("NOA Escalation Date").Index), _
                                        manualData(i, iaTablePrev.ListColumns("Escalation Name").Index), _
                                        manualData(i, iaTablePrev.ListColumns("Final Status").Index), _
                                        manualData(i, iaTablePrev.ListColumns("Comments").Index))
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    '---------------------------------------------------------------------------------
    ' 9. READ DATA FROM PORTFOLIOTABLE
    '---------------------------------------------------------------------------------
    portfolioData = portfolioTable.DataBodyRange.Value
    numRowsPortfolio = UBound(portfolioData, 1)
    
    Dim colGCI As Long, colRegion As Long, colManager As Long
    Dim colTriggerNonTrigger As Long, colNavSource As Long
    Dim colPriContact As Long, colSecContact As Long, colWksMissing As Long
    
    ' Retrieve column indexes once for efficiency
    colGCI = portfolioTable.ListColumns("Fund Manager GCI").Index
    colRegion = portfolioTable.ListColumns("Region").Index
    colManager = portfolioTable.ListColumns("Fund Manager").Index
    colTriggerNonTrigger = portfolioTable.ListColumns("Trigger/Non-Trigger").Index
    colNavSource = portfolioTable.ListColumns("NAV Source").Index
    colPriContact = portfolioTable.ListColumns("Primary Client Contact").Index
    colSecContact = portfolioTable.ListColumns("Secondary Client Contact").Index
    colWksMissing = portfolioTable.ListColumns("Wks Missing").Index
    
    '---------------------------------------------------------------------------------
    '10. LOOP THROUGH PORTFOLIOTABLE ROWS & BUILD DICTIONARIES
    '---------------------------------------------------------------------------------
    For i = 1 To numRowsPortfolio
        
        gci = SafeStringValue(portfolioData(i, colGCI))
        
        Dim region As String
        Dim manager As String
        Dim triggerStatus As String
        Dim navSource As String
        Dim primaryContact As String
        Dim secondaryContact As String
        Dim wksMissing As String
        
        region = SafeStringValue(portfolioData(i, colRegion))
        manager = SafeStringValue(portfolioData(i, colManager))
        triggerStatus = SafeStringValue(portfolioData(i, colTriggerNonTrigger))
        navSource = SafeStringValue(portfolioData(i, colNavSource))
        primaryContact = SafeStringValue(portfolioData(i, colPriContact))
        secondaryContact = SafeStringValue(portfolioData(i, colSecContact))
        wksMissing = SafeStringValue(portfolioData(i, colWksMissing))
        
        ' Add to dictionary if not already present
        If Not uniqueGCI.Exists(gci) And gci <> "" Then
            uniqueGCI.Add gci, True
            
            regionData.Add gci, region
            managerData.Add gci, manager
            triggerData.Add gci, triggerStatus
            navSourceData.Add gci, navSource
            clientContactData.Add gci, _
                primaryContact & IIf(primaryContact <> "" And secondaryContact <> "", ";", "") & secondaryContact
            
            triggerCountData.Add gci, 0
            nonTriggerCountData.Add gci, 0
            missingTriggerData.Add gci, 0
            missingNonTriggerData.Add gci, 0
        End If
        
        ' Tally counts only if gci is not blank
        If gci <> "" Then
            If triggerStatus = "Trigger" Then
                triggerCountData(gci) = triggerCountData(gci) + 1
                If wksMissing <> "" Then
                    missingTriggerData(gci) = missingTriggerData(gci) + 1
                End If
            ElseIf triggerStatus = "Non-Trigger" Then
                nonTriggerCountData(gci) = nonTriggerCountData(gci) + 1
                If wksMissing <> "" Then
                    missingNonTriggerData(gci) = missingNonTriggerData(gci) + 1
                End If
            End If
        End If
    Next i
    
    '---------------------------------------------------------------------------------
    '11. BUILD THE FINAL ARRAY FOR IA_TABLE (20 COLUMNS)
    '---------------------------------------------------------------------------------
    numUniqueGCI = uniqueGCI.Count
    If numUniqueGCI = 0 Then
        MsgBox "No unique GCIs found in the PortfolioTable. Nothing to populate.", vbExclamation
        GoTo CleanUp
    End If
    
    ReDim iaData(1 To numUniqueGCI, 1 To 20)
    
    j = 1
    For Each gci In uniqueGCI.Keys
        
        ' 1  - Fund Manager GCI
        ' 2  - Region
        ' 3  - Fund Manager
        ' 4  - Trigger/Non-Trigger
        ' 5  - NAV Source
        ' 6  - Client Contact(s)
        ' 7  - Trigger
        ' 8  - Non-Trigger
        ' 9  - Total Funds
        '10 - Missing Trigger
        '11 - Missing Non-Trigger
        '12 - Total Missing
        '13 - Days to Report (manual)
        '14 - 1st Client Outreach Date (manual)
        '15 - 2nd Client Outreach Date (manual)
        '16 - OA Escalation Date (manual)
        '17 - NOA Escalation Date (manual)
        '18 - Escalation Name (manual)
        '19 - Final Status (manual)
        '20 - Comments (manual)
        
        iaData(j, 1) = gci
        iaData(j, 2) = regionData(gci)
        iaData(j, 3) = managerData(gci)
        iaData(j, 4) = triggerData(gci)
        iaData(j, 5) = navSourceData(gci)
        iaData(j, 6) = clientContactData(gci)
        
        iaData(j, 7) = triggerCountData(gci)         ' "Trigger"
        iaData(j, 8) = nonTriggerCountData(gci)      ' "Non-Trigger"
        iaData(j, 9) = iaData(j, 7) + iaData(j, 8)   ' "Total Funds"
        
        iaData(j, 10) = missingTriggerData(gci)      ' "Missing Trigger"
        iaData(j, 11) = missingNonTriggerData(gci)   ' "Missing Non-Trigger"
        iaData(j, 12) = iaData(j, 10) + iaData(j, 11)' "Total Missing"
        
        If wantManualData And manualColumns.Exists(gci) Then
            Dim manualValues As Variant
            manualValues = manualColumns(gci)
            
            iaData(j, 13) = manualValues(0)   ' Days to Report
            iaData(j, 14) = manualValues(1)   ' 1st Client Outreach Date
            iaData(j, 15) = manualValues(2)   ' 2nd Client Outreach Date
            iaData(j, 16) = manualValues(3)   ' OA Escalation Date
            iaData(j, 17) = manualValues(4)   ' NOA Escalation Date
            iaData(j, 18) = manualValues(5)   ' Escalation Name
            iaData(j, 19) = manualValues(6)   ' Final Status
            iaData(j, 20) = manualValues(7)   ' Comments
        Else
            ' Leave manual columns blank if we are NOT loading from previous workbook
            iaData(j, 13) = ""
            iaData(j, 14) = ""
            iaData(j, 15) = ""
            iaData(j, 16) = ""
            iaData(j, 17) = ""
            iaData(j, 18) = ""
            iaData(j, 19) = ""
            iaData(j, 20) = ""
        End If
        
        j = j + 1
    Next gci
    
    '---------------------------------------------------------------------------------
    '12. WRITE FINAL ARRAY INTO IA_TABLE
    '---------------------------------------------------------------------------------
    iaTable.DataBodyRange.Resize(numUniqueGCI, 20).Value = iaData
    
    '---------------------------------------------------------------------------------
    '13. COMPLETION MESSAGE
    '---------------------------------------------------------------------------------
    MsgBox "IA_Table has been created/populated successfully.", vbInformation

CleanUp:

    '---------------------------------------------------------------------------------
    '14. RESTORE APPLICATION SETTINGS & CLOSE PREVIOUS WORKBOOK IF OPEN
    '---------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If Not wbPrevious Is Nothing Then
        wbPrevious.Close SaveChanges:=False
    End If
    
End Sub

'---------------------------------------------------------------------------------------------
' HELPER PROCEDURE: Checks if a given column name exists in a ListObject (table).
' If not, displays a message and stops the code so you can see the issue.
'---------------------------------------------------------------------------------------------
Private Sub CheckColumnExists(ByVal lo As ListObject, ByVal colName As String)
    Dim foundCol As ListColumn
    On Error Resume Next
    Set foundCol = lo.ListColumns(colName)
    On Error GoTo 0
    
    If foundCol Is Nothing Then
        MsgBox "Column '" & colName & "' not found in table '" & lo.Name & "'" & vbCrLf & _
               "Sheet: " & lo.Parent.Name, vbCritical
        Stop    ' Forces code to halt here in debug mode
    End If
End Sub

'---------------------------------------------------------------------------------------------
' HELPER FUNCTION: Safely converts a cell value to string, replacing #N/A (or any error) with ""
'---------------------------------------------------------------------------------------------
Private Function SafeStringValue(cellValue As Variant) As String
    If IsError(cellValue) Then
        SafeStringValue = ""  ' Replace any Excel error (#N/A, #VALUE!, etc.) with blank
    Else
        SafeStringValue = CStr(cellValue)
    End If
End Function
