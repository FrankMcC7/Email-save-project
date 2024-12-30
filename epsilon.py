Option Explicit

Sub MacroEpsilon()

    On Error GoTo ErrorHandler
    
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
    
    Dim previousFile As String
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
    
    ' If "Dataset" is truly not used, you may remove the references for it.
    '
    ' Dim wsDataset As Worksheet, datasetTable As ListObject
    ' Set wsDataset = ThisWorkbook.Sheets("Dataset")
    ' Set datasetTable = wsDataset.ListObjects("DatasetTable")
    
    '---------------------------------------------------------------------------------
    ' 4. GET PORTFOLIOTABLE
    '---------------------------------------------------------------------------------
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo ErrorHandler
    
    If portfolioTable Is Nothing Then
        MsgBox "Table 'PortfolioTable' not found on 'Portfolio' sheet.", vbCritical
        GoTo CleanUp
    End If
    
    '---------------------------------------------------------------------------------
    ' 5. CHECK IF IA_TABLE EXISTS; IF NOT, CREATE IT
    '---------------------------------------------------------------------------------
    On Error Resume Next
    Set iaTable = wsIA.ListObjects("IA_Table")
    On Error GoTo ErrorHandler
    
    If iaTable Is Nothing Then
        ' Create the required header row in row 1 (columns A:T = 20 columns)
        Dim requiredCols As Variant
        requiredCols = Array( _
            "Fund Manager GCI", "Region", "Fund Manager", "Trigger/Non-Trigger", "NAV Source", _
            "Client Contact(s)", "Trigger", "Non-Trigger", "Total Funds", "Missing Trigger", _
            "Missing Non-Trigger", "Total Missing", "Days to Report", "1st Client Outreach Date", _
            "2nd Client Outreach Date", "OA Escalation Date", "NOA Escalation Date", _
            "Escalation Name", "Final Status", "Comments")
        
        ' Place headers in Row 1
        Dim col As Long
        For col = LBound(requiredCols) To UBound(requiredCols)
            wsIA.Cells(1, col + 1).Value = requiredCols(col)
        Next col
        
        ' Create a ListObject (table) from A1:T1 (1 row of headers, no data rows yet)
        Set iaTable = wsIA.ListObjects.Add(SourceType:=xlSrcRange, _
                                           Source:=wsIA.Range("A1:T1"), _
                                           XlListObjectHasHeaders:=xlYes)
        iaTable.Name = "IA_Table"
    End If
    
    '---------------------------------------------------------------------------------
    ' 6. CLEAR IA_TABLE EXCEPT HEADERS
    '---------------------------------------------------------------------------------
    If Not iaTable.DataBodyRange Is Nothing Then
        iaTable.DataBodyRange.Delete
    End If
    
    '---------------------------------------------------------------------------------
    ' 7. PROMPT USER FOR MANUAL DATA IMPORT
    '---------------------------------------------------------------------------------
    userChoice = MsgBox( _
        "Would you like to load manual data from a previous version?", _
        vbYesNo + vbQuestion, _
        "Load Manual Data?")
    wantManualData = (userChoice = vbYes)
    
    '---------------------------------------------------------------------------------
    ' 8. INITIALIZE DICTIONARIES
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
    ' 9. IF USER WANTS MANUAL DATA, OPEN PREVIOUS WORKBOOK
    '---------------------------------------------------------------------------------
    If wantManualData Then
        
        previousFile = Application.GetOpenFilename( _
            "Excel Files (*.xls*), *.xls*", , "Select the previous version of IA_Table")
        
        If previousFile = "False" Then
            MsgBox "No file selected. Manual columns will remain blank.", vbExclamation
            wantManualData = False
        Else
            ' Open the previous workbook
            Set wbPrevious = Workbooks.Open(previousFile)
            
            On Error Resume Next
            Set wsIAPrev = wbPrevious.Sheets("IA_Level")
            On Error GoTo ErrorHandler
            
            If wsIAPrev Is Nothing Then
                MsgBox "Sheet 'IA_Level' not found in the previous version. Manual columns will remain blank.", vbCritical
                wbPrevious.Close SaveChanges:=False
                wantManualData = False
            Else
                Set iaTablePrev = wsIAPrev.ListObjects("IA_Table")
                If iaTablePrev Is Nothing Then
                    MsgBox "IA_Table not found in the previous version. Manual columns will remain blank.", vbCritical
                    wbPrevious.Close SaveChanges:=False
                    wantManualData = False
                Else
                    If Not iaTablePrev.DataBodyRange Is Nothing Then
                        numRowsIAPrev = iaTablePrev.DataBodyRange.Rows.Count
                        If numRowsIAPrev > 0 Then
                            manualData = iaTablePrev.DataBodyRange.Value
                            
                            ' Loop through old IA_Table rows
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
    '10. READ CURRENT PORTFOLIOTABLE DATA
    '---------------------------------------------------------------------------------
    If portfolioTable.DataBodyRange Is Nothing Then
        MsgBox "No data found in PortfolioTable.", vbExclamation
        GoTo CleanUp
    End If
    
    portfolioData = portfolioTable.DataBodyRange.Value
    numRowsPortfolio = UBound(portfolioData, 1)
    
    For i = 1 To numRowsPortfolio
        
        gci = portfolioData(i, portfolioTable.ListColumns("Fund Manager GCI").Index)
        
        Dim region As String
        Dim manager As String
        Dim triggerStatus As String
        Dim navSource As String
        Dim primaryContact As String
        Dim secondaryContact As String
        Dim wksMissing As String
        
        region = portfolioData(i, portfolioTable.ListColumns("Region").Index)
        manager = portfolioData(i, portfolioTable.ListColumns("Fund Manager").Index)
        triggerStatus = portfolioData(i, portfolioTable.ListColumns("Trigger/Non-Trigger").Index)
        navSource = portfolioData(i, portfolioTable.ListColumns("NAV Source").Index)
        primaryContact = portfolioData(i, portfolioTable.ListColumns("Primary Client Contact").Index)
        secondaryContact = portfolioData(i, portfolioTable.ListColumns("Secondary Client Contact").Index)
        wksMissing = portfolioData(i, portfolioTable.ListColumns("Wks Missing").Index)
        
        ' Add to dictionary if not already present
        If Not uniqueGCI.Exists(gci) Then
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
        
        ' Tally counts
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
    Next i
    
    '---------------------------------------------------------------------------------
    '11. BUILD THE FINAL ARRAY FOR IA_TABLE (20 COLUMNS)
    '---------------------------------------------------------------------------------
    numUniqueGCI = uniqueGCI.Count
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
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
    
End Sub
