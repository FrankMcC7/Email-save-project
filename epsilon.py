Option Explicit

Sub MacroEpsilon()

    '---------------------------------------------------------------------------------
    ' 1. DECLARE & INITIALIZE
    '---------------------------------------------------------------------------------
    Dim wsPortfolio As Worksheet, wsIA As Worksheet
    Dim portfolioTable As ListObject, iaTable As ListObject
    
    ' Only one declaration of wbPrevious here:
    Dim wbPrevious As Workbook, wsIAPrev As Worksheet
    Dim iaTablePrev As ListObject
    
    ' Dictionaries for storing data by GCI
    Dim uniqueGCI As Object
    Dim regionData As Object, managerData As Object
    Dim triggerCountData As Object, nonTriggerCountData As Object
    Dim missingTriggerData As Object, missingNonTriggerData As Object
    
    ' We use a dictionary of dictionaries for NAV Source (to store multiple NAVs per GCI)
    Dim navSourceData As Object
    
    Dim clientContactData As Object, manualColumns As Object
    
    Dim portfolioData As Variant, manualData As Variant, iaData() As Variant
    Dim userChoice As VbMsgBoxResult
    Dim wantManualData As Boolean
    
    Dim previousFile As Variant
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
    
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    
    If portfolioTable Is Nothing Then
        MsgBox "Table 'PortfolioTable' not found on the 'Portfolio' sheet." & vbCrLf & _
               "Please verify the table name/spelling.", vbCritical
        GoTo CleanUp
    End If
    
    If portfolioTable.DataBodyRange Is Nothing Then
        MsgBox "No data found in 'PortfolioTable'. Please ensure it has at least one row.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Confirm required columns in PortfolioTable
    CheckColumnExists portfolioTable, "Fund Manager GCI"
    CheckColumnExists portfolioTable, "Region"
    CheckColumnExists portfolioTable, "Fund Manager"
    CheckColumnExists portfolioTable, "Trigger/Non-Trigger"
    CheckColumnExists portfolioTable, "NAV Source"
    CheckColumnExists portfolioTable, "Primary Client Contact"
    CheckColumnExists portfolioTable, "Secondary Client Contact"
    CheckColumnExists portfolioTable, "Wks Missing"
    
    '---------------------------------------------------------------------------------
    ' 5. CHECK IF IA_TABLE EXISTS; IF NOT, CREATE IT (20 columns)
    '---------------------------------------------------------------------------------
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
    Set triggerCountData = CreateObject("Scripting.Dictionary")
    Set nonTriggerCountData = CreateObject("Scripting.Dictionary")
    Set missingTriggerData = CreateObject("Scripting.Dictionary")
    Set missingNonTriggerData = CreateObject("Scripting.Dictionary")
    
    ' navSourceData is a Dictionary of mini-dictionaries
    Set navSourceData = CreateObject("Scripting.Dictionary")
    
    Set clientContactData = CreateObject("Scripting.Dictionary")
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
            Set wbPrevious = Workbooks.Open(CStr(previousFile))
            
            On Error Resume Next
            Set wsIAPrev = wbPrevious.Sheets("IA_Level")
            On Error GoTo 0
            
            If wsIAPrev Is Nothing Then
                MsgBox "Sheet 'IA_Level' not found in the previous file. Manual columns will remain blank.", vbCritical
                wbPrevious.Close SaveChanges:=False
                wantManualData = False
            Else
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
                            
                            ' Check columns in old IA_Table
                            CheckColumnExists iaTablePrev, "Fund Manager GCI"
                            CheckColumnExists iaTablePrev, "Days to Report"
                            CheckColumnExists iaTablePrev, "1st Client Outreach Date"
                            CheckColumnExists iaTablePrev, "2nd Client Outreach Date"
                            CheckColumnExists iaTablePrev, "OA Escalation Date"
                            CheckColumnExists iaTablePrev, "NOA Escalation Date"
                            CheckColumnExists iaTablePrev, "Escalation Name"
                            CheckColumnExists iaTablePrev, "Final Status"
                            CheckColumnExists iaTablePrev, "Comments"
                            
                            Dim prevGCI As String
                            Dim r As Long
                            For r = 1 To numRowsIAPrev
                                prevGCI = manualData(r, iaTablePrev.ListColumns("Fund Manager GCI").Index)
                                
                                If prevGCI <> "" Then
                                    manualColumns.Add prevGCI, Array( _
                                        manualData(r, iaTablePrev.ListColumns("Days to Report").Index), _
                                        manualData(r, iaTablePrev.ListColumns("1st Client Outreach Date").Index), _
                                        manualData(r, iaTablePrev.ListColumns("2nd Client Outreach Date").Index), _
                                        manualData(r, iaTablePrev.ListColumns("OA Escalation Date").Index), _
                                        manualData(r, iaTablePrev.ListColumns("NOA Escalation Date").Index), _
                                        manualData(r, iaTablePrev.ListColumns("Escalation Name").Index), _
                                        manualData(r, iaTablePrev.ListColumns("Final Status").Index), _
                                        manualData(r, iaTablePrev.ListColumns("Comments").Index))
                                End If
                            Next r
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
    Dim colPriContact As Long, colSecContact As Long
    Dim colWksMissing As Long
    
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
    Dim fundGCI As String, navSrc As String
    Dim region As String, manager As String
    Dim triggerStatus As String
    Dim primaryContact As String, secondaryContact As String
    Dim wksMissing As String
    
    For i = 1 To numRowsPortfolio
        
        fundGCI = SafeStringValue(portfolioData(i, colGCI))
        region = SafeStringValue(portfolioData(i, colRegion))
        manager = SafeStringValue(portfolioData(i, colManager))
        triggerStatus = SafeStringValue(portfolioData(i, colTriggerNonTrigger))
        navSrc = SafeStringValue(portfolioData(i, colNavSource))
        
        primaryContact = SafeStringValue(portfolioData(i, colPriContact))
        secondaryContact = SafeStringValue(portfolioData(i, colSecContact))
        wksMissing = SafeStringValue(portfolioData(i, colWksMissing))
        
        If fundGCI <> "" Then
            
            ' If first time seeing this GCI, set up data placeholders
            If Not uniqueGCI.Exists(fundGCI) Then
                uniqueGCI.Add fundGCI, True
                regionData.Add fundGCI, region
                managerData.Add fundGCI, manager
                
                triggerCountData.Add fundGCI, 0
                nonTriggerCountData.Add fundGCI, 0
                missingTriggerData.Add fundGCI, 0
                missingNonTriggerData.Add fundGCI, 0
                
                ' For storing multiple NAV sources:
                Dim navDict As Object
                Set navDict = CreateObject("Scripting.Dictionary")
                navSourceData.Add fundGCI, navDict
                
                clientContactData.Add fundGCI, _
                    primaryContact & IIf(primaryContact <> "" And secondaryContact <> "", ";", "") & secondaryContact
            End If
            
            ' Record this NAV Source if not blank
            If navSrc <> "" Then
                If Not navSourceData(fundGCI).Exists(navSrc) Then
                    navSourceData(fundGCI).Add navSrc, True
                End If
            End If
            
            ' Tally Trigger / Non-Trigger counts
            If triggerStatus = "Trigger" Then
                triggerCountData(fundGCI) = triggerCountData(fundGCI) + 1
                If wksMissing <> "" Then
                    missingTriggerData(fundGCI) = missingTriggerData(fundGCI) + 1
                End If
            ElseIf triggerStatus = "Non-Trigger" Then
                nonTriggerCountData(fundGCI) = nonTriggerCountData(fundGCI) + 1
                If wksMissing <> "" Then
                    missingNonTriggerData(fundGCI) = missingNonTriggerData(fundGCI) + 1
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
    Dim manualValues As Variant
    
    For Each gci In uniqueGCI.Keys
        
        ' (1) Fund Manager GCI
        iaData(j, 1) = gci
        
        ' (2) Region
        iaData(j, 2) = regionData(gci)
        
        ' (3) Fund Manager
        iaData(j, 3) = managerData(gci)
        
        ' (4) Trigger/Non-Trigger:
        If triggerCountData(gci) > 0 And nonTriggerCountData(gci) > 0 Then
            iaData(j, 4) = "Both"
        ElseIf triggerCountData(gci) > 0 Then
            iaData(j, 4) = "Trigger"
        ElseIf nonTriggerCountData(gci) > 0 Then
            iaData(j, 4) = "Non-Trigger"
        Else
            iaData(j, 4) = ""
        End If
        
        ' (5) NAV Source: if no NAV sources, use "[No NAV Source]" else join with ";"
        Dim arrNav() As String
        arrNav = navSourceData(gci).Keys  ' All distinct NAV Source values
        If navSourceData(gci).Count = 0 Then
            ' No NAV Source found for this GCI
            iaData(j, 5) = "[No NAV Source]"
        Else
            iaData(j, 5) = Join(arrNav, ";")
        End If
        
        ' (6) Client Contact(s)
        iaData(j, 6) = clientContactData(gci)
        
        ' (7) # of Trigger
        iaData(j, 7) = triggerCountData(gci)
        
        ' (8) # of Non-Trigger
        iaData(j, 8) = nonTriggerCountData(gci)
        
        ' (9) Total Funds
        iaData(j, 9) = iaData(j, 7) + iaData(j, 8)
        
        ' (10) Missing Trigger
        iaData(j, 10) = missingTriggerData(gci)
        
        ' (11) Missing Non-Trigger
        iaData(j, 11) = missingNonTriggerData(gci)
        
        ' (12) Total Missing
        iaData(j, 12) = iaData(j, 10) + iaData(j, 11)
        
        ' (13–20) Manual columns if user chose to load them
        If wantManualData And manualColumns.Exists(gci) Then
            manualValues = manualColumns(gci)
            iaData(j, 13) = manualValues(0)  ' Days to Report
            iaData(j, 14) = manualValues(1)  ' 1st Client Outreach Date
            iaData(j, 15) = manualValues(2)  ' 2nd Client Outreach Date
            iaData(j, 16) = manualValues(3)  ' OA Escalation Date
            iaData(j, 17) = manualValues(4)  ' NOA Escalation Date
            iaData(j, 18) = manualValues(5)  ' Escalation Name
            iaData(j, 19) = manualValues(6)  ' Final Status
            iaData(j, 20) = manualValues(7)  ' Comments
        Else
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
    '12. WRITE FINAL ARRAY INTO IA_TABLE (RESIZE PROPERLY)
    '---------------------------------------------------------------------------------
    Dim targetRange As Range
    
    If iaTable.DataBodyRange Is Nothing Then
        ' The table has only headers and no data rows:
        
        Dim startCell As Range
        Set startCell = iaTable.HeaderRowRange.Offset(1, 0).Cells(1, 1)
        
        Set targetRange = startCell.Resize(numUniqueGCI, 20)
        targetRange.Value = iaData
        
        iaTable.Resize iaTable.Range.Resize(numUniqueGCI + 1, 20)
    Else
        ' Table has some rows (or zero if just cleared)
        
        Dim existingBody As Range
        Set existingBody = iaTable.DataBodyRange
        
        Dim firstDataCell As Range
        Set firstDataCell = existingBody.Cells(1, 1)
        
        Set targetRange = firstDataCell.Resize(numUniqueGCI, 20)
        targetRange.Value = iaData
        
        iaTable.Resize iaTable.HeaderRowRange.Resize(numUniqueGCI + 1, 20)
    End If
    
    MsgBox "IA_Table has been created/populated successfully.", vbInformation

CleanUp:
    
    '---------------------------------------------------------------------------------
    '13. RESTORE APPLICATION SETTINGS & CLOSE PREVIOUS WORKBOOK IF OPEN
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
' If not, displays a message and stops the code so you can see the issue in debug.
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
' HELPER FUNCTION: Converts a cell value to string, replacing #N/A (or any error) with ""
'---------------------------------------------------------------------------------------------
Private Function SafeStringValue(cellValue As Variant) As String
    If IsError(cellValue) Then
        SafeStringValue = ""  ' Replace any Excel error (#N/A, #VALUE!, etc.) with blank
    Else
        SafeStringValue = CStr(cellValue)
    End If
End Function
