Sub MacroEpsilon()
    Dim wsPortfolio As Worksheet, wsIA As Worksheet, wsDataset As Worksheet
    Dim portfolioTable As ListObject, iaTable As ListObject, datasetTable As ListObject
    Dim previousFile As String
    Dim wbPrevious As Workbook, wsIAPrev As Worksheet, iaTablePrev As ListObject
    Dim numRowsPortfolio As Long, numRowsIAPrev As Long
    Dim uniqueGCI As Object
    Dim regionData As Object, managerData As Object, triggerData As Object
    Dim navSourceData As Object, clientContactData As Object, ecaAnalystData As Object
    Dim triggerCountData As Object, nonTriggerCountData As Object
    Dim missingTriggerData As Object, missingNonTriggerData As Object
    Dim manualColumns As Object
    Dim i As Long, j As Long
    Dim gci As Variant

    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Initialize worksheets and tables
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set wsIA = ThisWorkbook.Sheets("IA_Level")
    Set wsDataset = ThisWorkbook.Sheets("Dataset")
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    Set iaTable = wsIA.ListObjects("IA_Table")
    Set datasetTable = wsDataset.ListObjects("DatasetTable")

    ' Validate tables
    If portfolioTable Is Nothing Or iaTable Is Nothing Or datasetTable Is Nothing Then
        MsgBox "One or more required tables (PortfolioTable, IA_Table, DatasetTable) are missing.", vbCritical
        GoTo CleanUp
    End If

    ' Clear IA_Table except headers
    If Not iaTable.DataBodyRange Is Nothing Then iaTable.DataBodyRange.Delete

    ' Prompt user to select the previous version of IA_Table
    previousFile = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select the previous version of IA_Table")
    If previousFile = "False" Then
        MsgBox "No file selected.", vbExclamation
        GoTo CleanUp
    End If
    Set wbPrevious = Workbooks.Open(previousFile)
    Set wsIAPrev = wbPrevious.Sheets("IA_Level")
    Set iaTablePrev = wsIAPrev.ListObjects("IA_Table")

    If iaTablePrev Is Nothing Then
        MsgBox "IA_Table does not exist in the previous version.", vbCritical
        wbPrevious.Close SaveChanges:=False
        GoTo CleanUp
    End If

    ' Initialize dictionaries
    Set uniqueGCI = CreateObject("Scripting.Dictionary")
    Set regionData = CreateObject("Scripting.Dictionary")
    Set managerData = CreateObject("Scripting.Dictionary")
    Set triggerData = CreateObject("Scripting.Dictionary")
    Set navSourceData = CreateObject("Scripting.Dictionary")
    Set clientContactData = CreateObject("Scripting.Dictionary")
    Set ecaAnalystData = CreateObject("Scripting.Dictionary")
    Set triggerCountData = CreateObject("Scripting.Dictionary")
    Set nonTriggerCountData = CreateObject("Scripting.Dictionary")
    Set missingTriggerData = CreateObject("Scripting.Dictionary")
    Set missingNonTriggerData = CreateObject("Scripting.Dictionary")
    Set manualColumns = CreateObject("Scripting.Dictionary")

    ' Read manual data from previous IA_Table
    numRowsIAPrev = iaTablePrev.DataBodyRange.Rows.Count
    If numRowsIAPrev > 0 Then
        Dim manualData As Variant
        manualData = iaTablePrev.DataBodyRange.Value
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

    ' Read PortfolioTable data into dictionaries
    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    Dim portfolioData As Variant
    portfolioData = portfolioTable.DataBodyRange.Value
    For i = 1 To numRowsPortfolio
        Dim region As String, manager As String
        Dim triggerStatus As String, navSource As String
        Dim primaryContact As String, secondaryContact As String, wksMissing As String
        gci = portfolioData(i, portfolioTable.ListColumns("Fund Manager GCI").Index)
        region = portfolioData(i, portfolioTable.ListColumns("Region").Index)
        manager = portfolioData(i, portfolioTable.ListColumns("Fund Manager").Index)
        triggerStatus = portfolioData(i, portfolioTable.ListColumns("Trigger/Non-Trigger").Index)
        navSource = portfolioData(i, portfolioTable.ListColumns("NAV Source").Index)
        primaryContact = portfolioData(i, portfolioTable.ListColumns("Primary Client Contact").Index)
        secondaryContact = portfolioData(i, portfolioTable.ListColumns("Secondary Client Contact").Index)
        wksMissing = portfolioData(i, portfolioTable.ListColumns("Wks Missing").Index)

        If Not uniqueGCI.Exists(gci) Then
            uniqueGCI.Add gci, Nothing
            regionData.Add gci, region
            managerData.Add gci, manager
            triggerData.Add gci, triggerStatus
            navSourceData.Add gci, navSource
            clientContactData.Add gci, primaryContact & IIf(primaryContact <> "" And secondaryContact <> "", ";", "") & secondaryContact
            triggerCountData.Add gci, 0
            nonTriggerCountData.Add gci, 0
            missingTriggerData.Add gci, 0
            missingNonTriggerData.Add gci, 0
        End If

        If triggerStatus = "Trigger" Then
            triggerCountData(gci) = triggerCountData(gci) + 1
            If wksMissing <> "" Then missingTriggerData(gci) = missingTriggerData(gci) + 1
        ElseIf triggerStatus = "Non-Trigger" Then
            nonTriggerCountData(gci) = nonTriggerCountData(gci) + 1
            If wksMissing <> "" Then missingNonTriggerData(gci) = missingNonTriggerData(gci) + 1
        End If
    Next i

    ' Prepare data for IA_Table
    Dim iaData() As Variant
    Dim numUniqueGCI As Long
    numUniqueGCI = uniqueGCI.Count
    ReDim iaData(1 To numUniqueGCI, 1 To 20)

    j = 1
    For Each gci In uniqueGCI.Keys
        iaData(j, 1) = gci
        iaData(j, 2) = regionData(gci)
        iaData(j, 3) = managerData(gci)
        iaData(j, 4) = triggerData(gci)
        iaData(j, 5) = navSourceData(gci)
        iaData(j, 6) = clientContactData(gci)
        iaData(j, 7) = triggerCountData(gci)
        iaData(j, 8) = nonTriggerCountData(gci)
        iaData(j, 9) = iaData(j, 7) + iaData(j, 8) ' Total Funds
        iaData(j, 10) = missingTriggerData(gci)
        iaData(j, 11) = missingNonTriggerData(gci)
        iaData(j, 12) = iaData(j, 10) + iaData(j, 11) ' Total Missing

        ' Populate manual columns if available
        If manualColumns.exists(gci) Then
            Dim manualValues As Variant
            manualValues = manualColumns(gci)
            iaData(j, 13) = manualValues(0) ' Days to Report
            iaData(j, 14) = manualValues(1) ' 1st Client Outreach Date
            iaData(j, 15) = manualValues(2) ' 2nd Client Outreach Date
            iaData(j, 16) = manualValues(3) ' OA Escalation Date
            iaData(j, 17) = manualValues(4) ' NOA Escalation Date
            iaData(j, 18) = manualValues(5) ' Escalation Name
            iaData(j, 19) = manualValues(6) ' Final Status
            iaData(j, 20) = manualValues(7) ' Comments
        End If
        j = j + 1
    Next gci

    ' Write data to IA_Table
    iaTable.DataBodyRange.Resize(numUniqueGCI, 20).Value = iaData

    ' Completion message
    MsgBox "IA_Table has been populated successfully.", vbInformation

CleanUp:
    ' Reset application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Not wbPrevious Is Nothing Then wbPrevious.Close SaveChanges:=False
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
