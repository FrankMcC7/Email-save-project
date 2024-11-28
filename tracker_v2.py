Sub ProcessAllData()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, allFundsTable As ListObject, nonTriggerTable As ListObject
    Dim dictFundGCI As Object, dictDataset As Object
    Dim fundGCIArray As Variant, resultArray As Variant
    Dim i As Long, lastRowPortfolio As Long
    Dim triggerColIndex As Long ' Column index for "Trigger/Non-Trigger"
    Dim startTime As Double
    Dim wsDataset As Worksheet
    Dim datasetTable As ListObject
    Dim portfolioFundManagerGCI As Variant
    Dim portfolioFamily As Variant
    Dim portfolioECAAnalyst As Variant
    Dim datasetFundManagerGCI As Variant
    Dim datasetFamily As Variant
    Dim datasetECAAnalyst As Variant
    Dim numRowsPortfolio As Long
    Dim numRowsDataset As Long
    Dim colNames As Object
    Dim headers As Variant
    Dim sourceHeaders As Variant, destHeaders As Variant
    Dim reviewStatusCol As Long, fundGCIColIndex As Long, iaGCIColIndex As Long
    Dim totalRows As Long
    Dim allFundsData As Variant
    Dim businessUnitColIndex As Long
    Dim fundGCIKey As String
    Dim key As String

    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Record the start time
    startTime = Timer

    ' === Set up the Portfolio sheet ===
    Set wbMaster = ThisWorkbook
    Set wsPortfolio = wbMaster.Sheets("Portfolio")

    ' Ensure the Portfolio sheet is converted to a table
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    If portfolioTable Is Nothing Then
        ' Convert the range to a table if not already one
        If wsPortfolio.UsedRange.Rows.Count > 1 Then
            Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, wsPortfolio.UsedRange, , xlYes)
        Else
            MsgBox "The Portfolio sheet does not contain headers or data.", vbCritical
            GoTo CleanUp
        End If
        portfolioTable.Name = "PortfolioTable"
    End If

    ' Clear existing data in the Portfolio table except headers
    If Not portfolioTable.DataBodyRange Is Nothing Then
        portfolioTable.DataBodyRange.Delete
    End If

    ' === Step 1: Process Trigger.csv ===
    triggerFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select Trigger.csv")
    If triggerFile = "False" Then GoTo CleanUp ' User canceled

    Set wbTrigger = Workbooks.Open(triggerFile)
    Set triggerSheet = wbTrigger.Sheets(1)

    ' Convert Trigger data to a table if it isn't already one
    On Error Resume Next
    Set triggerTable = triggerSheet.ListObjects(1)
    On Error GoTo 0
    If triggerTable Is Nothing Then
        Set triggerTable = triggerSheet.ListObjects.Add(xlSrcRange, triggerSheet.UsedRange, , xlYes)
        triggerTable.Name = "TriggerTable"
    End If

    ' Copy specific columns from Trigger.csv to Portfolio
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer")

    ' Calculate the row to start pasting data
    lastRowPortfolio = portfolioTable.HeaderRowRange.Row + 1

    For i = LBound(headers) To UBound(headers)
        With triggerTable.ListColumns(headers(i)).DataBodyRange
            wsPortfolio.Cells(lastRowPortfolio, portfolioTable.ListColumns(headers(i)).Index).Resize(.Rows.Count, 1).Value = .Value
        End With
    Next i

    ' Replace Region values and populate Trigger/Non-Trigger column
    triggerColIndex = portfolioTable.ListColumns("Trigger/Non-Trigger").Index
    With portfolioTable.DataBodyRange
        .Columns(portfolioTable.ListColumns("Region").Index).Replace What:="US", Replacement:="AMRS", LookAt:=xlWhole
        .Columns(portfolioTable.ListColumns("Region").Index).Replace What:="ASIA", Replacement:="APAC", LookAt:=xlWhole
        .Columns(triggerColIndex).Value = "Trigger"
    End With

    wbTrigger.Close SaveChanges:=False

    ' === Step 2: Process All Funds.csv ===
    allFundsFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select all fund.csv")
    If allFundsFile = "False" Then GoTo CleanUp ' User canceled

    Set wbAllFunds = Workbooks.Open(allFundsFile)
    Set allFundsSheet = wbAllFunds.Sheets(1)

    ' Delete the first row (assuming it's a header row)
    allFundsSheet.Rows(1).Delete

    ' Read All Funds data into arrays to improve performance
    totalRows = allFundsSheet.Cells(allFundsSheet.Rows.Count, "A").End(xlUp).Row

    ' Read the data into an array
    allFundsData = allFundsSheet.Range("A1").Resize(totalRows, allFundsSheet.UsedRange.Columns.Count).Value

    ' Get column indices
    Set colNames = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(allFundsData, 2)
        colNames(allFundsData(1, i)) = i
    Next i

    reviewStatusCol = colNames("Review Status")
    fundGCIColIndex = colNames("Fund GCI")
    iaGCIColIndex = colNames("IA GCI")

    ' Initialize dictionary
    Set dictFundGCI = CreateObject("Scripting.Dictionary")

    ' Loop through the data and add to dictionary where Review Status is "Approved"
    For i = 2 To UBound(allFundsData, 1)
        If allFundsData(i, reviewStatusCol) = "Approved" Then
            If Not IsEmpty(allFundsData(i, fundGCIColIndex)) And Not dictFundGCI.Exists(CStr(allFundsData(i, fundGCIColIndex))) Then
                dictFundGCI.Add CStr(allFundsData(i, fundGCIColIndex)), allFundsData(i, iaGCIColIndex)
            End If
        End If
    Next i

    ' Match Fund GCI in Portfolio and write IA GCI to Fund Manager GCI
    numRowsPortfolio = portfolioTable.ListRows.Count
    If numRowsPortfolio > 0 Then
        fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
        ReDim portfolioFundManagerGCI(1 To numRowsPortfolio, 1 To 1)

        For i = 1 To numRowsPortfolio
            fundGCIKey = CStr(fundGCIArray(i, 1))
            If dictFundGCI.Exists(fundGCIKey) Then
                portfolioFundManagerGCI(i, 1) = dictFundGCI(fundGCIKey)
            Else
                portfolioFundManagerGCI(i, 1) = "No Match Found"
            End If
        Next i

        ' Write results back to Fund Manager GCI column
        portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = portfolioFundManagerGCI
    End If

    wbAllFunds.Close SaveChanges:=False

    ' === Step 3: Process Non-Trigger.csv ===
    nonTriggerFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select Non-Trigger.csv")
    If nonTriggerFile = "False" Then GoTo CleanUp ' User canceled

    Set wbNonTrigger = Workbooks.Open(nonTriggerFile)
    Set nonTriggerSheet = wbNonTrigger.Sheets(1)

    ' Convert Non-Trigger data to a table
    On Error Resume Next
    Set nonTriggerTable = nonTriggerSheet.ListObjects(1)
    On Error GoTo 0
    If nonTriggerTable Is Nothing Then
        Set nonTriggerTable = nonTriggerSheet.ListObjects.Add(xlSrcRange, nonTriggerSheet.UsedRange, , xlYes)
        nonTriggerTable.Name = "NonTriggerTable"
    End If

    ' Delete rows where 'Business Unit' = 'FI-ASIA'
    businessUnitColIndex = nonTriggerTable.ListColumns("Business Unit").Index

    ' Loop backwards through the table's ListRows
    For i = nonTriggerTable.ListRows.Count To 1 Step -1
        If Trim(nonTriggerTable.ListRows(i).Range.Cells(1, businessUnitColIndex).Value) = "FI-ASIA" Then
            nonTriggerTable.ListRows(i).Delete
        End If
    Next i

    ' Append Non-Trigger data to Portfolio
    sourceHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Weeks Missing")
    destHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Wks Missing")

    ' Find the last row with data in the Portfolio table
    If portfolioTable.ListRows.Count > 0 Then
        lastRowPortfolio = portfolioTable.Range.Row + portfolioTable.Range.Rows.Count - 1
    Else
        lastRowPortfolio = portfolioTable.HeaderRowRange.Row
    End If

    Dim numRowsToCopy As Long
    numRowsToCopy = nonTriggerTable.DataBodyRange.Rows.Count

    If numRowsToCopy > 0 Then
        For i = LBound(sourceHeaders) To UBound(sourceHeaders)
            With nonTriggerTable.ListColumns(sourceHeaders(i)).DataBodyRange
                wsPortfolio.Cells(lastRowPortfolio + 1, portfolioTable.ListColumns(destHeaders(i)).Index).Resize(.Rows.Count, 1).Value = .Value
            End With
        Next i

        ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
        With wsPortfolio
            .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), .Cells(lastRowPortfolio + numRowsToCopy, triggerColIndex)).Value = "Non-Trigger"
        End With
    Else
        MsgBox "No data to copy from Non-Trigger.csv after deleting 'FI-ASIA' rows.", vbInformation
    End If

    wbNonTrigger.Close SaveChanges:=False

    ' === Step 4: Populate 'Family' and 'ECA India Analyst' from 'Dataset' sheet ===

    ' Set up Dataset sheet
    Set wsDataset = wbMaster.Sheets("Dataset")

    ' Ensure the Dataset sheet is converted to a table
    On Error Resume Next
    Set datasetTable = wsDataset.ListObjects("DatasetTable")
    On Error GoTo 0
    If datasetTable Is Nothing Then
        If wsDataset.UsedRange.Rows.Count > 1 Then
            Set datasetTable = wsDataset.ListObjects.Add(xlSrcRange, wsDataset.UsedRange, , xlYes)
        Else
            MsgBox "The Dataset sheet does not contain headers or data.", vbCritical
            GoTo CleanUp
        End If
        datasetTable.Name = "DatasetTable"
    End If

    ' Read Dataset data into arrays
    numRowsDataset = datasetTable.DataBodyRange.Rows.Count
    datasetFundManagerGCI = datasetTable.ListColumns("Fund Manager GCI").DataBodyRange.Value
    datasetFamily = datasetTable.ListColumns("Family").DataBodyRange.Value
    datasetECAAnalyst = datasetTable.ListColumns("ECA India Analyst").DataBodyRange.Value

    ' Create a dictionary for Dataset using Fund Manager GCI as key
    Set dictDataset = CreateObject("Scripting.Dictionary")
    For i = 1 To numRowsDataset
        key = CStr(datasetFundManagerGCI(i, 1))
        If Not dictDataset.Exists(key) Then
            dictDataset.Add key, Array(datasetFamily(i, 1), datasetECAAnalyst(i, 1))
        End If
    Next i

    ' Read Portfolio data into arrays
    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    If numRowsPortfolio > 0 Then
        portfolioFundManagerGCI = portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value
        portfolioFamily = portfolioTable.ListColumns("Family").DataBodyRange.Value
        portfolioECAAnalyst = portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Value

        ' Ensure arrays are properly dimensioned
        If Not IsArray(portfolioFundManagerGCI) Then ReDim portfolioFundManagerGCI(1 To 1, 1 To 1)
        If Not IsArray(portfolioFamily) Then ReDim portfolioFamily(1 To 1, 1 To 1)
        If Not IsArray(portfolioECAAnalyst) Then ReDim portfolioECAAnalyst(1 To 1, 1 To 1)

        ' Populate 'Family' and 'ECA India Analyst' in PortfolioTable
        For i = 1 To numRowsPortfolio
            ' Check for empty or error values in Fund Manager GCI
            If Not IsError(portfolioFundManagerGCI(i, 1)) And Not IsEmpty(portfolioFundManagerGCI(i, 1)) Then
                key = CStr(portfolioFundManagerGCI(i, 1))

                ' Update 'Family' only if it is empty or contains "No Match Found"
                If IsError(portfolioFamily(i, 1)) Or IsEmpty(portfolioFamily(i, 1)) Or portfolioFamily(i, 1) = "" Or portfolioFamily(i, 1) = "No Match Found" Then
                    If dictDataset.Exists(key) Then
                        portfolioFamily(i, 1) = dictDataset(key)(0)
                    Else
                        portfolioFamily(i, 1) = "No Match Found"
                    End If
                End If

                ' Always update 'ECA India Analyst' from DatasetTable
                If dictDataset.Exists(key) Then
                    portfolioECAAnalyst(i, 1) = dictDataset(key)(1)
                Else
                    portfolioECAAnalyst(i, 1) = "No Match Found"
                End If
            Else
                ' Handle empty or error values
                portfolioFamily(i, 1) = "No Match Found"
                portfolioECAAnalyst(i, 1) = "No Match Found"
            End If
        Next i

        ' Write updated data back to PortfolioTable
        portfolioTable.ListColumns("Family").DataBodyRange.Value = portfolioFamily
        portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Value = portfolioECAAnalyst
    End If

    ' === Completion Message ===
    MsgBox "Data from all files has been processed successfully!" & vbCrLf & _
           "Time taken: " & Format(Timer - startTime, "0.00") & " seconds.", vbInformation

CleanUp:
    ' Reset Application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
