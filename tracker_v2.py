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
        Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, wsPortfolio.UsedRange, , xlYes)
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
    Dim headers As Variant
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
    Dim allFundsData As Variant
    Dim reviewStatusCol As Long, fundGCIColIndex As Long, iaGCIColIndex As Long
    Dim totalRows As Long

    totalRows = allFundsSheet.Cells(allFundsSheet.Rows.Count, "A").End(xlUp).Row

    ' Read the data into an array
    allFundsData = allFundsSheet.Range("A1").Resize(totalRows, allFundsSheet.UsedRange.Columns.Count).Value

    ' Get column indices
    Dim colNames As Object
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
            If Not IsEmpty(allFundsData(i, fundGCIColIndex)) And Not dictFundGCI.exists(allFundsData(i, fundGCIColIndex)) Then
                dictFundGCI.Add allFundsData(i, fundGCIColIndex), allFundsData(i, iaGCIColIndex)
            End If
        End If
    Next i

    ' Match Fund GCI in Portfolio and write IA GCI to Fund Manager GCI
    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim portfolioFundManagerGCI(1 To numRowsPortfolio, 1 To 1)

    For i = 1 To numRowsPortfolio
        If dictFundGCI.exists(fundGCIArray(i, 1)) Then
            portfolioFundManagerGCI(i, 1) = dictFundGCI(fundGCIArray(i, 1))
        Else
            portfolioFundManagerGCI(i, 1) = "No Match Found"
        End If
    Next i

    ' Write results back to Fund Manager GCI column
    portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = portfolioFundManagerGCI

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

    ' Filter out rows where FI-ASIA is present
    nonTriggerTable.Range.AutoFilter Field:=nonTriggerTable.ListColumns("Region").Index, Criteria1:="<>FI-ASIA"

    ' Append Non-Trigger data to Portfolio
    Dim sourceHeaders As Variant, destHeaders As Variant
    sourceHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Weeks Missing")
    destHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Wks Missing")

    ' Find the last row with data in the Portfolio table
    If portfolioTable.ListRows.Count > 0 Then
        lastRowPortfolio = portfolioTable.Range.Row + portfolioTable.Range.Rows.Count - 1
    Else
        lastRowPortfolio = portfolioTable.HeaderRowRange.Row
    End If

    Dim numRowsToCopy As Long
    numRowsToCopy = nonTriggerTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count

    For i = LBound(sourceHeaders) To UBound(sourceHeaders)
        With nonTriggerTable.ListColumns(sourceHeaders(i)).DataBodyRange.SpecialCells(xlCellTypeVisible)
            wsPortfolio.Cells(lastRowPortfolio + 1, portfolioTable.ListColumns(destHeaders(i)).Index).Resize(.Rows.Count, 1).Value = .Value
        End With
    Next i

    ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
    With wsPortfolio
        .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), .Cells(lastRowPortfolio + numRowsToCopy, triggerColIndex)).Value = "Non-Trigger"
    End With

    wbNonTrigger.Close SaveChanges:=False

    ' === Step 4: Populate 'Family' and 'ECA India Analyst' from 'Dataset' sheet ===

    ' Set up Dataset sheet
    Set wsDataset = wbMaster.Sheets("Dataset")

    ' Ensure the Dataset sheet is converted to a table
    On Error Resume Next
    Set datasetTable = wsDataset.ListObjects("DatasetTable")
    On Error GoTo 0
    If datasetTable Is Nothing Then
        Set datasetTable = wsDataset.ListObjects.Add(xlSrcRange, wsDataset.UsedRange, , xlYes)
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
        If Not dictDataset.exists(datasetFundManagerGCI(i, 1)) Then
            dictDataset.Add datasetFundManagerGCI(i, 1), Array(datasetFamily(i, 1), datasetECAAnalyst(i, 1))
        End If
    Next i

    ' Read Portfolio data into arrays
    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    portfolioFundManagerGCI = portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value
    portfolioFamily = portfolioTable.ListColumns("Family").DataBodyRange.Value
    portfolioECAAnalyst = portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Value

    ' Populate 'Family' and 'ECA India Analyst' in PortfolioTable
    For i = 1 To numRowsPortfolio
        ' Update 'Family' only if it is empty
        If IsEmpty(portfolioFamily(i, 1)) Or portfolioFamily(i, 1) = "" Then
            If dictDataset.exists(portfolioFundManagerGCI(i, 1)) Then
                portfolioFamily(i, 1) = dictDataset(portfolioFundManagerGCI(i, 1))(0)
            Else
                portfolioFamily(i, 1) = "No Match Found"
            End If
        End If

        ' Always update 'ECA India Analyst' from DatasetTable
        If dictDataset.exists(portfolioFundManagerGCI(i, 1)) Then
            portfolioECAAnalyst(i, 1) = dictDataset(portfolioFundManagerGCI(i, 1))(1)
        Else
            portfolioECAAnalyst(i, 1) = "No Match Found"
        End If
    Next i

    ' Write updated data back to PortfolioTable
    portfolioTable.ListColumns("Family").DataBodyRange.Value = portfolioFamily
    portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Value = portfolioECAAnalyst

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