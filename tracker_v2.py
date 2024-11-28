Sub ProcessAllData()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, nonTriggerTable As ListObject
    Dim dictFundGCI As Object, dictDataset As Object
    Dim fundGCIArray As Variant
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

    ' Clear the sheet except for headers
    With wsPortfolio
        If .AutoFilterMode Then .AutoFilterMode = False
        If .FilterMode Then .ShowAllData

        Dim lastUsedRow As Long
        lastUsedRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        ' Assuming headers are in row 1
        If lastUsedRow > 1 Then
            .Rows("2:" & lastUsedRow).ClearContents
        End If
    End With

    ' Ensure the Portfolio sheet is converted to a table
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    If portfolioTable Is Nothing Then
        ' Check if headers exist in row 1
        Dim headerRow As Range
        Set headerRow = wsPortfolio.Rows(1)
        If Application.WorksheetFunction.CountA(headerRow) > 0 Then
            ' Determine the last used column
            Dim lastCol As Long
            lastCol = wsPortfolio.Cells(1, wsPortfolio.Columns.Count).End(xlToLeft).Column

            ' Define the range including only the header row
            Dim dataRange As Range
            Set dataRange = wsPortfolio.Range(wsPortfolio.Cells(1, 1), wsPortfolio.Cells(1, lastCol))

            ' Convert the header row to a table
            Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
            portfolioTable.Name = "PortfolioTable"
        Else
            MsgBox "No headers found in the Portfolio sheet.", vbCritical
            GoTo CleanUp
        End If
    Else
        ' If the table exists, ensure it starts from row 1
        If portfolioTable.HeaderRowRange.Row > 1 Then
            MsgBox "The PortfolioTable does not start at row 1. Please adjust the sheet.", vbCritical
            GoTo CleanUp
        End If
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
            If Not IsEmpty(allFundsData(i, fundGCIColIndex)) And Not dictFundGCI.Exists(allFundsData(i, fundGCIColIndex)) Then
                dictFundGCI.Add allFundsData(i, fundGCIColIndex), allFundsData(i, iaGCIColIndex)
            End If
        End If
    Next i

    ' Match Fund GCI in Portfolio and write IA GCI to Fund Manager GCI
    numRowsPortfolio = portfolioTable.ListRows.Count
    fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim portfolioFundManagerGCI(1 To numRowsPortfolio, 1 To 1)

    For i = 1 To numRowsPortfolio
        If dictFundGCI.Exists(fundGCIArray(i, 1)) Then
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
    Set nonTriggerTable = nonTriggerSheet.ListObjects("NonTriggerTable")
    On Error GoTo 0
    If nonTriggerTable Is Nothing Then
        Set nonTriggerTable = nonTriggerSheet.ListObjects.Add(xlSrcRange, nonTriggerSheet.UsedRange, , xlYes)
        nonTriggerTable.Name = "NonTriggerTable"
    End If

    ' Clear any existing filters
    If nonTriggerTable.AutoFilter.FilterMode Then
        nonTriggerTable.AutoFilter.ShowAllData
    End If

    ' Verify the "Business Unit" column exists
    Dim businessUnitCol As ListColumn
    Set businessUnitCol = Nothing
    On Error Resume Next
    Set businessUnitCol = nonTriggerTable.ListColumns("Business Unit")
    On Error GoTo 0

    If businessUnitCol Is Nothing Then
        MsgBox "The column 'Business Unit' does not exist in Non-Trigger.csv.", vbCritical
        GoTo CleanUp
    End If

    ' Apply the filter using wildcards to account for variations
    nonTriggerTable.Range.AutoFilter Field:=businessUnitCol.Index, Criteria1:="<>*FI-ASIA*", Operator:=xlAnd

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

    Dim destRow As Long
    destRow = lastRowPortfolio + 1

    Dim visibleRowCount As Long
    visibleRowCount = 0

    ' Loop through each area of the visible data
    Dim area As Range
    Dim rowOffset As Long
    rowOffset = 0

    For i = LBound(sourceHeaders) To UBound(sourceHeaders)
        Dim sourceColumn As Range
        On Error Resume Next
        Set sourceColumn = nonTriggerTable.ListColumns(sourceHeaders(i)).DataBodyRange.SpecialCells(xlCellTypeVisible)
        If Err.Number <> 0 Then
            ' No visible cells in this column
            Err.Clear
            GoTo NextHeader
        End If
        On Error GoTo 0

        ' Loop through each area of visible cells
        For Each area In sourceColumn.Areas
            wsPortfolio.Cells(destRow + rowOffset, portfolioTable.ListColumns(destHeaders(i)).Index).Resize(area.Rows.Count, 1).Value = area.Value
            If i = LBound(sourceHeaders) Then
                visibleRowCount = visibleRowCount + area.Rows.Count
                rowOffset = rowOffset + area.Rows.Count
            End If
        Next area
        rowOffset = 0 ' Reset for next column
NextHeader:
    Next i

    ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
    If visibleRowCount > 0 Then
        With wsPortfolio
            .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), .Cells(lastRowPortfolio + visibleRowCount, triggerColIndex)).Value = "Non-Trigger"
        End With
    Else
        MsgBox "No data to copy from Non-Trigger.csv after applying the filter.", vbExclamation
    End If

    ' Clear filters
    If nonTriggerTable.AutoFilter.FilterMode Then
        nonTriggerTable.AutoFilter.ShowAllData
    End If

    wbNonTrigger.Close SaveChanges:=False

    ' === Remove Extra Empty Rows from PortfolioTable ===
    ' Optional: Remove empty rows to ensure DataBodyRange is accurate
    Dim tblRow As ListRow
    For i = portfolioTable.ListRows.Count To 1 Step -1
        Set tblRow = portfolioTable.ListRows(i)
        If Application.WorksheetFunction.CountA(tblRow.Range) = 0 Then
            tblRow.Delete
        End If
    Next i

    ' === Step 4: Populate 'Family' and 'ECA India Analyst' from 'Dataset' sheet ===

    ' Set up Dataset sheet
    Set wsDataset = wbMaster.Sheets("Dataset")

    ' Ensure the Dataset sheet is converted to a table
    ' [Previous code remains unchanged]

    ' Read Dataset data into arrays
    ' [Previous code remains unchanged]

    ' Create a dictionary for Dataset using Fund Manager GCI as key
    ' [Previous code remains unchanged]

    ' === Adjusted Code Starts Here ===

    ' Find the last row with data in 'Fund GCI' column of PortfolioTable
    Dim fundGCIRange As Range
    Set fundGCIRange = portfolioTable.ListColumns("Fund GCI").DataBodyRange

    Dim lastDataRow As Long
    With fundGCIRange
        If Application.WorksheetFunction.CountA(.Cells) > 0 Then
            lastDataRow = .Cells(.Rows.Count).End(xlUp).Row - .Row + 1
        Else
            lastDataRow = 0
        End If
    End With

    If lastDataRow = 0 Then
        MsgBox "No data in PortfolioTable to update.", vbInformation
        GoTo CleanUp
    End If

    numRowsPortfolio = lastDataRow

    ' Now read the data into arrays
    portfolioFundManagerGCI = portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Cells(1, 1).Resize(numRowsPortfolio, 1).Value
    portfolioFamily = portfolioTable.ListColumns("Family").DataBodyRange.Cells(1, 1).Resize(numRowsPortfolio, 1).Value
    portfolioECAAnalyst = portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Cells(1, 1).Resize(numRowsPortfolio, 1).Value

    ' Populate 'Family' and 'ECA India Analyst' in PortfolioTable
    For i = 1 To numRowsPortfolio
        ' Update 'Family' only if it is empty
        If IsEmpty(portfolioFamily(i, 1)) Or portfolioFamily(i, 1) = "" Then
            If dictDataset.Exists(portfolioFundManagerGCI(i, 1)) Then
                portfolioFamily(i, 1) = dictDataset(portfolioFundManagerGCI(i, 1))(0)
            Else
                portfolioFamily(i, 1) = "No Match Found"
            End If
        End If

        ' Always update 'ECA India Analyst' from DatasetTable
        If dictDataset.Exists(portfolioFundManagerGCI(i, 1)) Then
            portfolioECAAnalyst(i, 1) = dictDataset(portfolioFundManagerGCI(i, 1))(1)
        Else
            portfolioECAAnalyst(i, 1) = "No Match Found"
        End If
    Next i

    ' Write updated data back to PortfolioTable
    portfolioTable.ListColumns("Family").DataBodyRange.Cells(1, 1).Resize(numRowsPortfolio, 1).Value = portfolioFamily
    portfolioTable.ListColumns("ECA India Analyst").DataBodyRange.Cells(1, 1).Resize(numRowsPortfolio, 1).Value = portfolioECAAnalyst

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
