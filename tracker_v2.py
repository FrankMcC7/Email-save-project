Sub ProcessAllData()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, allFundsTable As ListObject, nonTriggerTable As ListObject
    Dim dictFundGCI As Object
    Dim fundGCIArray As Variant, resultArray As Variant
    Dim i As Long, lastRowPortfolio As Long
    Dim triggerColIndex As Long ' Single declaration for use across sections
    Dim startTime As Double
    
    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Record the start time
    startTime = Timer
    
    ' Set up the Portfolio sheet
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

    ' Delete the first row
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
    Dim portfolioFundGCI As Variant
    Dim portfolioIA_GCI As Variant
    Dim numRowsPortfolio As Long

    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    portfolioFundGCI = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim portfolioIA_GCI(1 To numRowsPortfolio, 1 To 1)

    For i = 1 To numRowsPortfolio
        If dictFundGCI.exists(portfolioFundGCI(i, 1)) Then
            portfolioIA_GCI(i, 1) = dictFundGCI(portfolioFundGCI(i, 1))
        Else
            portfolioIA_GCI(i, 1) = "No Match Found"
        End If
    Next i

    ' Write results back to Fund Manager GCI column
    portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = portfolioIA_GCI

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

    ' Clear filters in Non-Trigger table
    If nonTriggerTable.ShowAutoFilter Then nonTriggerTable.AutoFilter.ShowAllData

    wbNonTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv, All Funds.csv, and Non-Trigger.csv has been processed successfully!" & vbCrLf & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds.", vbInformation

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
