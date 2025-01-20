Option Explicit

Sub ProcessAllData()
    ' Variable declarations
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, allFundsTable As ListObject, nonTriggerTable As ListObject
    Dim dictFundGCI As Object, dictDataset As Object, dictLatestNAVDate As Object
    Dim fundGCIArray As Variant, resultArray As Variant, navDateArray As Variant
    Dim i As Long, lastRowPortfolio As Long
    Dim triggerColIndex As Long
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
    Dim baseFolder As String
    Dim headers As Variant
    Dim fundGCICol As Range, iaGCICol As Range, latestNAVDateCol As Range
    Dim lastRowPF As Long
    Dim sourceHeaders As Variant, destHeaders As Variant
    Dim numRowsToCopy As Long
    Dim arrIdx As Long
    Dim nonTriggerLObj As ListObject
    
    ' Define file paths - MODIFY THIS PATH TO MATCH YOUR ENVIRONMENT
    baseFolder = "C:\Data\NAV Reports\"  
    triggerFile = baseFolder & "Trigger.csv"
    allFundsFile = baseFolder & "All Funds.csv"
    nonTriggerFile = baseFolder & "Non-Trigger.csv"

    On Error GoTo ErrorHandler

    ' Verify file existence
    If Dir(triggerFile) = "" Then
        MsgBox "Trigger.csv not found at: " & triggerFile, vbCritical
        Exit Sub
    End If
    If Dir(allFundsFile) = "" Then
        MsgBox "All Funds.csv not found at: " & allFundsFile, vbCritical
        Exit Sub
    End If
    If Dir(nonTriggerFile) = "" Then
        MsgBox "Non-Trigger.csv not found at: " & nonTriggerFile, vbCritical
        Exit Sub
    End If

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
        Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, wsPortfolio.UsedRange, , xlYes)
        portfolioTable.Name = "PortfolioTable"
    End If

    ' Add new columns if they don't exist
    On Error Resume Next
    If portfolioTable.ListColumns("Latest NAV Date").Index = 0 Then
        portfolioTable.ListColumns.Add.Name = "Latest NAV Date"
    End If
    If portfolioTable.ListColumns("Required NAV Date").Index = 0 Then
        portfolioTable.ListColumns.Add.Name = "Required NAV Date"
    End If
    On Error GoTo 0

    ' Clear existing data in the Portfolio table except headers
    If Not portfolioTable.DataBodyRange Is Nothing Then
        portfolioTable.DataBodyRange.Delete
    End If

    ' === Step 1: Process Trigger.csv ===
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
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer", "Req NAV Date")
    lastRowPortfolio = portfolioTable.HeaderRowRange.Row + 1
    For i = LBound(headers) To UBound(headers)
        With triggerTable.ListColumns(headers(i)).DataBodyRange
            If headers(i) = "Req NAV Date" Then
                wsPortfolio.Cells(lastRowPortfolio, portfolioTable.ListColumns("Required NAV Date").Index).Resize(.Rows.Count, 1).Value = .Value
            Else
                wsPortfolio.Cells(lastRowPortfolio, portfolioTable.ListColumns(headers(i)).Index).Resize(.Rows.Count, 1).Value = .Value
            End If
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
    Set wbAllFunds = Workbooks.Open(allFundsFile)
    Set allFundsSheet = wbAllFunds.Sheets(1)

    ' Delete the first row
    allFundsSheet.Rows(1).Delete

    ' Convert All Funds data to a table
    On Error Resume Next
    Set allFundsTable = allFundsSheet.ListObjects(1)
    On Error GoTo 0
    If allFundsTable Is Nothing Then
        Set allFundsTable = allFundsSheet.ListObjects.Add(xlSrcRange, allFundsSheet.UsedRange, , xlYes)
        allFundsTable.Name = "AllFundsTable"
    End If

    ' Filter Review Status to Approved
    allFundsTable.Range.AutoFilter Field:=allFundsTable.ListColumns("Review Status").Index, Criteria1:="Approved"

    ' Load Fund GCI, IA GCI, and Latest NAV Date into dictionaries
    Set dictFundGCI = CreateObject("Scripting.Dictionary")
    Set dictLatestNAVDate = CreateObject("Scripting.Dictionary")
    Set fundGCICol = allFundsTable.ListColumns("Fund GCI").DataBodyRange
    Set iaGCICol = allFundsTable.ListColumns("IA GCI").DataBodyRange
    Set latestNAVDateCol = allFundsTable.ListColumns("Latest NAV Date").DataBodyRange

    For i = 1 To fundGCICol.Rows.Count
        If Not IsEmpty(fundGCICol.Cells(i, 1).Value) Then
            If Not dictFundGCI.exists(fundGCICol.Cells(i, 1).Value) Then
                dictFundGCI.Add fundGCICol.Cells(i, 1).Value, iaGCICol.Cells(i, 1).Value
            End If
            If Not dictLatestNAVDate.exists(fundGCICol.Cells(i, 1).Value) Then
                dictLatestNAVDate.Add fundGCICol.Cells(i, 1).Value, latestNAVDateCol.Cells(i, 1).Value
            End If
        End If
    Next i

    ' Match Fund GCI in Portfolio and write IA GCI and Latest NAV Date
    lastRowPF = portfolioTable.DataBodyRange.Rows.Count
    fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim resultArray(1 To lastRowPF, 1 To 1)
    ReDim navDateArray(1 To lastRowPF, 1 To 1)

    For i = 1 To UBound(fundGCIArray, 1)
        If dictFundGCI.exists(fundGCIArray(i, 1)) Then
            resultArray(i, 1) = dictFundGCI(fundGCIArray(i, 1))
            navDateArray(i, 1) = dictLatestNAVDate(fundGCIArray(i, 1))
        Else
            resultArray(i, 1) = "No Match Found"
            navDateArray(i, 1) = "No Match Found"
        End If
    Next i

    ' Write results back to Fund Manager GCI and Latest NAV Date columns
    portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = resultArray
    portfolioTable.ListColumns("Latest NAV Date").DataBodyRange.Value = navDateArray

    ' Clear filters in All Funds table
    If allFundsTable.ShowAutoFilter Then allFundsTable.AutoFilter.ShowAllData

    wbAllFunds.Close SaveChanges:=False

    ' === Step 3: Process Non-Trigger.csv ===
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
    sourceHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Weeks Missing", "Required NAV Date")
    destHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Wks Missing", "Required NAV Date")

    lastRowPortfolio = portfolioTable.DataBodyRange.Rows.Count + portfolioTable.HeaderRowRange.Row

    numRowsToCopy = nonTriggerTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count

    For arrIdx = LBound(sourceHeaders) To UBound(sourceHeaders)
        With nonTriggerTable.ListColumns(sourceHeaders(arrIdx)).DataBodyRange.SpecialCells(xlCellTypeVisible)
            wsPortfolio.Cells(lastRowPortfolio + 1, portfolioTable.ListColumns(destHeaders(arrIdx)).Index).Resize(.Rows.Count, 1).Value = .Value
        End With
    Next arrIdx

    ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
    With wsPortfolio
        .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), .Cells(lastRowPortfolio + numRowsToCopy, triggerColIndex)).Value = "Non-Trigger"
    End With

    ' Clear filters in Non-Trigger table
    Set nonTriggerLObj = nonTriggerTable
    If nonTriggerLObj.AutoFilter.FilterMode Then nonTriggerLObj.AutoFilter.ShowAllData

    wbNonTrigger.Close SaveChanges:=False

    MsgBox "Data processed successfully from:" & vbCrLf & _
           "- " & triggerFile & vbCrLf & _
           "- " & allFundsFile & vbCrLf & _
           "- " & nonTriggerFile, vbInformation

CleanUp:
    ' Reset Application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Clean up objects
    Set wsPortfolio = Nothing
    Set wbMaster = Nothing
    Set wbTrigger = Nothing
    Set wbAllFunds = Nothing
    Set wbNonTrigger = Nothing
    Set triggerSheet = Nothing
    Set allFundsSheet = Nothing
    Set nonTriggerSheet = Nothing
    Set triggerTable = Nothing
    Set portfolioTable = Nothing
    Set allFundsTable = Nothing
    Set nonTriggerTable = Nothing
    Set dictFundGCI = Nothing
    Set dictDataset = Nothing
    Set dictLatestNAVDate = Nothing
    Set fundGCICol = Nothing
    Set iaGCICol = Nothing
    Set latestNAVDateCol = Nothing
    Set nonTriggerLObj = Nothing

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Current file being processed: " & _
           IIf(wbTrigger Is Nothing, _
               IIf(wbAllFunds Is Nothing, _
                   IIf(wbNonTrigger Is Nothing, "Unknown", nonTriggerFile), _
                   allFundsFile), _
               triggerFile), vbCritical
    Resume CleanUp
End Sub
