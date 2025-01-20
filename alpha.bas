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

    ' -- (A) Create the two new columns in PortfolioTable if they do not exist --
    Dim colCheck As ListColumn
    On Error Resume Next
    Set colCheck = portfolioTable.ListColumns("Latest NAV Date")
    If colCheck Is Nothing Then
        Set colCheck = portfolioTable.ListColumns.Add
        colCheck.Name = "Latest NAV Date"
    End If
    Set colCheck = Nothing

    Set colCheck = portfolioTable.ListColumns("Required NAV Date")
    If colCheck Is Nothing Then
        Set colCheck = portfolioTable.ListColumns.Add
        colCheck.Name = "Required NAV Date"
    End If
    On Error GoTo 0

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

    ' -- (B) Copy specific columns from Trigger.csv to Portfolio (including the new columns) --
    Dim sourceHeadersTrig As Variant, destHeadersTrig As Variant
    sourceHeadersTrig = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer", _
                              "Latest NAV Date", "Req NAV Date")
    destHeadersTrig = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer", _
                            "Latest NAV Date", "Required NAV Date")

    lastRowPortfolio = portfolioTable.HeaderRowRange.Row + 1
    For i = LBound(sourceHeadersTrig) To UBound(sourceHeadersTrig)
        On Error Resume Next
        ' Attempt to find the column in Trigger.csv
        Dim srcCol As ListColumn
        Dim srcColIndex As Long
        Set srcCol = triggerTable.ListColumns(sourceHeadersTrig(i))
        srcColIndex = 0
        If Not srcCol Is Nothing Then srcColIndex = srcCol.Index
        On Error GoTo 0

        If srcColIndex <> 0 Then
            ' Copy that column from Trigger.csv to the matching column in Portfolio
            wsPortfolio.Cells(lastRowPortfolio, portfolioTable.ListColumns(destHeadersTrig(i)).Index) _
                       .Resize(srcCol.DataBodyRange.Rows.Count, 1).Value = srcCol.DataBodyRange.Value
        End If
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

    ' Load Fund GCI and IA GCI from All Funds into a dictionary
    Set dictFundGCI = CreateObject("Scripting.Dictionary")
    Dim fundGCICol As Range, iaGCICol As Range
    Set fundGCICol = allFundsTable.ListColumns("Fund GCI").DataBodyRange
    Set iaGCICol = allFundsTable.ListColumns("IA GCI").DataBodyRange

    For i = 1 To fundGCICol.Rows.Count
        If Not IsEmpty(fundGCICol.Cells(i, 1).Value) And Not dictFundGCI.Exists(fundGCICol.Cells(i, 1).Value) Then
            dictFundGCI.Add fundGCICol.Cells(i, 1).Value, iaGCICol.Cells(i, 1).Value
        End If
    Next i

    ' Match Fund GCI in Portfolio and write IA GCI to Fund Manager GCI
    Dim lastRowPF As Long
    lastRowPF = portfolioTable.DataBodyRange.Rows.Count
    fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim resultArray(1 To lastRowPF, 1 To 1)

    For i = 1 To UBound(fundGCIArray, 1)
        If dictFundGCI.Exists(fundGCIArray(i, 1)) Then
            resultArray(i, 1) = dictFundGCI(fundGCIArray(i, 1))
        Else
            resultArray(i, 1) = "No Match Found"
        End If
    Next i

    ' Write results back to Fund Manager GCI column
    portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = resultArray

    ' Clear filters in All Funds table
    If allFundsTable.ShowAutoFilter Then allFundsTable.AutoFilter.ShowAllData

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

    ' (C) Append Non-Trigger data to Portfolio (including the new columns)
    Dim sourceHeaders As Variant, destHeaders As Variant
    sourceHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", _
                          "Credit Officer", "Weeks Missing", "Latest NAV Date", "Required NAV Date")
    destHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", _
                        "Credit Officer", "Wks Missing", "Latest NAV Date", "Required NAV Date")

    lastRowPortfolio = portfolioTable.DataBodyRange.Rows.Count + portfolioTable.HeaderRowRange.Row

    Dim numRowsToCopy As Long
    numRowsToCopy = nonTriggerTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count

    Dim arrIdx As Long
    For arrIdx = LBound(sourceHeaders) To UBound(sourceHeaders)
        On Error Resume Next
        Dim srcNC As ListColumn
        Dim srcNCIndex As Long
        Set srcNC = nonTriggerTable.ListColumns(sourceHeaders(arrIdx))
        srcNCIndex = 0
        If Not srcNC Is Nothing Then srcNCIndex = srcNC.Index
        On Error GoTo 0

        If srcNCIndex <> 0 Then
            With nonTriggerTable.ListColumns(sourceHeaders(arrIdx)).DataBodyRange.SpecialCells(xlCellTypeVisible)
                wsPortfolio.Cells(lastRowPortfolio + 1, _
                                  portfolioTable.ListColumns(destHeaders(arrIdx)).Index).Resize(.Rows.Count, 1).Value = .Value
            End With
        End If
    Next arrIdx

    ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
    With wsPortfolio
        .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), _
               .Cells(lastRowPortfolio + numRowsToCopy, triggerColIndex)).Value = "Non-Trigger"
    End With

    ' Clear filters in Non-Trigger table
    Dim nonTriggerLObj As ListObject
    Set nonTriggerLObj = nonTriggerTable
    If nonTriggerLObj.AutoFilter.FilterMode Then nonTriggerLObj.AutoFilter.ShowAllData

    wbNonTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv, All Funds.csv, and Non-Trigger.csv has been processed successfully!" & _
           vbCrLf & "New columns [Latest NAV Date] and [Required NAV Date] have been added.", vbInformation

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
