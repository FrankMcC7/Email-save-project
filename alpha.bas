Sub ProcessAllData()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, allFundsTable As ListObject, nonTriggerTable As ListObject
    
    ' Dictionaries for IA GCI and Latest NAV Date
    Dim dictIA_GCI As Object
    Dim dictNAVDate As Object
    
    Dim fundGCIArray As Variant
    Dim resultArray As Variant
    Dim i As Long, lastRowPortfolio As Long
    Dim triggerColIndex As Long
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

    '------------------------------------------------------------------------------
    ' (A) ADD TWO NEW COLUMNS IF THEY DON'T EXIST ALREADY:
    '       1. "Latest NAV Date"
    '       2. "Required NAV Date"
    '------------------------------------------------------------------------------
    Dim newCol As ListColumn
    
    On Error Resume Next
    Set newCol = portfolioTable.ListColumns("Latest NAV Date")
    If newCol Is Nothing Then
        Set newCol = portfolioTable.ListColumns.Add
        newCol.Name = "Latest NAV Date"
    End If
    
    Set newCol = portfolioTable.ListColumns("Required NAV Date")
    If newCol Is Nothing Then
        Set newCol = portfolioTable.ListColumns.Add
        newCol.Name = "Required NAV Date"
    End If
    On Error GoTo 0

    '------------------------------------------------------------------------------
    ' STEP 1: PROCESS Trigger.csv
    '------------------------------------------------------------------------------
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

    '--- RENAME "Req NAV Date" COLUMN TO "Required NAV Date" (so we can copy it 1:1) ---
    On Error Resume Next
    triggerTable.ListColumns("Req NAV Date").Name = "Required NAV Date"
    On Error GoTo 0

    ' Copy specific columns from Trigger.csv to Portfolio (including "Required NAV Date" now)
    Dim trigHeaders As Variant
    trigHeaders = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer", "Required NAV Date")

    ' Where to paste in the Portfolio
    lastRowPortfolio = portfolioTable.HeaderRowRange.Row + 1

    For i = LBound(trigHeaders) To UBound(trigHeaders)
        With triggerTable.ListColumns(trigHeaders(i)).DataBodyRange
            wsPortfolio.Cells(lastRowPortfolio, portfolioTable.ListColumns(trigHeaders(i)).Index) _
                .Resize(.Rows.Count, 1).Value = .Value
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

    '------------------------------------------------------------------------------
    ' STEP 2: PROCESS All Funds.csv
    '   - Filter "Approved"
    '   - Dictionary match for IA GCI and LATEST NAV DATE via "Fund GCI"
    '------------------------------------------------------------------------------
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

    ' --- Create Dictionaries ---
    Set dictIA_GCI = CreateObject("Scripting.Dictionary")
    Set dictNAVDate = CreateObject("Scripting.Dictionary")

    Dim fundGCICol As Range, iaGCICol As Range, navDateCol As Range
    Set fundGCICol = allFundsTable.ListColumns("Fund GCI").DataBodyRange
    Set iaGCICol = allFundsTable.ListColumns("IA GCI").DataBodyRange
    Set navDateCol = allFundsTable.ListColumns("Latest NAV Date").DataBodyRange  ' Make sure the column name matches EXACTLY

    ' --- Read each row into dictionaries by Fund GCI ---
    Dim keyVal As Variant
    For i = 1 To fundGCICol.Rows.Count
        keyVal = fundGCICol.Cells(i, 1).Value
        If Not IsEmpty(keyVal) Then
            If Not dictIA_GCI.Exists(keyVal) Then
                dictIA_GCI.Add keyVal, iaGCICol.Cells(i, 1).Value
            End If
            If Not dictNAVDate.Exists(keyVal) Then
                dictNAVDate.Add keyVal, navDateCol.Cells(i, 1).Value
            End If
        End If
    Next i

    ' 1) Match Fund GCI in Portfolio and write IA GCI to Fund Manager GCI
    Dim lastRowPF As Long
    lastRowPF = portfolioTable.DataBodyRange.Rows.Count
    fundGCIArray = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value

    ReDim resultArray(1 To lastRowPF, 1 To 1)

    For i = 1 To UBound(fundGCIArray, 1)
        If dictIA_GCI.exists(fundGCIArray(i, 1)) Then
            resultArray(i, 1) = dictIA_GCI(fundGCIArray(i, 1))
        Else
            resultArray(i, 1) = "No Match Found"
        End If
    Next i
    portfolioTable.ListColumns("Fund Manager GCI").DataBodyRange.Value = resultArray

    ' 2) Match Fund GCI in Portfolio and write Latest NAV Date
    ReDim resultArray(1 To lastRowPF, 1 To 1)
    For i = 1 To UBound(fundGCIArray, 1)
        If dictNAVDate.exists(fundGCIArray(i, 1)) Then
            resultArray(i, 1) = dictNAVDate(fundGCIArray(i, 1))
        Else
            resultArray(i, 1) = ""  ' or "No Match Found", up to you
        End If
    Next i
    portfolioTable.ListColumns("Latest NAV Date").DataBodyRange.Value = resultArray

    ' Clear filters in All Funds table
    If allFundsTable.ShowAutoFilter Then allFundsTable.AutoFilter.ShowAllData

    wbAllFunds.Close SaveChanges:=False

    '------------------------------------------------------------------------------
    ' STEP 3: PROCESS Non-Trigger.csv
    '   - Filter out "FI-ASIA" in Region
    '   - Copy columns (including "Required NAV Date") to Portfolio
    '------------------------------------------------------------------------------
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
    sourceHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Weeks Missing", "Required NAV Date")
    destHeaders = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", "Fund GCI", "Fund Name", "Credit Officer", "Wks Missing", "Required NAV Date")

    lastRowPortfolio = portfolioTable.DataBodyRange.Rows.Count + portfolioTable.HeaderRowRange.Row

    Dim numRowsToCopy As Long
    numRowsToCopy = nonTriggerTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count

    Dim arrIdx As Long
    For arrIdx = LBound(sourceHeaders) To UBound(sourceHeaders)
        With nonTriggerTable.ListColumns(sourceHeaders(arrIdx)).DataBodyRange.SpecialCells(xlCellTypeVisible)
            wsPortfolio.Cells(lastRowPortfolio + 1, portfolioTable.ListColumns(destHeaders(arrIdx)).Index) _
                .Resize(.Rows.Count, 1).Value = .Value
        End With
    Next arrIdx

    ' Fill Trigger/Non-Trigger column with 'Non-Trigger'
    With wsPortfolio
        .Range(.Cells(lastRowPortfolio + 1, triggerColIndex), .Cells(lastRowPortfolio + numRowsToCopy, triggerColIndex)).Value = "Non-Trigger"
    End With

    ' Clear filters in Non-Trigger table
    If nonTriggerTable.AutoFilter.FilterMode Then
        nonTriggerTable.AutoFilter.ShowAllData
    End If

    wbNonTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv, All Funds.csv, and Non-Trigger.csv has been processed successfully!" & vbCrLf & _
           "Two new columns added: 'Latest NAV Date' and 'Required NAV Date'.", vbInformation

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
