Sub UpdatePortfolioTable()
    Dim wbTrigger As Workbook, wbNonTrigger As Workbook
    Dim wsTrigger As Worksheet, wsNonTrigger As Worksheet, wsPortfolio As Worksheet
    Dim loTrigger As ListObject, loNonTrigger As ListObject, loPortfolio As ListObject
    Dim filePath As String
    Dim portRow As ListRow
    Dim key As Variant
    Dim trgMatch As Variant, nonTrigMatch As Variant
    Dim lastRow As Long, lastCol As Long
    Dim lastRowNT As Long, lastColNT As Long
    
    Application.ScreenUpdating = False
    
    ' Set reference to the Portfolio sheet and its table
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set loPortfolio = wsPortfolio.ListObjects("PortfolioTable")
    
    ' ************* Part 1: Update Trigger rows *************
    ' Prompt user to locate the trigger file
    filePath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*")
    If filePath = "False" Then
        MsgBox "No trigger file selected. Exiting."
        Exit Sub
    End If
    
    Set wbTrigger = Workbooks.Open(filePath)
    ' Assume data is in the first worksheet
    Set wsTrigger = wbTrigger.Worksheets(1)
    
    ' If no table exists in the trigger file, convert the used range to a table
    On Error Resume Next
    Set loTrigger = wsTrigger.ListObjects(1)
    On Error GoTo 0
    If loTrigger Is Nothing Then
        lastRow = wsTrigger.Cells(wsTrigger.Rows.Count, 1).End(xlUp).Row
        lastCol = wsTrigger.Cells(1, wsTrigger.Columns.Count).End(xlToLeft).Column
        Set loTrigger = wsTrigger.ListObjects.Add(xlSrcRange, wsTrigger.Range(wsTrigger.Cells(1, 1), wsTrigger.Cells(lastRow, lastCol)), , xlYes)
    End If
    
    ' Get column indexes from the trigger file table
    Dim trgFundGCI_Col As Long, trgLatestNAV_Col As Long, trgRequiredNAV_Col As Long
    trgFundGCI_Col = loTrigger.ListColumns("Fund GCI").Index
    trgLatestNAV_Col = loTrigger.ListColumns("Latest NAV Date").Index
    trgRequiredNAV_Col = loTrigger.ListColumns("Required NAV Date").Index
    
    ' Get column indexes from the PortfolioTable
    Dim portFundGCI_Col As Long, portFlag_Col As Long, portLatestNAV_Col As Long, portRequiredNAV_Col As Long
    portFundGCI_Col = loPortfolio.ListColumns("Fund GCI").Index
    portFlag_Col = loPortfolio.ListColumns("Trigger/Non-Trigger").Index
    portLatestNAV_Col = loPortfolio.ListColumns("Latest NAV Date").Index
    portRequiredNAV_Col = loPortfolio.ListColumns("Required NAV Date").Index
    
    ' Loop through PortfolioTable rows where Trigger/Non-Trigger = "Trigger"
    For Each portRow In loPortfolio.ListRows
        If portRow.Range(1, portFlag_Col).Value = "Trigger" Then
            key = portRow.Range(1, portFundGCI_Col).Value
            On Error Resume Next
            trgMatch = Application.Match(key, loTrigger.ListColumns("Fund GCI").DataBodyRange, 0)
            On Error GoTo 0
            If Not IsError(trgMatch) Then
                ' Update using corresponding columns from trigger file
                portRow.Range(1, portLatestNAV_Col).Value = loTrigger.DataBodyRange(trgMatch, trgLatestNAV_Col).Value
                portRow.Range(1, portRequiredNAV_Col).Value = loTrigger.DataBodyRange(trgMatch, trgRequiredNAV_Col).Value
            End If
        End If
    Next portRow
    
    wbTrigger.Close False  ' Close trigger file without saving
    
    ' ************* Part 2: Update Non-Trigger rows *************
    ' Prompt user to locate the Non-Trigger file
    filePath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*")
    If filePath = "False" Then
        MsgBox "No Non-Trigger file selected. Exiting."
        Exit Sub
    End If
    
    Set wbNonTrigger = Workbooks.Open(filePath)
    Set wsNonTrigger = wbNonTrigger.Worksheets(1)
    
    ' Convert data to table if not already a table
    On Error Resume Next
    Set loNonTrigger = wsNonTrigger.ListObjects(1)
    On Error GoTo 0
    If loNonTrigger Is Nothing Then
        lastRowNT = wsNonTrigger.Cells(wsNonTrigger.Rows.Count, 1).End(xlUp).Row
        lastColNT = wsNonTrigger.Cells(1, wsNonTrigger.Columns.Count).End(xlToLeft).Column
        Set loNonTrigger = wsNonTrigger.ListObjects.Add(xlSrcRange, wsNonTrigger.Range(wsNonTrigger.Cells(1, 1), wsNonTrigger.Cells(lastRowNT, lastColNT)), , xlYes)
    End If
    
    ' Get column indexes for Non-Trigger file table
    Dim nonTrigFundGCI_Col As Long, nonTrigLatestNAV2_Col As Long, nonTrigRequiredNAV3_Col As Long
    nonTrigFundGCI_Col = loNonTrigger.ListColumns("Fund GCI").Index
    nonTrigLatestNAV2_Col = loNonTrigger.ListColumns("Latest NAV Date2").Index
    nonTrigRequiredNAV3_Col = loNonTrigger.ListColumns("Required NAV Date3").Index
    
    ' Loop through PortfolioTable rows where Trigger/Non-Trigger = "Non-Trigger"
    For Each portRow In loPortfolio.ListRows
        If portRow.Range(1, portFlag_Col).Value = "Non-Trigger" Then
            key = portRow.Range(1, portFundGCI_Col).Value
            On Error Resume Next
            nonTrigMatch = Application.Match(key, loNonTrigger.ListColumns("Fund GCI").DataBodyRange, 0)
            On Error GoTo 0
            If Not IsError(nonTrigMatch) Then
                ' Update using mapping from Non-Trigger file:
                ' Portfolio Latest NAV Date from "Latest NAV Date2"
                ' Portfolio Required NAV Date from "Required NAV Date3"
                portRow.Range(1, portLatestNAV_Col).Value = loNonTrigger.DataBodyRange(nonTrigMatch, nonTrigLatestNAV2_Col).Value
                portRow.Range(1, portRequiredNAV_Col).Value = loNonTrigger.DataBodyRange(nonTrigMatch, nonTrigRequiredNAV3_Col).Value
            End If
        End If
    Next portRow
    
    wbNonTrigger.Close False  ' Close non-trigger file without saving
    
    Application.ScreenUpdating = True
    MsgBox "PortfolioTable updated successfully."
End Sub