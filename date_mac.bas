Sub UpdatePortfolioTable_Optimized()
    Dim wbTrigger As Workbook, wbNonTrigger As Workbook
    Dim wsTrigger As Worksheet, wsNonTrigger As Worksheet, wsPortfolio As Worksheet
    Dim loTrigger As ListObject, loNonTrigger As ListObject, loPortfolio As ListObject
    Dim filePathTrigger As String, filePathNonTrigger As String
    Dim dictTrigger As Object, dictNonTrigger As Object
    Dim portRow As ListRow, key As Variant
    Dim flagValue As String
    Dim trgData As Variant, nonTrigData As Variant
    Dim calcMode As XlCalculation
    
    ' Improve performance by turning off screen updating, events and calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp
    
    ' Prompt user for both files at the start
    filePathTrigger = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Select Trigger File")
    If filePathTrigger = "False" Then
        MsgBox "No trigger file selected. Exiting."
        GoTo CleanUp
    End If
    
    filePathNonTrigger = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Select Non-Trigger File")
    If filePathNonTrigger = "False" Then
        MsgBox "No non-trigger file selected. Exiting."
        GoTo CleanUp
    End If
    
    ' Open both workbooks
    Set wbTrigger = Workbooks.Open(filePathTrigger)
    Set wbNonTrigger = Workbooks.Open(filePathNonTrigger)
    
    ' Set up Trigger file worksheet and table (assume first worksheet)
    Set wsTrigger = wbTrigger.Worksheets(1)
    On Error Resume Next
    Set loTrigger = wsTrigger.ListObjects(1)
    On Error GoTo 0
    If loTrigger Is Nothing Then
        Dim lastRow As Long, lastCol As Long
        lastRow = wsTrigger.Cells(wsTrigger.Rows.Count, 1).End(xlUp).Row
        lastCol = wsTrigger.Cells(1, wsTrigger.Columns.Count).End(xlToLeft).Column
        Set loTrigger = wsTrigger.ListObjects.Add(xlSrcRange, wsTrigger.Range(wsTrigger.Cells(1, 1), wsTrigger.Cells(lastRow, lastCol)), , xlYes)
    End If
    
    ' Set up Non-Trigger file worksheet and table (assume first worksheet)
    Set wsNonTrigger = wbNonTrigger.Worksheets(1)
    On Error Resume Next
    Set loNonTrigger = wsNonTrigger.ListObjects(1)
    On Error GoTo 0
    If loNonTrigger Is Nothing Then
        Dim lastRowNT As Long, lastColNT As Long
        lastRowNT = wsNonTrigger.Cells(wsNonTrigger.Rows.Count, 1).End(xlUp).Row
        lastColNT = wsNonTrigger.Cells(1, wsNonTrigger.Columns.Count).End(xlToLeft).Column
        Set loNonTrigger = wsNonTrigger.ListObjects.Add(xlSrcRange, wsNonTrigger.Range(wsNonTrigger.Cells(1, 1), wsNonTrigger.Cells(lastRowNT, lastColNT)), , xlYes)
    End If
    
    ' Build dictionary for Trigger file data
    Set dictTrigger = CreateObject("Scripting.Dictionary")
    Dim trgFundGCI_Col As Long, trgLatestNAV_Col As Long, trgReqNAV_Col As Long
    trgFundGCI_Col = loTrigger.ListColumns("Fund GCI").Index
    trgLatestNAV_Col = loTrigger.ListColumns("Latest NAV Date").Index
    ' Use the column "Req NAV Date" for updating PortfolioTable's Required NAV Date
    trgReqNAV_Col = loTrigger.ListColumns("Req NAV Date").Index
    
    Dim trgRow As Range, trgKey As Variant
    For Each trgRow In loTrigger.DataBodyRange.Rows
        trgKey = trgRow.Cells(1, trgFundGCI_Col).Value
        If Not dictTrigger.Exists(trgKey) Then
            dictTrigger.Add trgKey, Array( _
                trgRow.Cells(1, trgLatestNAV_Col).Value, _
                trgRow.Cells(1, trgReqNAV_Col).Value)
        End If
    Next trgRow
    
    ' Build dictionary for Non-Trigger file data
    Set dictNonTrigger = CreateObject("Scripting.Dictionary")
    Dim nonTrigFundGCI_Col As Long, nonTrigLatestNAV2_Col As Long, nonTrigRequiredNAV3_Col As Long
    nonTrigFundGCI_Col = loNonTrigger.ListColumns("Fund GCI").Index
    nonTrigLatestNAV2_Col = loNonTrigger.ListColumns("Latest NAV Date2").Index
    nonTrigRequiredNAV3_Col = loNonTrigger.ListColumns("Required NAV Date3").Index
    
    Dim nonTrigRow As Range, nonTrigKey As Variant
    For Each nonTrigRow In loNonTrigger.DataBodyRange.Rows
        nonTrigKey = nonTrigRow.Cells(1, nonTrigFundGCI_Col).Value
        If Not dictNonTrigger.Exists(nonTrigKey) Then
            dictNonTrigger.Add nonTrigKey, Array( _
                nonTrigRow.Cells(1, nonTrigLatestNAV2_Col).Value, _
                nonTrigRow.Cells(1, nonTrigRequiredNAV3_Col).Value)
        End If
    Next nonTrigRow
    
    ' Set reference to Portfolio table in the current workbook
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Set loPortfolio = wsPortfolio.ListObjects("PortfolioTable")
    
    Dim portFundGCI_Col As Long, portFlag_Col As Long, portLatestNAV_Col As Long, portRequiredNAV_Col As Long
    portFundGCI_Col = loPortfolio.ListColumns("Fund GCI").Index
    portFlag_Col = loPortfolio.ListColumns("Trigger/Non-Trigger").Index
    portLatestNAV_Col = loPortfolio.ListColumns("Latest NAV Date").Index
    portRequiredNAV_Col = loPortfolio.ListColumns("Required NAV Date").Index
    
    ' Loop through PortfolioTable and update rows based on the Trigger/Non-Trigger flag
    For Each portRow In loPortfolio.ListRows
        key = portRow.Range(1, portFundGCI_Col).Value
        flagValue = portRow.Range(1, portFlag_Col).Value
        
        If flagValue = "Trigger" Then
            If dictTrigger.Exists(key) Then
                trgData = dictTrigger(key)
                portRow.Range(1, portLatestNAV_Col).Value = trgData(0)
                portRow.Range(1, portRequiredNAV_Col).Value = trgData(1)
            End If
        ElseIf flagValue = "Non-Trigger" Then
            If dictNonTrigger.Exists(key) Then
                nonTrigData = dictNonTrigger(key)
                portRow.Range(1, portLatestNAV_Col).Value = nonTrigData(0)
                portRow.Range(1, portRequiredNAV_Col).Value = nonTrigData(1)
            End If
        End If
    Next portRow
    
    ' Close both workbooks without saving changes
    wbTrigger.Close False
    wbNonTrigger.Close False
    
    MsgBox "PortfolioTable updated successfully."
    
CleanUp:
    ' Restore application settings
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub