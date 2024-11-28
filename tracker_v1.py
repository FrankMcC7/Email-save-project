Sub ImportTriggerDataWithAutomaticTableCreation()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook
    Dim triggerFile As String
    Dim triggerSheet As Worksheet
    Dim triggerTable As ListObject
    Dim portfolioTable As ListObject
    Dim headers As Variant
    Dim colPortfolio As ListColumn
    Dim pasteRow As ListRow
    Dim i As Long
    
    ' Set the Portfolio sheet in the master file
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
    
    ' Prompt for the Trigger.csv file location
    triggerFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select Trigger.csv")
    If triggerFile = "False" Then Exit Sub ' User canceled

    ' Open the Trigger.csv file
    Set wbTrigger = Workbooks.Open(triggerFile)
    Set triggerSheet = wbTrigger.Sheets(1)
    
    ' Ensure the Trigger file data is converted to a table
    On Error Resume Next
    Set triggerTable = triggerSheet.ListObjects(1)
    On Error GoTo 0
    If triggerTable Is Nothing Then
        Set triggerTable = triggerSheet.ListObjects.Add(xlSrcRange, triggerSheet.UsedRange, , xlYes)
        triggerTable.Name = "TriggerTable"
    End If
    
    ' Headers to copy from Trigger.csv and paste into Portfolio
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer")
    
    ' Add new rows to the Portfolio table
    For i = 1 To triggerTable.ListRows.Count
        Set pasteRow = portfolioTable.ListRows.Add
        ' Copy data for each header
        For Each colPortfolio In portfolioTable.ListColumns
            On Error Resume Next
            pasteRow.Range(colPortfolio.Index).Value = triggerTable.ListColumns(colPortfolio.Name).DataBodyRange(i, 1).Value
            On Error GoTo 0
        Next colPortfolio
    Next i
    
    ' Hygiene changes
    With portfolioTable.DataBodyRange
        ' Replace 'US' with 'AMRS' and 'ASIA' with 'APAC' in the Region column
        For Each pasteRow In portfolioTable.ListRows
            Select Case pasteRow.Range(1).Value
                Case "US": pasteRow.Range(1).Value = "AMRS"
                Case "ASIA": pasteRow.Range(1).Value = "APAC"
            End Select
            ' Add 'Trigger' to the Trigger/Non-Trigger column
            pasteRow.Range(portfolioTable.ListColumns("Trigger/Non-Trigger").Index).Value = "Trigger"
        Next pasteRow
    End With

    ' Close the Trigger.csv file
    wbTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv has been successfully imported into the Portfolio table with hygiene changes.", vbInformation
End Sub