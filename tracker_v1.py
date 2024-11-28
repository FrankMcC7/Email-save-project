Sub ImportTriggerDataWithAccurateRowDetection()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook
    Dim triggerFile As String
    Dim triggerSheet As Worksheet
    Dim triggerTable As ListObject
    Dim portfolioTable As ListObject
    Dim headers As Variant
    Dim pasteStartRow As Long
    Dim i As Long
    Dim lastCol As Long
    Dim portfolioHeaderRange As Range
    Dim triggerDataRange As Range
    Dim triggerLastRow As Long, triggerLastCol As Long
    Dim dataRowCount As Long
    
    ' Set the Portfolio sheet
    Set wbMaster = ThisWorkbook
    Set wsPortfolio = wbMaster.Sheets("Portfolio")
    
    ' Ensure the Portfolio sheet is converted to a table
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    If portfolioTable Is Nothing Then
        ' Determine the last used column in row 1 (headers)
        lastCol = wsPortfolio.Cells(1, wsPortfolio.Columns.Count).End(xlToLeft).Column
        ' Define the header range
        Set portfolioHeaderRange = wsPortfolio.Range(wsPortfolio.Cells(1, 1), wsPortfolio.Cells(1, lastCol))
        ' Convert the header range to a table
        Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, portfolioHeaderRange, , xlYes)
        portfolioTable.Name = "PortfolioTable"
    Else
        ' Clear existing data in the Portfolio table except headers
        If portfolioTable.ListRows.Count > 0 Then
            portfolioTable.DataBodyRange.Delete
        End If
    End If

    ' Clear any formatting or data beyond the table
    wsPortfolio.Cells(portfolioTable.Range.Row + portfolioTable.Range.Rows.Count, 1).Resize(wsPortfolio.Rows.Count - portfolioTable.Range.Row - portfolioTable.Range.Rows.Count + 1, wsPortfolio.Columns.Count).Clear

    ' Set pasteStartRow to the row immediately after the header row
    pasteStartRow = portfolioTable.HeaderRowRange.Row + 1 ' Should be row 2
    
    ' Prompt for the Trigger.csv file location
    triggerFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select Trigger.csv")
    If triggerFile = "False" Then Exit Sub ' User canceled
    
    ' Open the Trigger.csv file
    Set wbTrigger = Workbooks.Open(triggerFile)
    Set triggerSheet = wbTrigger.Sheets(1)
    
    ' Ensure the Trigger file data is converted to a table
    On Error Resume Next
    Set triggerTable = triggerSheet.ListObjects("TriggerTable")
    On Error GoTo 0
    If triggerTable Is Nothing Then
        ' Determine the last used row and column in Trigger sheet
        triggerLastRow = triggerSheet.Cells(triggerSheet.Rows.Count, 1).End(xlUp).Row
        triggerLastCol = triggerSheet.Cells(1, triggerSheet.Columns.Count).End(xlToLeft).Column
        ' Define the data range
        Set triggerDataRange = triggerSheet.Range(triggerSheet.Cells(1, 1), triggerSheet.Cells(triggerLastRow, triggerLastCol))
        ' Convert Trigger data to a table
        Set triggerTable = triggerSheet.ListObjects.Add(xlSrcRange, triggerDataRange, , xlYes)
        triggerTable.Name = "TriggerTable"
    End If
    
    ' Headers to copy from Trigger.csv
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer")
    
    ' Copy and paste data in bulk
    For i = LBound(headers) To UBound(headers)
        ' Check if the header exists in both tables
        If Not triggerTable.ListColumns(headers(i)) Is Nothing And Not portfolioTable.ListColumns(headers(i)) Is Nothing Then
            With triggerTable.ListColumns(headers(i)).DataBodyRange
                ' Paste data starting from pasteStartRow
                wsPortfolio.Cells(pasteStartRow, portfolioTable.ListColumns(headers(i)).Index).Resize(.Rows.Count, 1).Value = .Value
            End With
        Else
            MsgBox "Header '" & headers(i) & "' not found in both tables.", vbCritical
            Exit Sub
        End If
    Next i
    
    ' Hygiene changes
    With portfolioTable.Range
        Dim regionColIndex As Long, triggerColIndex As Long
        Dim regionRange As Range, triggerRange As Range
        
        ' Get column indices for Region and Trigger/Non-Trigger
        regionColIndex = portfolioTable.ListColumns("Region").Index
        triggerColIndex = portfolioTable.ListColumns("Trigger/Non-Trigger").Index
        
        ' Calculate the number of data rows
        dataRowCount = triggerTable.DataBodyRange.Rows.Count
        
        ' Get ranges for Region and Trigger/Non-Trigger columns
        Set regionRange = wsPortfolio.Cells(pasteStartRow, regionColIndex).Resize(dataRowCount, 1)
        Set triggerRange = wsPortfolio.Cells(pasteStartRow, triggerColIndex).Resize(dataRowCount, 1)
        
        ' Replace 'US' with 'AMRS' and 'ASIA' with 'APAC' in Region column
        regionRange.Replace What:="US", Replacement:="AMRS", LookAt:=xlWhole
        regionRange.Replace What:="ASIA", Replacement:="APAC", LookAt:=xlWhole
        
        ' Fill 'Trigger' in the Trigger/Non-Trigger column for all rows
        triggerRange.Value = "Trigger"
    End With
    
    ' Close the Trigger.csv file
    wbTrigger.Close SaveChanges:=False
    
    MsgBox "Portfolio table has been reset, and data from Trigger.csv has been successfully imported with hygiene changes.", vbInformation
End Sub