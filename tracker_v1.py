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
    
    ' Set the Portfolio sheet
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
        portfolioTable.DataBodyRange.ClearContents
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
        ' Convert Trigger data to a table if it isn't already one
        Set triggerTable = triggerSheet.ListObjects.Add(xlSrcRange, triggerSheet.UsedRange, , xlYes)
        triggerTable.Name = "TriggerTable"
    End If
    
    ' Determine the starting row for pasting in the Portfolio table
    If portfolioTable.ListRows.Count = 0 Then
        pasteStartRow = portfolioTable.HeaderRowRange.Row + 1 ' Start immediately below headers
    Else
        pasteStartRow = portfolioTable.DataBodyRange.Rows(1).Row + portfolioTable.ListRows.Count
    End If

    ' Headers to copy from Trigger.csv
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer")
    
    ' Copy and paste data in bulk
    For i = LBound(headers) To UBound(headers)
        With triggerTable.ListColumns(headers(i)).DataBodyRange
            wsPortfolio.Cells(pasteStartRow, portfolioTable.ListColumns(headers(i)).Index).Resize(.Rows.Count, 1).Value = .Value
        End With
    Next i

    ' Hygiene changes
    With portfolioTable.DataBodyRange
        Dim regionColIndex As Long, triggerColIndex As Long
        Dim regionRange As Range, triggerRange As Range
        
        ' Get column indices for Region and Trigger/Non-Trigger
        regionColIndex = portfolioTable.ListColumns("Region").Index
        triggerColIndex = portfolioTable.ListColumns("Trigger/Non-Trigger").Index
        
        ' Get ranges for Region and Trigger/Non-Trigger columns
        Set regionRange = .Columns(regionColIndex)
        Set triggerRange = .Columns(triggerColIndex).Resize(triggerTable.ListRows.Count)
        
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