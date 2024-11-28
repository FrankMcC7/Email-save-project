Sub ImportTriggerDataWithOptimizedHygiene()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook
    Dim triggerFile As String
    Dim triggerSheet As Worksheet
    Dim triggerTable As ListObject
    Dim portfolioTable As ListObject
    Dim headers As Variant
    Dim colPortfolio As ListColumn
    Dim i As Long
    
    ' Set the Portfolio sheet and table in the master file
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
        With portfolioTable.ListRows.Add
            For Each colPortfolio In portfolioTable.ListColumns
                On Error Resume Next
                .Range(colPortfolio.Index).Value = triggerTable.ListColumns(colPortfolio.Name).DataBodyRange(i, 1).Value
                On Error GoTo 0
            Next colPortfolio
        End With
    Next i
    
    ' Optimized hygiene changes
    With portfolioTable.DataBodyRange
        Dim regionColIndex As Long, triggerColIndex As Long
        Dim regionRange As Range, triggerRange As Range
        
        ' Get column indices for Region and Trigger/Non-Trigger
        regionColIndex = portfolioTable.ListColumns("Region").Index
        triggerColIndex = portfolioTable.ListColumns("Trigger/Non-Trigger").Index
        
        ' Get ranges for Region and Trigger/Non-Trigger columns
        Set regionRange = .Columns(regionColIndex)
        Set triggerRange = .Columns(triggerColIndex)
        
        ' Replace 'US' with 'AMRS' and 'ASIA' with 'APAC' in Region column
        regionRange.Replace What:="US", Replacement:="AMRS", LookAt:=xlWhole
        regionRange.Replace What:="ASIA", Replacement:="APAC", LookAt:=xlWhole
        
        ' Fill 'Trigger' in the Trigger/Non-Trigger column for all rows
        triggerRange.Value = "Trigger"
    End With

    ' Close the Trigger.csv file
    wbTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv has been successfully imported into the Portfolio table with hygiene changes.", vbInformation
End Sub