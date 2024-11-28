Sub ImportTriggerDataWithHygiene()
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook
    Dim triggerFile As String
    Dim triggerSheet As Worksheet
    Dim lastRowTrigger As Long, pasteRow As Long
    Dim headers As Variant
    Dim headerDict As Object
    Dim colTrigger As Long, colPortfolio As Long
    Dim regionCol As Long, triggerCol As Long
    Dim i As Long
    
    ' Set the Portfolio sheet in the master file
    Set wbMaster = ThisWorkbook
    Set wsPortfolio = wbMaster.Sheets("Portfolio")
    
    ' Prompt for the Trigger.csv file location
    triggerFile = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select Trigger.csv")
    If triggerFile = "False" Then Exit Sub ' User canceled

    ' Open the Trigger.csv file
    Set wbTrigger = Workbooks.Open(triggerFile)
    Set triggerSheet = wbTrigger.Sheets(1)
    
    ' Find the last row in the Trigger file
    lastRowTrigger = triggerSheet.Cells(triggerSheet.Rows.Count, 1).End(xlUp).Row

    ' Headers to copy from Trigger.csv and paste into Portfolio
    headers = Array("Region", "Fund Manager", "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer")
    
    ' Create a dictionary to map Trigger file headers to their column numbers
    Set headerDict = CreateObject("Scripting.Dictionary")
    For i = 1 To triggerSheet.Cells(1, triggerSheet.Columns.Count).End(xlToLeft).Column
        headerDict(triggerSheet.Cells(1, i).Value) = i
    Next i
    
    ' Find the next empty row in the Portfolio sheet
    pasteRow = wsPortfolio.Cells(wsPortfolio.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Loop through each header and copy data from Trigger file to Portfolio
    For i = LBound(headers) To UBound(headers)
        If headerDict.exists(headers(i)) Then
            colTrigger = headerDict(headers(i)) ' Column number in Trigger.csv
            colPortfolio = Application.Match(headers(i), wsPortfolio.Rows(1), 0) ' Column number in Portfolio
            triggerSheet.Range(triggerSheet.Cells(2, colTrigger), triggerSheet.Cells(lastRowTrigger, colTrigger)).Copy _
                Destination:=wsPortfolio.Cells(pasteRow, colPortfolio)
        End If
    Next i

    ' Hygiene changes
    regionCol = Application.Match("Region", wsPortfolio.Rows(1), 0) ' Column number for Region
    triggerCol = Application.Match("Trigger/Non-Trigger", wsPortfolio.Rows(1), 0) ' Column number for Trigger/Non-Trigger

    With wsPortfolio
        Dim lastPasteRow As Long
        lastPasteRow = .Cells(.Rows.Count, 1).End(xlUp).Row ' Determine the last pasted row

        ' Replace 'US' with 'AMRS' and 'ASIA' with 'APAC' in the Region column
        For i = pasteRow To lastPasteRow
            Select Case .Cells(i, regionCol).Value
                Case "US"
                    .Cells(i, regionCol).Value = "AMRS"
                Case "ASIA"
                    .Cells(i, regionCol).Value = "APAC"
            End Select
        Next i

        ' Input 'Trigger' in the Trigger/Non-Trigger column for the pasted rows
        .Range(.Cells(pasteRow, triggerCol), .Cells(lastPasteRow, triggerCol)).Value = "Trigger"
    End With

    ' Close the Trigger.csv file
    wbTrigger.Close SaveChanges:=False

    MsgBox "Data from Trigger.csv has been successfully imported with hygiene changes.", vbInformation
End Sub