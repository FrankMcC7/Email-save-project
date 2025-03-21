Sub UpdateBetaFromAlpha()
    Dim mainWs As Worksheet, updateWs As Worksheet
    Dim mainTable As ListObject, updateTable As ListObject
    Dim updateWb As Workbook
    Dim fd As FileDialog
    Dim filePath As String
    
    ' Set reference to the main worksheet and table "Beta" in the current workbook.
    Set mainWs = ThisWorkbook.Sheets("Tracker")
    Set mainTable = mainWs.ListObjects("Beta")
    
    ' Prompt the user to select the update file.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the update file (must contain sheet 'Tracker' with table 'Alpha')"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Exiting macro."
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    
    ' Open the update workbook.
    Set updateWb = Workbooks.Open(filePath)
    
    ' Set reference to the update worksheet and table "Alpha".
    On Error Resume Next
    Set updateWs = updateWb.Sheets("Tracker")
    Set updateTable = updateWs.ListObjects("Alpha")
    On Error GoTo 0
    
    If updateTable Is Nothing Then
        MsgBox "Table 'Alpha' not found in sheet 'Tracker' of the update file."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Get column indexes for required columns in table Alpha.
    Dim colFundGCI_Alpha As Long, colECA As Long, colProspectus_Alpha As Long, colStatus_Alpha As Long
    On Error Resume Next
    colFundGCI_Alpha = updateTable.ListColumns("Fund GCI").Index
    colECA = updateTable.ListColumns("ECA").Index
    colProspectus_Alpha = updateTable.ListColumns("Prospectus").Index
    colStatus_Alpha = updateTable.ListColumns("Status").Index
    On Error GoTo 0
    If colFundGCI_Alpha = 0 Or colECA = 0 Or colProspectus_Alpha = 0 Or colStatus_Alpha = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'ECA', 'Prospectus', 'Status') not found in table 'Alpha'."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Get column indexes for required columns in table Beta.
    Dim colFundGCI_Beta As Long, colProspectus_Beta As Long, colStatus_Beta As Long
    On Error Resume Next
    colFundGCI_Beta = mainTable.ListColumns("Fund GCI").Index
    colProspectus_Beta = mainTable.ListColumns("Prospectus").Index
    colStatus_Beta = mainTable.ListColumns("Status").Index
    On Error GoTo 0
    If colFundGCI_Beta = 0 Or colProspectus_Beta = 0 Or colStatus_Beta = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'Prospectus', 'Status') not found in table 'Beta'."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim updateRow As ListRow
    Dim keyValue As Variant, ecaValue As String
    Dim prospectusValue As Variant, statusValue As Variant
    Dim foundCell As Range
    Dim rowIndex As Long
    
    ' Loop through each row in table "Alpha".
    For Each updateRow In updateTable.ListRows
        ecaValue = CStr(updateRow.Range.Cells(1, colECA).Value)
        ' Only update rows where ECA is "Amit" or "Revanth".
        If ecaValue = "Amit" Or ecaValue = "Revanth" Then
            keyValue = updateRow.Range.Cells(1, colFundGCI_Alpha).Value
            prospectusValue = updateRow.Range.Cells(1, colProspectus_Alpha).Value
            statusValue = updateRow.Range.Cells(1, colStatus_Alpha).Value
            
            ' Find matching Fund GCI in table "Beta".
            Set foundCell = mainTable.DataBodyRange.Columns(colFundGCI_Beta).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                ' Calculate the relative row index within table "Beta".
                rowIndex = foundCell.Row - mainTable.DataBodyRange.Rows(1).Row + 1
                ' Update the Prospectus and Status columns in table "Beta".
                mainTable.DataBodyRange.Cells(rowIndex, colProspectus_Beta).Value = prospectusValue
                mainTable.DataBodyRange.Cells(rowIndex, colStatus_Beta).Value = statusValue
            End If
        End If
    Next updateRow
    
    ' Close the update workbook without saving any changes.
    updateWb.Close SaveChanges:=False
    MsgBox "Update completed successfully!"
End Sub
