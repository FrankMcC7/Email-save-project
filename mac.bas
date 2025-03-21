Sub UpdateBetaFromSource()
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
        .Title = "Select the update file (must contain sheet 'Tracker' with a table)"
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
    
    ' Check if the update workbook contains a sheet named "Tracker".
    On Error Resume Next
    Set updateWs = updateWb.Sheets("Tracker")
    On Error GoTo 0
    If updateWs Is Nothing Then
        MsgBox "Sheet 'Tracker' not found in the update file."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Use the first table available in the "Tracker" sheet.
    If updateWs.ListObjects.Count = 0 Then
        MsgBox "No table found in sheet 'Tracker' in the update file."
        updateWb.Close SaveChanges:=False
        Exit Sub
    Else
        Set updateTable = updateWs.ListObjects(1)
    End If
    
    ' Verify required columns in the update table.
    Dim colFundGCI_Upd As Long, colECA As Long, colProspectus_Upd As Long, colStatus_Upd As Long
    Dim colFileName_Upd As Long, colOutreachDate_Upd As Long, colComments_Upd As Long
    On Error Resume Next
    colFundGCI_Upd = updateTable.ListColumns("Fund GCI").Index
    colECA = updateTable.ListColumns("ECA").Index
    colProspectus_Upd = updateTable.ListColumns("Prospectus").Index
    colStatus_Upd = updateTable.ListColumns("Status").Index
    colFileName_Upd = updateTable.ListColumns("File Name").Index
    colOutreachDate_Upd = updateTable.ListColumns("Outreach Date").Index
    colComments_Upd = updateTable.ListColumns("Comments").Index
    On Error GoTo 0
    If colFundGCI_Upd = 0 Or colECA = 0 Or colProspectus_Upd = 0 Or colStatus_Upd = 0 _
       Or colFileName_Upd = 0 Or colOutreachDate_Upd = 0 Or colComments_Upd = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'ECA', 'Prospectus', 'Status', 'File Name', 'Outreach Date', 'Comments') not found in the update table."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Verify required columns in main table "Beta".
    Dim colFundGCI_Beta As Long, colProspectus_Beta As Long, colStatus_Beta As Long
    Dim colFileName_Beta As Long, colOutreachDate_Beta As Long, colComments_Beta As Long
    On Error Resume Next
    colFundGCI_Beta = mainTable.ListColumns("Fund GCI").Index
    colProspectus_Beta = mainTable.ListColumns("Prospectus").Index
    colStatus_Beta = mainTable.ListColumns("Status").Index
    colFileName_Beta = mainTable.ListColumns("File Name").Index
    colOutreachDate_Beta = mainTable.ListColumns("Outreach Date").Index
    colComments_Beta = mainTable.ListColumns("Comments").Index
    On Error GoTo 0
    If colFundGCI_Beta = 0 Or colProspectus_Beta = 0 Or colStatus_Beta = 0 _
       Or colFileName_Beta = 0 Or colOutreachDate_Beta = 0 Or colComments_Beta = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'Prospectus', 'Status', 'File Name', 'Outreach Date', 'Comments') not found in the main table 'Beta'."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Prompt user to input the filter criteria for the ECA column.
    Dim filterCriteria As String, criteriaArray() As String
    filterCriteria = InputBox("Enter filter criteria for the ECA column. For multiple values, separate them by commas:", "ECA Filter Criteria")
    If Trim(filterCriteria) = "" Then
        MsgBox "No filter criteria provided. Exiting macro."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    criteriaArray = Split(filterCriteria, ",")
    
    Dim updateRow As ListRow
    Dim keyValue As Variant, ecaValue As String
    Dim prospectusValue As Variant, statusValue As Variant
    Dim fileNameValue As Variant, outreachDateValue As Variant, commentsValue As Variant
    Dim foundCell As Range
    Dim rowIndex As Long
    Dim iCriteria As Long, matchFound As Boolean
    
    ' Loop through each row in the update table.
    For Each updateRow In updateTable.ListRows
        ecaValue = CStr(updateRow.Range.Cells(1, colECA).Value)
        matchFound = False
        
        ' Check if the current ECA value matches any of the filter criteria.
        For iCriteria = LBound(criteriaArray) To UBound(criteriaArray)
            If Trim(criteriaArray(iCriteria)) = ecaValue Then
                matchFound = True
                Exit For
            End If
        Next iCriteria
        
        If matchFound Then
            keyValue = updateRow.Range.Cells(1, colFundGCI_Upd).Value
            prospectusValue = updateRow.Range.Cells(1, colProspectus_Upd).Value
            statusValue = updateRow.Range.Cells(1, colStatus_Upd).Value
            fileNameValue = updateRow.Range.Cells(1, colFileName_Upd).Value
            outreachDateValue = updateRow.Range.Cells(1, colOutreachDate_Upd).Value
            commentsValue = updateRow.Range.Cells(1, colComments_Upd).Value
            
            ' Find matching Fund GCI in the main table "Beta".
            Set foundCell = mainTable.DataBodyRange.Columns(colFundGCI_Beta).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole)
            If Not foundCell Is Nothing Then
                ' Calculate the relative row index within table "Beta".
                rowIndex = foundCell.Row - mainTable.DataBodyRange.Rows(1).Row + 1
                ' Update the Prospectus, Status, File Name, Outreach Date, and Comments columns in "Beta".
                mainTable.DataBodyRange.Cells(rowIndex, colProspectus_Beta).Value = prospectusValue
                mainTable.DataBodyRange.Cells(rowIndex, colStatus_Beta).Value = statusValue
                mainTable.DataBodyRange.Cells(rowIndex, colFileName_Beta).Value = fileNameValue
                mainTable.DataBodyRange.Cells(rowIndex, colOutreachDate_Beta).Value = outreachDateValue
                mainTable.DataBodyRange.Cells(rowIndex, colComments_Beta).Value = commentsValue
            End If
        End If
    Next updateRow
    
    ' Close the update workbook without saving changes.
    updateWb.Close SaveChanges:=False
    MsgBox "Update completed successfully!"
End Sub
