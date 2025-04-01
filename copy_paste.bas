Sub ImportCSVFilesToSheets()
    Dim triggerFilePath As String, nonTriggerFilePath As String
    Dim triggerWB As Workbook, nonTriggerWB As Workbook
    Dim currentWB As Workbook
    Dim wsAgreement As Worksheet, wsNotAgreement As Worksheet
    Dim fd As FileDialog
    Dim tbl As ListObject
    Dim lastRow As Long, lastCol As Long

    Set currentWB = ThisWorkbook

    ' Prompt user to select the Trigger CSV file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select Trigger CSV File"
    If fd.Show = -1 Then
        triggerFilePath = fd.SelectedItems(1)
    Else
        MsgBox "No trigger file selected. Exiting macro."
        Exit Sub
    End If

    ' Prompt user to select the Non-Trigger CSV file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select Non-Trigger CSV File"
    If fd.Show = -1 Then
        nonTriggerFilePath = fd.SelectedItems(1)
    Else
        MsgBox "No non-trigger file selected. Exiting macro."
        Exit Sub
    End If

    ' Check if "Agreement" sheet exists; if not, create it; if yes, clear its contents.
    On Error Resume Next
    Set wsAgreement = currentWB.Sheets("Agreement")
    On Error GoTo 0
    If wsAgreement Is Nothing Then
        Set wsAgreement = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.Count))
        wsAgreement.Name = "Agreement"
    Else
        wsAgreement.Cells.Clear
        ' Remove existing tables if any
        Dim lo As ListObject
        For Each lo In wsAgreement.ListObjects
            lo.Unlist
        Next lo
    End If

    ' Check if "Not_Agreement" sheet exists; if not, create it; if yes, clear its contents.
    On Error Resume Next
    Set wsNotAgreement = currentWB.Sheets("Not_Agreement")
    On Error GoTo 0
    If wsNotAgreement Is Nothing Then
        Set wsNotAgreement = currentWB.Sheets.Add(After:=currentWB.Sheets(currentWB.Sheets.Count))
        wsNotAgreement.Name = "Not_Agreement"
    Else
        wsNotAgreement.Cells.Clear
        For Each lo In wsNotAgreement.ListObjects
            lo.Unlist
        Next lo
    End If

    ' Open the Trigger CSV file and copy its data to the Agreement sheet
    Set triggerWB = Workbooks.Open(triggerFilePath)
    triggerWB.Sheets(1).UsedRange.Copy wsAgreement.Range("A1")
    triggerWB.Close SaveChanges:=False

    ' Open the Non-Trigger CSV file and copy its data to the Not_Agreement sheet
    Set nonTriggerWB = Workbooks.Open(nonTriggerFilePath)
    nonTriggerWB.Sheets(1).UsedRange.Copy wsNotAgreement.Range("A1")
    nonTriggerWB.Close SaveChanges:=False

    ' Convert data on the Agreement sheet to a table named "Agreement"
    With wsAgreement
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(lastRow, lastCol)), , xlYes)
        tbl.Name = "Agreement"
    End With

    ' Convert data on the Not_Agreement sheet to a table named "Not_Agreement"
    With wsNotAgreement
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(1, 1), .Cells(lastRow, lastCol)), , xlYes)
        tbl.Name = "Not_Agreement"
    End With

    MsgBox "Data imported and tables created successfully!"
End Sub