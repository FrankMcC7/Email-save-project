Private Sub CommandButton1_Click()
    Dim searchDate As String
    Dim searchSenderName As String
    Dim searchSenderEmail As String
    Dim searchSubject As String
    Dim sourceSheet As Worksheet
    Dim searchSheet As Worksheet
    Dim rowNum As Long
    Dim resultRow As Long
    Dim matchFound As Boolean

    ' Get input values from the UserForm
    searchDate = Me.TextBox1.Value
    searchSenderName = Me.TextBox2.Value
    searchSenderEmail = Me.TextBox3.Value
    searchSubject = Me.TextBox4.Value

    ' Set source and search sheets
    Set sourceSheet = ThisWorkbook.Sheets("Email Logs")
    Set searchSheet = ThisWorkbook.Sheets("Search Email")

    ' Clear previous search results
    searchSheet.Cells.Clear
    searchSheet.Cells(1, 1).Value = "Search Results"
    searchSheet.Cells(2, 1).Value = "Date"
    searchSheet.Cells(2, 2).Value = "Sender Name"
    searchSheet.Cells(2, 3).Value = "Sender Email"
    searchSheet.Cells(2, 4).Value = "Subject"
    searchSheet.Rows(2).Font.Bold = True
    searchSheet.Rows(2).HorizontalAlignment = xlCenter

    ' Initialize result row counter (starts below the header)
    resultRow = 3
    matchFound = False

    ' Loop through rows in the source sheet
    For rowNum = 3 To sourceSheet.UsedRange.Rows.Count
        Dim matchRow As Boolean
        matchRow = True

        ' Check each column for matches based on user input
        If searchDate <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 1).Value, searchDate, vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchSenderName <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 2).Value, searchSenderName, vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchSenderEmail <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 3).Value, searchSenderEmail, vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchSubject <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 4).Value, searchSubject, vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If

        ' If the row matches all criteria, copy it to the results sheet
        If matchRow Then
            searchSheet.Cells(resultRow, 1).Value = sourceSheet.Cells(rowNum, 1).Value ' Date
            searchSheet.Cells(resultRow, 2).Value = sourceSheet.Cells(rowNum, 2).Value ' Sender Name
            searchSheet.Cells(resultRow, 3).Value = sourceSheet.Cells(rowNum, 3).Value ' Sender Email

            ' Add hyperlink to the Subject column
            With searchSheet.Cells(resultRow, 4)
                .Value = sourceSheet.Cells(rowNum, 4).Value ' Subject
                .Hyperlinks.Add Anchor:=searchSheet.Cells(resultRow, 4), _
                                Address:=sourceSheet.Cells(rowNum, 4).Hyperlinks(1).Address, _
                                TextToDisplay:=sourceSheet.Cells(rowNum, 4).Value
                .Font.Color = RGB(0, 0, 255)
                .Font.Underline = xlUnderlineStyleSingle
            End With

            resultRow = resultRow + 1
            matchFound = True
        End If
    Next rowNum

    ' Apply formatting to results sheet
    With searchSheet
        .Columns("A:D").AutoFit
        .Rows("2:" & resultRow - 1).RowHeight = 15
        .Cells.HorizontalAlignment = xlLeft
        .Cells.VerticalAlignment = xlCenter
    End With

    ' Notify user of the result
    If matchFound Then
        MsgBox "Search completed. Results are displayed in the 'Search Email' tab.", vbInformation
    Else
        MsgBox "No matching records found.", vbExclamation
    End If

    ' Close the UserForm
    Unload Me
End Sub