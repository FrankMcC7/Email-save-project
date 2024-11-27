Sub SearchEmails()
    Dim searchValue As String
    Dim sourceSheet As Worksheet
    Dim resultsSheet As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim resultRow As Long

    ' Prompt for search value
    searchValue = InputBox("Enter the search term (Date, Sender Name, Sender Email, or Subject):", "Search Emails")
    If searchValue = "" Then Exit Sub

    ' Set source and results sheets
    Set sourceSheet = ThisWorkbook.Sheets("Email Logs")
    On Error Resume Next
    Set resultsSheet = ThisWorkbook.Sheets("Search Results")
    On Error GoTo 0

    ' Create the "Search Results" sheet if it doesn't exist
    If resultsSheet Is Nothing Then
        Set resultsSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultsSheet.Name = "Search Results"
        resultsSheet.Cells(1, 1).Value = "Date"
        resultsSheet.Cells(1, 2).Value = "Sender Name"
        resultsSheet.Cells(1, 3).Value = "Sender Email"
        resultsSheet.Cells(1, 4).Value = "Subject"
        resultsSheet.Cells(1, 4).Font.Underline = xlUnderlineStyleSingle
        resultsSheet.Cells(1, 4).Font.Color = RGB(0, 0, 255)
    Else
        resultsSheet.Cells.Clear
        resultsSheet.Cells(1, 1).Value = "Date"
        resultsSheet.Cells(1, 2).Value = "Sender Name"
        resultsSheet.Cells(1, 3).Value = "Sender Email"
        resultsSheet.Cells(1, 4).Value = "Subject"
        resultsSheet.Cells(1, 4).Font.Underline = xlUnderlineStyleSingle
        resultsSheet.Cells(1, 4).Font.Color = RGB(0, 0, 255)
    End If

    ' Set the header font for the results sheet
    resultsSheet.Rows(1).Font.Bold = True

    ' Define the range to search in the source sheet
    Set searchRange = sourceSheet.UsedRange

    ' Initialize result row counter
    resultRow = 2

    ' Loop through each cell in the source sheet and search for the value
    For Each cell In searchRange
        If InStr(1, cell.Value, searchValue, vbTextCompare) > 0 Then
            ' Copy the entire row to the results sheet
            resultsSheet.Cells(resultRow, 1).Value = sourceSheet.Cells(cell.Row, 1).Value ' Date
            resultsSheet.Cells(resultRow, 2).Value = sourceSheet.Cells(cell.Row, 2).Value ' Sender Name
            resultsSheet.Cells(resultRow, 3).Value = sourceSheet.Cells(cell.Row, 3).Value ' Sender Email

            ' Add hyperlink to the Subject column
            With resultsSheet.Cells(resultRow, 4)
                .Value = sourceSheet.Cells(cell.Row, 4).Value ' Subject
                .Hyperlinks.Add Anchor:=resultsSheet.Cells(resultRow, 4), _
                                Address:=sourceSheet.Cells(cell.Row, 4).Hyperlinks(1).Address, _
                                TextToDisplay:=sourceSheet.Cells(cell.Row, 4).Value
                .Font.Color = RGB(0, 0, 255)
                .Font.Underline = xlUnderlineStyleSingle
            End With

            resultRow = resultRow + 1
        End If
    Next cell

    ' Notify user of completion
    MsgBox "Search completed. Results are available in the 'Search Results' tab.", vbInformation
End Sub