Sub EnhancedSearchForm()
    Dim searchForm As Object
    Dim searchDate As String
    Dim searchSenderName As String
    Dim searchSenderEmail As String
    Dim searchSubject As String
    Dim sourceSheet As Worksheet
    Dim searchSheet As Worksheet
    Dim searchRange As Range
    Dim rowNum As Long
    Dim resultRow As Long
    Dim matchFound As Boolean

    ' Set source and search sheets
    Set sourceSheet = ThisWorkbook.Sheets("Email Logs")
    Set searchSheet = ThisWorkbook.Sheets("Search Email")

    ' Create search form
    Set searchForm = CreateObject("Scripting.Dictionary")
    searchDate = InputBox("Enter Date to search (leave blank to skip):", "Search Criteria")
    searchSenderName = InputBox("Enter Sender Name to search (leave blank to skip):", "Search Criteria")
    searchSenderEmail = InputBox("Enter Sender Email to search (leave blank to skip):", "Search Criteria")
    searchSubject = InputBox("Enter Subject to search (leave blank to skip):", "Search Criteria")

    ' Add criteria to the searchForm dictionary
    searchForm.Add "Date", searchDate
    searchForm.Add "Sender Name", searchSenderName
    searchForm.Add "Sender Email", searchSenderEmail
    searchForm.Add "Subject", searchSubject

    ' If all criteria are blank, exit
    If Len(searchDate & searchSenderName & searchSenderEmail & searchSubject) = 0 Then
        MsgBox "No search criteria entered. Exiting search.", vbExclamation
        Exit Sub
    End If

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
        If searchForm("Date") <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 1).Value, searchForm("Date"), vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchForm("Sender Name") <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 2).Value, searchForm("Sender Name"), vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchForm("Sender Email") <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 3).Value, searchForm("Sender Email"), vbTextCompare) = 0 Then
                matchRow = False
            End If
        End If
        If searchForm("Subject") <> "" Then
            If InStr(1, sourceSheet.Cells(rowNum, 4).Value, searchForm("Subject"), vbTextCompare) = 0 Then
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
End Sub