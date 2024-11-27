Sub SearchEmails()
    Dim ws As Worksheet
    Dim searchValue As String
    Dim foundRange As Range
    Dim firstAddress As String
    Dim searchColumn As String
    Dim searchCriteria As String
    Dim resultsSheet As Worksheet
    Dim resultRow As Long

    ' Set the worksheet where the email log is stored
    Set ws = ThisWorkbook.Sheets("Email Logs")
    
    ' Prompt user to select the column to search
    searchCriteria = InputBox("Enter the search criteria: (1 for Date, 2 for Sender Email, 3 for Subject)", "Search Emails")
    If searchCriteria = "" Then Exit Sub
    
    ' Map the input to a column letter
    Select Case searchCriteria
        Case "1": searchColumn = "A" ' Date
        Case "2": searchColumn = "B" ' Sender Email
        Case "3": searchColumn = "C" ' Subject
        Case Else
            MsgBox "Invalid input. Please enter 1, 2, or 3.", vbExclamation
            Exit Sub
    End Select

    ' Prompt the user for the search term
    searchValue = InputBox("Enter the value to search for:", "Search Emails")
    If searchValue = "" Then Exit Sub

    ' Clear previous highlights and filters
    Call ClearHighlights

    ' Create or activate the "Search Results" sheet
    On Error Resume Next
    Set resultsSheet = ThisWorkbook.Sheets("Search Results")
    If resultsSheet Is Nothing Then
        Set resultsSheet = ThisWorkbook.Sheets.Add
        resultsSheet.Name = "Search Results"
    End If
    On Error GoTo 0
    resultsSheet.Cells.Clear
    resultsSheet.Range("A1:C1").Value = Array("Date", "Sender Email", "Subject")
    resultRow = 2

    ' Search for the value in the selected column
    Set foundRange = ws.Columns(searchColumn).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

    If Not foundRange Is Nothing Then
        firstAddress = foundRange.Address
        Do
            ' Highlight the matching row
            foundRange.EntireRow.Interior.Color = RGB(255, 255, 0) ' Yellow highlight

            ' Copy matching row to "Search Results" sheet
            resultsSheet.Cells(resultRow, 1).Value = ws.Cells(foundRange.Row, 1).Value ' Date
            resultsSheet.Cells(resultRow, 2).Value = ws.Cells(foundRange.Row, 2).Value ' Sender Email
            resultsSheet.Cells(resultRow, 3).Value = ws.Cells(foundRange.Row, 3).Value ' Subject
            resultRow = resultRow + 1

            Set foundRange = ws.Columns(searchColumn).FindNext(foundRange)
        Loop While Not foundRange Is Nothing And foundRange.Address <> firstAddress

        ' Autofit columns in the "Search Results" sheet
        resultsSheet.Columns("A:C").AutoFit

        MsgBox "Search complete. Matching rows have been highlighted and copied to the 'Search Results' sheet.", vbInformation
    Else
        MsgBox "No matching records found.", vbExclamation
    End If
End Sub

Sub ClearHighlights()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Email Logs")
    ws.Cells.Interior.ColorIndex = xlNone
    ws.AutoFilterMode = False
    MsgBox "All highlights and filters have been cleared.", vbInformation
End Sub