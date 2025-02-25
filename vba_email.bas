Sub EmailSearchResults()
    Dim wsSearch As Worksheet
    Dim lastRow As Long, rowNum As Long
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim attachPath As String
    Dim recipientEmails As String
    
    ' Prompt user for recipient email(s). Could also come from a UserForm textbox.
    recipientEmails = InputBox("Enter the recipient email addresses (comma-separated if multiple):", _
                               "Email Search Results")
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient provided. Macro aborted.", vbExclamation
        Exit Sub
    End If
    
    ' Reference the "Search Email" sheet
    Set wsSearch = ThisWorkbook.Sheets("Search Email")
    
    ' Determine the last row of data
    lastRow = wsSearch.Cells(wsSearch.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found. Please run the search first.", vbInformation
        Exit Sub
    End If
    
    ' Try to get an Outlook Application object
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If outlookApp Is Nothing Then
        MsgBox "Outlook could not be accessed. Please make sure Outlook is installed.", vbCritical
        Exit Sub
    End If
    
    ' Create the mail item
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria. " & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best regards," & vbNewLine & "Your Company Name"
        
        ' Loop through all matching rows and attach .msg files
        For rowNum = 3 To lastRow
            ' Check if there's a hyperlink in Column 4 (Subject)
            If wsSearch.Cells(rowNum, 4).Hyperlinks.Count > 0 Then
                attachPath = wsSearch.Cells(rowNum, 4).Hyperlinks(1).Address
                
                ' (Optional) Validate attachPath is a .msg file or actual file path
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                Else
                    Debug.Print "Could not find file: " & attachPath
                End If
            End If
        Next rowNum
        
        ' Display the email for the user to review before sending:
        .Display  ' or use .Send to send directly
    End With
    
    ' Cleanup
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    MsgBox "Your email has been created (or sent).", vbInformation
End Sub
