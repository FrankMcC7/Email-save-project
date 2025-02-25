Sub EmailSearchResults_AttachUNCFiles_AutoUserEmail()
    Dim wsSearch As Worksheet
    Dim lastRow As Long, rowNum As Long
    
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim attachPath As String
    Dim recipientEmails As String
    Dim debugMsg As String
    Dim userEmail As String
    
    ' Set the worksheet
    Set wsSearch = ThisWorkbook.Sheets("Search Email")
    
    '----------------------------------------------------------------------
    ' 1. Get Active User's Email OR Ask for It
    '----------------------------------------------------------------------
    
    ' Attempt to fetch the active user's system email (Excel UserName or Windows Username)
    On Error Resume Next
    userEmail = Application.UserName  ' Sometimes contains full name instead of email
    If InStr(1, userEmail, "@") = 0 Then
        ' If no email is found in Application.UserName, try Windows environment variable
        userEmail = Environ("USERNAME") & "@yourcompany.com" ' Adjust domain if needed
    End If
    On Error GoTo 0
    
    ' Ask user if email isn't detected automatically
    recipientEmails = userEmail
    If Len(recipientEmails) = 0 Or InStr(1, recipientEmails, "@") = 0 Then
        recipientEmails = InputBox("Enter your email address (or multiple separated by commas):", "Email Search")
    End If
    
    ' If still empty, exit
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient email provided. Please enter an email address.", vbExclamation
        Exit Sub
    End If
    
    '----------------------------------------------------------------------
    ' 2. Identify last row with search results
    '----------------------------------------------------------------------
    lastRow = wsSearch.Cells(wsSearch.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found. Please run the search first.", vbInformation
        Exit Sub
    End If
    
    '----------------------------------------------------------------------
    ' 3. Initialize Outlook
    '----------------------------------------------------------------------
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If outlookApp Is Nothing Then
        MsgBox "Could not start Outlook. Please ensure Outlook is installed.", vbCritical
        Exit Sub
    End If
    
    '----------------------------------------------------------------------
    ' 4. Create a new email
    '----------------------------------------------------------------------
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria. " & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best regards," & vbNewLine & "Your Company Name"
        
        debugMsg = "Checking files:" & vbNewLine
        
        '----------------------------------------------------------------------
        ' 5. Loop through search results and attach .msg files
        '----------------------------------------------------------------------
        For rowNum = 3 To lastRow
            ' Check if there's a hyperlink in Column 4 (Subject)
            If wsSearch.Cells(rowNum, 4).Hyperlinks.Count > 0 Then
                attachPath = wsSearch.Cells(rowNum, 4).Hyperlinks(1).Address
                
                ' Ensure UNC path is properly formatted
                attachPath = Replace(attachPath, "%20", " ") ' Convert spaces
                attachPath = Replace(attachPath, "/", "\") ' Ensure backslashes
                
                ' Debug print to check what paths are being read
                debugMsg = debugMsg & vbNewLine & attachPath
                
                ' Ensure file exists before attaching
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                Else
                    debugMsg = debugMsg & " - NOT FOUND!"
                End If
            Else
                debugMsg = debugMsg & vbNewLine & "No hyperlink found in row " & rowNum
            End If
        Next rowNum
        
        '----------------------------------------------------------------------
        ' 6. Display or send the email
        '----------------------------------------------------------------------
        .Display ' Show to user before sending
        ' .Send ' Uncomment to send automatically
        
    End With
    
    '----------------------------------------------------------------------
    ' 7. Debug Output
    '----------------------------------------------------------------------
    Debug.Print debugMsg ' Open Immediate Window (Ctrl+G) to check
    
    ' Cleanup
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Your email has been created. Check the Immediate Window (Ctrl+G) for missing files.", vbInformation
End Sub
