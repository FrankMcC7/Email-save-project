Sub EmailSearchResults_WithFileNormalization()
    Dim wsSearch As Worksheet
    Dim lastRow As Long, rowNum As Long
    
    Dim outlookApp As Object  ' Late bound Outlook.Application
    Dim outlookMail As Object ' Late bound Outlook.MailItem
    
    Dim rawPath As String
    Dim recipientEmails As String
    
    '-----------------------------------------------------------------------
    ' 1. WHERE DOES THE RECIPIENT COME FROM?
    '    (A) Hard-code a cell in "Search Email" with the user’s email address
    '    (B) OR prompt with InputBox
    '-----------------------------------------------------------------------
    
    Set wsSearch = ThisWorkbook.Sheets("Search Email")
    
    ' Example (A) – read from a cell in "Search Email", e.g., A1
    recipientEmails = wsSearch.Range("A1").Value
    
    ' If instead you want an InputBox, comment the above line and uncomment:
    ' recipientEmails = InputBox("Enter recipient email(s):", "Email Search Results")
    
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient specified. Please supply an email address in 'Search Email'! (or InputBox).", vbExclamation
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------
    ' 2. DETERMINE HOW MANY ROWS OF SEARCH RESULTS
    '-----------------------------------------------------------------------
    lastRow = wsSearch.Cells(wsSearch.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found. Please run the search first.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------
    ' 3. ACCESS OUTLOOK
    '-----------------------------------------------------------------------
    On Error Resume Next
    Set outlookApp = GetObject(Class:="Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If outlookApp Is Nothing Then
        MsgBox "Could not start or find Outlook. Please ensure Outlook is installed.", vbCritical
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------
    ' 4. CREATE A NEW EMAIL
    '-----------------------------------------------------------------------
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria. " & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best regards," & vbNewLine & _
                "Your Company Name"
        
        '-------------------------------------------------------------------
        ' 5. LOOP THROUGH SEARCH RESULTS TO ATTACH MSG FILES
        '    (Assuming Column 4 has a hyperlink to the .msg file)
        '-------------------------------------------------------------------
        For rowNum = 3 To lastRow
            ' Make sure there is actually a hyperlink in column 4 (D)
            If wsSearch.Cells(rowNum, 4).Hyperlinks.Count > 0 Then
                rawPath = wsSearch.Cells(rowNum, 4).Hyperlinks(1).Address
                
                '--- CLEAN UP THE LINK IF IT HAS file:/// or %20, etc. ---
                
                ' If it starts with "file:///" remove that portion
                If InStr(1, rawPath, "file:///", vbTextCompare) = 1 Then
                    rawPath = Replace(rawPath, "file:///", "")
                ElseIf InStr(1, rawPath, "file://", vbTextCompare) = 1 Then
                    rawPath = Replace(rawPath, "file://", "")
                End If
                
                ' Replace forward slashes with backslashes
                rawPath = Replace(rawPath, "/", "\")
                
                ' Decode %20 -> space (if any)
                rawPath = Replace(rawPath, "%20", " ")
                
                ' Now check if the file actually exists
                If Dir(rawPath) <> "" Then
                    .Attachments.Add rawPath
                Else
                    ' Optional: For debugging or logging
                    Debug.Print "Could not find file: " & rawPath
                End If
                
            End If
        Next rowNum
        
        '-------------------------------------------------------------------
        ' 6. DISPLAY OR SEND THE EMAIL
        '-------------------------------------------------------------------
        .Display  ' show the email so the user can review/edit before sending
        ' .Send   ' or send immediately without displaying
    End With
    
    '-----------------------------------------------------------------------
    ' 7. CLEAN UP
    '-----------------------------------------------------------------------
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "An Outlook email has been created with the matching .msg attachments, if found.", vbInformation
End Sub
