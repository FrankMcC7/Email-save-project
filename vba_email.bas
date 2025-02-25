Public Sub EmailSearchResults()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim outlookApp As Object      ' Late-bound Outlook.Application
    Dim outlookMail As Object     ' Late-bound Outlook.MailItem
    
    Dim attachPath As String
    Dim userName As String
    Dim defaultEmail As String
    Dim recipientEmails As String
    
    Dim debugMsg As String
    
    '----------------------------------------------------
    ' 1) Reference the "Search Email" sheet
    '----------------------------------------------------
    Set ws = ThisWorkbook.Sheets("Search Email")
    
    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' If there are no results below row 2, exit
    If lastRow < 3 Then
        MsgBox "No search results found in 'Search Email'. Please run your search first.", vbInformation
        Exit Sub
    End If
    
    '----------------------------------------------------
    ' 2) Determine the user's email address
    '    (a) Build a default using the Windows username + domain
    '    (b) Prompt user to confirm or override
    '----------------------------------------------------
    userName = Environ("USERNAME")         ' e.g. "john.doe"
    defaultEmail = userName & "@mycompany.com"  ' Adjust domain as needed
    
    recipientEmails = InputBox( _
        Prompt:="Enter or confirm recipient email(s):", _
        Title:="Email Search Results", _
        Default:=defaultEmail _
    )
    
    ' If user cancels or clears the box, stop
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient email provided. Process aborted.", vbExclamation
        Exit Sub
    End If
    
    '----------------------------------------------------
    ' 3) Initialize Outlook
    '----------------------------------------------------
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
    
    '----------------------------------------------------
    ' 4) Create a new mail item
    '----------------------------------------------------
    Set outlookMail = outlookApp.CreateItem(0)  ' 0 = olMailItem
    
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria." & vbNewLine & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best Regards," & vbNewLine & "Your Company Name"
        
        ' We'll track debugging info here:
        debugMsg = "Attempting to attach files from 'Search Email' (rows 3 through " & lastRow & "):" & vbNewLine
        
        '------------------------------------------------
        ' 5) Loop through each row of search results
        '    and attach the file if it exists
        '------------------------------------------------
        For r = 3 To lastRow
            ' Does Column 4 (D) have a hyperlink?
            If ws.Cells(r, 4).Hyperlinks.Count > 0 Then
                attachPath = ws.Cells(r, 4).Hyperlinks(1).Address
                
                ' Normalize path if needed
                attachPath = Replace(attachPath, "%20", " ") ' decode spaces
                attachPath = Replace(attachPath, "/", "\")   ' ensure backslashes if any
                
                debugMsg = debugMsg & vbNewLine & "Row " & r & ": " & attachPath
                
                ' Check if file actually exists before attaching
                If Len(Dir(attachPath)) > 0 Then
                    .Attachments.Add attachPath
                Else
                    debugMsg = debugMsg & "  -> NOT FOUND!"
                End If
            Else
                debugMsg = debugMsg & vbNewLine & "Row " & r & ": No hyperlink in column 4"
            End If
        Next r
        
        '------------------------------------------------
        ' 6) Display the email for user review
        '    (or .Send to send immediately)
        '------------------------------------------------
        .Display
    End With
    
    '----------------------------------------------------
    ' 7) Debugging output
    '----------------------------------------------------
    Debug.Print debugMsg  ' Open VBA Immediate Window (Ctrl+G) to see details
    
    ' Cleanup
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Your email has been created. Open Outlook to review/send. " & vbCrLf & _
           "Check the VBA Immediate Window (Ctrl+G) for any missing files.", vbInformation
End Sub
