Public Sub EmailSearchResults_FixFilePrefix()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    
    Dim outlookApp As Object
    Dim outlookMail As Object
    
    Dim attachPath As String
    Dim debugMsg As String
    
    Dim userName As String
    Dim defaultEmail As String
    Dim recipientEmails As String
    
    ' 1) Reference the "Search Email" sheet
    Set ws = ThisWorkbook.Sheets("Search Email")
    
    ' 2) Check rows
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found in 'Search Email'. Please run your search first.", vbInformation
        Exit Sub
    End If
    
    ' 3) Ask user for an email address (or build from their Windows login)
    userName = Environ("USERNAME") ' e.g. "jsmith"
    defaultEmail = userName & "@mycompany.com"
    recipientEmails = InputBox( _
        "Enter or confirm recipient email(s):", _
        "Email Search Results", _
        defaultEmail)
    
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient email provided. Process aborted.", vbExclamation
        Exit Sub
    End If
    
    ' 4) Get Outlook
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
    
    ' 5) Create a new mail
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria." & vbNewLine & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best Regards," & vbNewLine & "Your Company Name"
        
        debugMsg = "Attaching files from rows 3 through " & lastRow & ":" & vbNewLine
        
        ' 6) Loop through each row in "Search Email"
        For r = 3 To lastRow
            If ws.Cells(r, 4).Hyperlinks.Count > 0 Then
                attachPath = ws.Cells(r, 4).Hyperlinks(1).Address
                
                ' --- Clean up the path ---
                
                ' Remove "file:///" or "file://"
                If InStr(1, attachPath, "file:///", vbTextCompare) = 1 Then
                    attachPath = Replace(attachPath, "file:///", "")
                ElseIf InStr(1, attachPath, "file://", vbTextCompare) = 1 Then
                    attachPath = Replace(attachPath, "file://", "")
                End If
                
                ' Convert any %20 to spaces
                attachPath = Replace(attachPath, "%20", " ")
                
                ' Convert any forward slashes to backslashes
                attachPath = Replace(attachPath, "/", "\")
                
                ' Log it in debug
                debugMsg = debugMsg & vbNewLine & "Row " & r & ": " & attachPath
                
                ' Check if file exists
                If Len(Dir(attachPath)) > 0 Then
                    .Attachments.Add attachPath
                Else
                    debugMsg = debugMsg & "  -> NOT FOUND!"
                End If
            Else
                debugMsg = debugMsg & vbNewLine & "Row " & r & ": No hyperlink in col D"
            End If
        Next r
        
        ' Show the email
        .Display
    End With
    
    ' 7) Print debug info to Immediate Window (Ctrl+G)
    Debug.Print debugMsg
    
    ' Cleanup
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Email created. Check attachments. " & _
           "See Immediate Window (Ctrl+G) for any missing paths.", vbInformation
End Sub
