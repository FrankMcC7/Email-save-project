Public Sub EmailSearchResults_RobustFilePathFix()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    
    Dim outlookApp As Object
    Dim outlookMail As Object
    
    Dim rawHyperlink As String
    Dim finalPath As String
    Dim debugMsg As String
    
    Dim userName As String
    Dim defaultEmail As String
    Dim recipientEmails As String
    
    '------------------------------
    '1) Reference the "Search Email" sheet
    '------------------------------
    Set ws = ThisWorkbook.Sheets("Search Email")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "No search results found in 'Search Email'. Please run your search first.", vbInformation
        Exit Sub
    End If
    
    '------------------------------
    '2) Ask for email address (or build from Windows username)
    '------------------------------
    userName = Environ("USERNAME") ' e.g. "jsmith"
    defaultEmail = userName & "@mycompany.com"
    recipientEmails = InputBox("Enter or confirm recipient email(s):", _
                               "Email Search Results", defaultEmail)
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient email provided. Process aborted.", vbExclamation
        Exit Sub
    End If
    
    '------------------------------
    '3) Initialize Outlook
    '------------------------------
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
    
    '------------------------------
    '4) Create a new mail
    '------------------------------
    Set outlookMail = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With outlookMail
        .To = recipientEmails
        .Subject = "Search Results: Emails from Excel"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the .msg files that matched your search criteria." & vbNewLine & _
                "Please review them as needed." & vbNewLine & vbNewLine & _
                "Best Regards," & vbNewLine & "Your Company Name"
        
        debugMsg = "File attachments (rows 3 to " & lastRow & "):" & vbNewLine
        
        '------------------------------
        '5) Loop rows: fix hyperlink path
        '------------------------------
        For r = 3 To lastRow
            If ws.Cells(r, 4).Hyperlinks.Count > 0 Then
                
                ' (A) Get the raw hyperlink
                rawHyperlink = ws.Cells(r, 4).Hyperlinks(1).Address
                
                ' Log the original:
                debugMsg = debugMsg & vbNewLine & _
                           "Row " & r & " - Original: " & rawHyperlink
                
                ' (B) Find the UNC portion (the first occurrence of "\\" )
                Dim pos As Long
                pos = InStr(1, rawHyperlink, "\\") ' find first double backslash
                
                If pos > 0 Then
                    ' Keep everything from the first "\\" onward
                    finalPath = Mid(rawHyperlink, pos)
                Else
                    ' No "\\" found; fallback to original
                    finalPath = rawHyperlink
                End If
                
                ' (C) Decode any %20 -> space
                finalPath = Replace(finalPath, "%20", " ")
                
                ' (D) Convert forward slashes to backslashes
                finalPath = Replace(finalPath, "/", "\")
                
                ' Log the final path we will test
                debugMsg = debugMsg & vbNewLine & _
                           "        Final: " & finalPath
                
                ' (E) Check if file exists
                If Len(Dir(finalPath)) > 0 Then
                    .Attachments.Add finalPath
                Else
                    debugMsg = debugMsg & "  -> NOT FOUND!"
                End If
                
            Else
                debugMsg = debugMsg & vbNewLine & _
                           "Row " & r & " - No hyperlink in col D"
            End If
        Next r
        
        '------------------------------
        '6) Show the email (or .Send to auto-send)
        '------------------------------
        .Display
    End With
    
    '------------------------------
    '7) Print debug info
    '------------------------------
    Debug.Print debugMsg
    
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Email created. Check attachments and the Immediate Window (Ctrl+G) for details.", vbInformation
End Sub
