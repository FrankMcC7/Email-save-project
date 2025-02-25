Sub EmailSearchResults_AttachUNCFiles()
    Dim wsSearch As Worksheet
    Dim lastRow As Long, rowNum As Long
    
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim attachPath As String
    Dim recipientEmails As String
    
    ' Set the worksheet
    Set wsSearch = ThisWorkbook.Sheets("Search Email")
    
    ' Get the recipient email from cell A1 (or prompt user)
    recipientEmails = wsSearch.Range("A1").Value
    If Len(recipientEmails) = 0 Then
        MsgBox "No recipient email found in A1. Please provide it.", vbExclamation
        Exit Sub
    End If
    
    ' Determine the last row of data
    lastRow = wsSearch.Cells(wsSearch.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found. Please run the search first.", vbInformation
        Exit Sub
    End If
    
    ' Initialize Outlook
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
    
    ' Create a new email
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
                
                ' UNC paths should be directly usable but ensure they exist
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                Else
                    Debug.Print "Could not find file: " & attachPath
                End If
            End If
        Next rowNum
        
        ' Display the email for user review before sending
        .Display ' or use .Send to send immediately
    End With
    
    ' Cleanup
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Your email has been created with the attached emails.", vbInformation
End Sub
