Sub AttachOutlookEmailsFromSearch()
    Dim ws As Worksheet
    Dim lastRow As Long, rowNum As Long
    
    Dim outlookApp As Object  ' Outlook.Application (late bound)
    Dim outlookNs As Object   ' Outlook.Namespace
    Dim sourceMailItem As Object ' Outlook.MailItem, but late bound
    Dim newMailItem As Object    ' Outlook.MailItem, for the outgoing email
    
    Dim outlookLink As String
    Dim entryID As String
    Dim tempFilePath As String
    
    Dim recipientEmail As String
    
    '-----------------------------------------------------------------------
    ' 1. Identify the sheet and recipient
    '-----------------------------------------------------------------------
    Set ws = ThisWorkbook.Sheets("Search Email")
    
    ' For example, read the recipient from cell A1 (adjust as needed):
    recipientEmail = ws.Range("A1").Value
    If Len(recipientEmail) = 0 Then
        MsgBox "No recipient email found in A1. Please provide it.", vbExclamation
        Exit Sub
    End If
    
    ' Determine last row with data in column A (or whichever column is relevant)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "No search results found in 'Search Email'.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------
    ' 2. Connect to Outlook
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
    
    ' Get a reference to MAPI namespace
    Set outlookNs = outlookApp.GetNamespace("MAPI")
    
    '-----------------------------------------------------------------------
    ' 3. Create a brand new email to which we'll attach the found items
    '-----------------------------------------------------------------------
    Set newMailItem = outlookApp.CreateItem(0) ' 0 = olMailItem
    
    With newMailItem
        .To = recipientEmail
        .Subject = "Search Results: Outlook Emails"
        .Body = "Dear user," & vbNewLine & vbNewLine & _
                "Attached are the Outlook emails that matched your search criteria." & vbNewLine & vbNewLine & _
                "Best regards," & vbNewLine & "Your Company Name"
    End With
    
    '-----------------------------------------------------------------------
    ' 4. Loop through the rows in "Search Email" and process "outlook:" links
    '-----------------------------------------------------------------------
    For rowNum = 3 To lastRow
        ' If there's a hyperlink in column 4 (the Subject)
        If ws.Cells(rowNum, 4).Hyperlinks.Count > 0 Then
            
            outlookLink = ws.Cells(rowNum, 4).Hyperlinks(1).Address
            
            ' Check if it starts with "outlook:"
            If InStr(1, outlookLink, "outlook:", vbTextCompare) = 1 Then
                
                ' Extract everything *after* "outlook:"
                ' e.g. "outlook:00000000F503C..." => "00000000F503C..."
                entryID = Mid(outlookLink, Len("outlook:") + 1)
                
                On Error Resume Next
                Set sourceMailItem = outlookNs.GetItemFromID(entryID)
                On Error GoTo 0
                
                If Not sourceMailItem Is Nothing Then
                    '-------------------------------------------------------------------
                    ' 5. Save the mail item as a temporary .msg file, then attach
                    '-------------------------------------------------------------------
                    tempFilePath = Environ("TEMP") & "\TempEmail_" & rowNum & ".msg"
                    sourceMailItem.SaveAs tempFilePath, 3 ' 3 = olMSG
                    
                    If Dir(tempFilePath) <> "" Then
                        newMailItem.Attachments.Add tempFilePath
                    End If
                Else
                    Debug.Print "Could not retrieve Outlook item from EntryID: " & entryID
                End If
            Else
                Debug.Print "Not an 'outlook:' hyperlink: " & outlookLink
            End If
        End If
    Next rowNum
    
    '-----------------------------------------------------------------------
    ' 6. Display or send the new email
    '-----------------------------------------------------------------------
    newMailItem.Display  ' show to user
    ' newMailItem.Send   ' to send directly without preview
    
    '-----------------------------------------------------------------------
    ' 7. Clean up
    '-----------------------------------------------------------------------
    Set sourceMailItem = Nothing
    Set newMailItem = Nothing
    Set outlookNs = Nothing
    Set outlookApp = Nothing
    
    MsgBox "Your Outlook email has been prepared with the actual emails attached.", vbInformation
End Sub
