Sub UpdateFrequencyFromSource()
    Dim sourceFilePath As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim sourceTbl As ListObject
    Dim destTbl As ListObject
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim i As Long
    
    ' Let the user select the source file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select the Source File"
    If fd.Show <> -1 Then
        MsgBox "No file selected. Exiting macro."
        Exit Sub
    End If
    sourceFilePath = fd.SelectedItems(1)
    
    ' Open the source workbook
    Set sourceWb = Workbooks.Open(sourceFilePath)
    
    ' Get Sheet2 from the source file
    On Error Resume Next
    Set sourceWs = sourceWb.Worksheets("Sheet2")
    On Error GoTo 0
    If sourceWs Is Nothing Then
        MsgBox "Sheet2 not found in the source file."
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Get the table named "Source" from Sheet2
    On Error Resume Next
    Set sourceTbl = sourceWs.ListObjects("Source")
    On Error GoTo 0
    If sourceTbl Is Nothing Then
        MsgBox "Table 'Source' not found in Sheet2 of the source file."
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Find the destination table "newTable" in the current workbook.
    Set destTbl = Nothing
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set destTbl = ws.ListObjects("newTable")
        On Error GoTo 0
        If Not destTbl Is Nothing Then Exit For
    Next ws
    If destTbl Is Nothing Then
        MsgBox "Destination table 'newTable' not found in the current workbook."
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Determine column positions in the source table based on header names.
    Dim srcFundGCI_Col As Long, srcPeriod_Col As Long, srcTrigger_Col As Long
    srcFundGCI_Col = 0: srcPeriod_Col = 0: srcTrigger_Col = 0
    For i = 1 To sourceTbl.HeaderRowRange.Columns.Count
        Select Case LCase(Trim(sourceTbl.HeaderRowRange.Cells(1, i).Value))
            Case "fund gci"
                srcFundGCI_Col = i
            Case "period"
                srcPeriod_Col = i
            Case "trigger value"
                srcTrigger_Col = i
        End Select
    Next i
    If srcFundGCI_Col = 0 Or srcPeriod_Col = 0 Or srcTrigger_Col = 0 Then
        MsgBox "One or more required columns (Fund GCI, Period, Trigger Value) not found in the source table."
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Determine column positions in the destination table ("newTable")
    Dim destFundGCI_Col As Long, destFrequency_Col As Long
    destFundGCI_Col = 0: destFrequency_Col = 0
    For i = 1 To destTbl.HeaderRowRange.Columns.Count
        Select Case LCase(Trim(destTbl.HeaderRowRange.Cells(1, i).Value))
            Case "fund gci"
                destFundGCI_Col = i
            Case "frequency"
                destFrequency_Col = i
        End Select
    Next i
    If destFundGCI_Col = 0 Or destFrequency_Col = 0 Then
        MsgBox "Destination table does not contain required columns (Fund GCI, Frequency)."
        sourceWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Loop through each row in the destination table and update the Frequency column.
    Dim destFundGCI As String
    Dim candidatePeriod As Variant
    Dim periodToUse As Variant
    Dim foundNonBlank As Boolean
    Dim srcRowIndex As Long
    Dim srcRowCount As Long
    Dim triggerVal As String
    
    Dim destData As Range
    Set destData = destTbl.DataBodyRange
    
    Dim destRowIndex As Long
    For destRowIndex = 1 To destData.Rows.Count
        destFundGCI = Trim(destData.Cells(destRowIndex, destFundGCI_Col).Value)
        If destFundGCI <> "" Then
            candidatePeriod = ""
            periodToUse = ""
            foundNonBlank = False
            
            srcRowCount = sourceTbl.DataBodyRange.Rows.Count
            For srcRowIndex = 1 To srcRowCount
                If Trim(sourceTbl.DataBodyRange.Cells(srcRowIndex, srcFundGCI_Col).Value) = destFundGCI Then
                    ' Capture the first period value for this GCI (if not already captured)
                    If candidatePeriod = "" Then
                        candidatePeriod = sourceTbl.DataBodyRange.Cells(srcRowIndex, srcPeriod_Col).Value
                    End If
                    triggerVal = Trim(sourceTbl.DataBodyRange.Cells(srcRowIndex, srcTrigger_Col).Value)
                    ' If trigger value is not blank, then use this period and exit loop.
                    If triggerVal <> "" Then
                        periodToUse = sourceTbl.DataBodyRange.Cells(srcRowIndex, srcPeriod_Col).Value
                        foundNonBlank = True
                        Exit For
                    End If
                End If
            Next srcRowIndex
            
            ' If no nonblank trigger was found, then use the candidate period (first matching row)
            If periodToUse = "" Then periodToUse = candidatePeriod
            
            ' Update the Frequency column in the destination table.
            destData.Cells(destRowIndex, destFrequency_Col).Value = periodToUse
        End If
    Next destRowIndex
    
    ' Close the source workbook without saving changes.
    sourceWb.Close SaveChanges:=False
    
    MsgBox "Frequency values updated successfully."
End Sub