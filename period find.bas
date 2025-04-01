Sub UpdateFrequencyFromSource_Faster()
    Dim sourceFilePath As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim sourceTbl As ListObject
    Dim destTbl As ListObject
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim i As Long
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Prompt user to select the source file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select the Source File"
    If fd.Show <> -1 Then
        MsgBox "No file selected. Exiting macro."
        GoTo CleanUp
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
        GoTo CleanUp
    End If
    
    ' Get the table named "Source" from Sheet2
    On Error Resume Next
    Set sourceTbl = sourceWs.ListObjects("Source")
    On Error GoTo 0
    If sourceTbl Is Nothing Then
        MsgBox "Table 'Source' not found in Sheet2 of the source file."
        sourceWb.Close SaveChanges:=False
        GoTo CleanUp
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
        GoTo CleanUp
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
        GoTo CleanUp
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
        GoTo CleanUp
    End If
    
    ' Read the source table data into an array
    Dim srcData As Variant
    Dim numSrcRows As Long
    srcData = sourceTbl.DataBodyRange.Value
    numSrcRows = UBound(srcData, 1)
    
    ' Build a dictionary for quick lookup:
    ' Key: Fund GCI, Value: Array(candidatePeriod, nonBlankPeriod)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim fundKey As String, candidatePeriod As Variant, nonBlankPeriod As Variant, triggerVal As String
    For i = 1 To numSrcRows
        fundKey = Trim(srcData(i, srcFundGCI_Col))
        If fundKey <> "" Then
            If Not dict.exists(fundKey) Then
                candidatePeriod = srcData(i, srcPeriod_Col)
                triggerVal = Trim(srcData(i, srcTrigger_Col))
                If triggerVal <> "" Then
                    nonBlankPeriod = srcData(i, srcPeriod_Col)
                Else
                    nonBlankPeriod = ""
                End If
                dict.Add fundKey, Array(candidatePeriod, nonBlankPeriod)
            Else
                ' If nonBlankPeriod is still empty and current row has non-blank trigger, update it.
                triggerVal = Trim(srcData(i, srcTrigger_Col))
                If dict(fundKey)(1) = "" And triggerVal <> "" Then
                    dict(fundKey)(1) = srcData(i, srcPeriod_Col)
                End If
            End If
        End If
    Next i
    
    ' Process the destination table using an array for speed.
    Dim destData As Variant
    Dim numDestRows As Long
    destData = destTbl.DataBodyRange.Value
    numDestRows = UBound(destData, 1)
    
    Dim currentFund As String, periodToUse As Variant
    For i = 1 To numDestRows
        currentFund = Trim(destData(i, destFundGCI_Col))
        If currentFund <> "" Then
            If dict.exists(currentFund) Then
                ' Use non-blank period if available; otherwise, use candidate period.
                If dict(currentFund)(1) <> "" Then
                    periodToUse = dict(currentFund)(1)
                Else
                    periodToUse = dict(currentFund)(0)
                End If
                destData(i, destFrequency_Col) = periodToUse
            Else
                destData(i, destFrequency_Col) = "" ' No match found
            End If
        End If
    Next i
    
    ' Write the updated data back to the destination table
    destTbl.DataBodyRange.Value = destData
    
    ' Close the source workbook without saving changes.
    sourceWb.Close SaveChanges:=False
    
    MsgBox "Frequency values updated successfully.", vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub