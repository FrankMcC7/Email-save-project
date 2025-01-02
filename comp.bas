Sub ProcessCSVData()
    On Error Resume Next
    
    'Check if ApprovedData sheet exists and delete
    Application.DisplayAlerts = False
    If SheetExists("ApprovedData") Then
        Sheets("ApprovedData").Delete
    End If
    Application.DisplayAlerts = True
    
    'Delete first row
    ActiveSheet.Rows(1).Delete
    
    'Remove blank rows
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
        End If
    Next i
    
    'Find Review Status column
    Dim statusCol As Long
    statusCol = FindColumn("Review Status")
    
    If statusCol = 0 Then
        MsgBox "Review Status column not found!", vbCritical
        Exit Sub
    End If
    
    'Filter and copy approved data
    ActiveSheet.Range("A1").CurrentRegion.AutoFilter Field:=statusCol, Criteria1:="Approved"
    
    If Not IsEmpty(ActiveSheet.Range("A2")) Then
        ActiveSheet.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
        Sheets.Add.Name = "ApprovedData"
        ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    Else
        MsgBox "No approved data found!", vbCritical
        Exit Sub
    End If
    
    'Random sampling loop
    Do
        For i = 1 To 5
            CreateRandomSample i
            DoEvents
        Next i
        
        Application.Wait Now + TimeValue("00:00:05")
        
        DeleteSampleSheets
        DoEvents
    Loop

End Sub

Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    
    SheetExists = False
End Function

Private Function FindColumn(headerName As String) As Long
    Dim cell As Range
    
    On Error Resume Next
    Set cell = ActiveSheet.Range("1:1").Find(headerName, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not cell Is Nothing Then
        FindColumn = cell.Column
    Else
        FindColumn = 0
    End If
End Function

Private Sub CreateRandomSample(sheetNum As Integer)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    'Create new sample sheet
    Application.DisplayAlerts = False
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Sample" & sheetNum
    Application.DisplayAlerts = True
    
    'Copy headers
    Sheets("ApprovedData").Range("A1").EntireRow.Copy ws.Range("A1")
    
    'Get random rows
    lastRow = Sheets("ApprovedData").Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow > 101 Then  'Only if we have more than 100 rows + header
        Dim randArr() As Long
        ReDim randArr(1 To 100)
        
        'Generate random numbers
        Dim j As Long
        For j = 1 To 100
            randArr(j) = Application.WorksheetFunction.RandBetween(2, lastRow)
        Next j
        
        'Copy random rows
        For j = 1 To 100
            Sheets("ApprovedData").Rows(randArr(j)).Copy ws.Rows(j + 1)
        Next j
    End If
End Sub

Private Sub DeleteSampleSheets()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Sample" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub
