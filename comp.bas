Sub ProcessCSVData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim i As Integer
    
    'Delete first row
    Rows(1).Delete
    
    'Convert to table with headers
    Range("A1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = "DataTable"
    
    'Remove blank rows
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    'Filter for Approved status
    ActiveSheet.ListObjects("DataTable").Range.AutoFilter Field:=Application.WorksheetFunction.Match("Review Status", Range("1:1"), 0), Criteria1:="Approved"
    
    'Copy visible cells to new sheet
    Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    Sheets.Add.Name = "ApprovedData"
    ActiveSheet.Paste
    Range("A1").CurrentRegion.RemoveSubtotal
    
    'Random sampling loop
    Do
        lastRow = Sheets("ApprovedData").Cells(Rows.Count, "A").End(xlUp).Row
        
        'Create 5 sample sheets
        For i = 1 To 5
            Sheets.Add.Name = "Sample" & i
            
            'Get 100 random rows
            With Sheets("ApprovedData")
                .Range("A1").EntireRow.Copy Sheets("Sample" & i).Range("A1")
                
                'Sample 100 random rows
                Dim usedRows As Collection
                Set usedRows = New Collection
                Dim j As Long, k As Long
                
                For j = 1 To 100
                    Do
                        k = Int((lastRow - 1) * Rnd + 2)
                        On Error Resume Next
                        usedRows.Add k, CStr(k)
                        If Err.Number = 0 Then Exit Do
                        Err.Clear
                    Loop
                    .Rows(k).Copy Sheets("Sample" & i).Rows(j + 1)
                Next j
            End With
        Next i
        
        'Wait 5 seconds
        Application.Wait Now + TimeValue("00:00:05")
        
        'Delete sample sheets
        For i = 1 To 5
            Application.DisplayAlerts = False
            Sheets("Sample" & i).Delete
            Application.DisplayAlerts = True
        Next i
        
        DoEvents
    Loop Until Application.UserCancelKey
    
End Sub
