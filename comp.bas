Sub ProcessCSVData()
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
    
    Do
        For i = 1 To 5
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Sample" & i
            Sheets("ApprovedData").Range("A1").EntireRow.Copy Sheets("Sample" & i).Range("A1")
            
            'Random sample using Advanced Filter
            With Sheets("ApprovedData")
                .Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
                    CopyToRange:=Sheets("Sample" & i).Range("A2"), _
                    Unique:=True, Random:=100
            End With
        Next i
        
        Application.Wait Now + TimeValue("00:00:05")
        
        For i = 1 To 5
            Application.DisplayAlerts = False
            Sheets("Sample" & i).Delete
            Application.DisplayAlerts = True
        Next i
        
        DoEvents
    Loop
End Sub
