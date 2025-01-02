Option Explicit

Sub ProcessCSVData()
    Dim lastRow As Long
    Dim i As Long
    Dim statusCol As Long
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    'Delete first row
    ActiveSheet.Rows(1).Delete
    
    'Remove blank rows
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountA(ActiveSheet.Rows(i)) = 0 Then
            ActiveSheet.Rows(i).Delete
        End If
    Next i
    
    'Find Review Status column
    statusCol = findColumnNumber("Review Status")
    If statusCol = 0 Then
        MsgBox "Review Status column not found!"
        Exit Sub
    End If
    
    'Filter for Approved
    With ActiveSheet
        .Range("A1").CurrentRegion.AutoFilter Field:=statusCol, Criteria1:="Approved"
        If Not IsEmpty(.Range("A2")) Then
            .Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
            Sheets.Add.Name = "ApprovedData"
            ActiveSheet.Range("A1").PasteSpecial xlPasteValues
        End If
    End With
    
    Do While True
        createSampleSheets
        Application.Wait Now + TimeValue("00:00:05")
        deleteSampleSheets
        DoEvents
    Loop
    
ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Function findColumnNumber(headerName As String) As Long
    Dim foundCell As Range
    Set foundCell = ActiveSheet.Range("1:1").Find(headerName, LookIn:=xlValues)
    If Not foundCell Is Nothing Then
        findColumnNumber = foundCell.Column
    Else
        findColumnNumber = 0
    End If
End Function

Sub createSampleSheets()
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim randRow As Long
    
    lastRow = Sheets("ApprovedData").Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To 5
        Sheets.Add.Name = "Sample" & i
        Sheets("ApprovedData").Rows(1).Copy Sheets("Sample" & i).Rows(1)
        
        For j = 2 To 101
            randRow = WorksheetFunction.RandBetween(2, lastRow)
            Sheets("ApprovedData").Rows(randRow).Copy Sheets("Sample" & i).Rows(j)
        Next j
    Next i
End Sub

Sub deleteSampleSheets()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Sample" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
End Sub
