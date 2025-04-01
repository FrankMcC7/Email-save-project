Sub BuildDashboard()
    Dim fDialog As FileDialog
    Dim sourceFilePath As String
    Dim wbSource As Workbook
    Dim wsData As Worksheet
    Dim wsDashboard As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    ' 1) Prompt the user to select the source data file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select the Source Data File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        
        If .Show <> -1 Then
            MsgBox "No file selected. Macro will end.", vbExclamation
            Exit Sub
        End If
        
        sourceFilePath = .SelectedItems(1)
    End With
    
    ' 2) Open the source workbook and copy data to this workbook
    Application.ScreenUpdating = False
    Set wbSource = Workbooks.Open(sourceFilePath)
    
    'Assume data is on the first worksheet in the source file:
    wbSource.Worksheets(1).Activate
    
    ' Find the last row and column of data in the source sheet
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Copy the data range
    Range(Cells(1, 1), Cells(lastRow, lastCol)).Copy
    
    ' Paste into the current workbook on a sheet named "RawData" (create if not exists)
    ThisWorkbook.Activate
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("RawData")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "RawData"
    Else
        ' Clear old data if needed
        wsData.Cells.Clear
    End If
    
    wsData.Range("A1").PasteSpecial xlPasteAll
    
    ' Close the source workbook
    wbSource.Close SaveChanges:=False
    
    ' 3) Create a pivot cache from the data
    '    (Optional) You could turn "RawData" into a Table and reference that.
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
              SourceType:=xlDatabase, _
              SourceData:=wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol)))
    
    ' 4) Create a new sheet called "Dashboard" for our pivot
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Dashboard").Delete ' Delete if exists
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsDashboard = ThisWorkbook.Worksheets.Add
    wsDashboard.Name = "Dashboard"
    
    ' 5) Insert the Pivot Table onto the "Dashboard" sheet
    Set pt = wsDashboard.PivotTables.Add(PivotCache:=pc, TableDestination:="Dashboard!A3", TableName:="Pivot_Main")
    
    ' 6) Add/Arrange the Pivot Fields (update field names to match your actual columns)
    With pt
        ' Clear any default fields first
        .PivotFields("SomeRowField").Orientation = xlHidden
        ' ...
        
        ' Example arrangement (Replace with your real field names)
        .PivotFields("NAV Date").Orientation = xlRowField
        .PivotFields("NAV Date").Position = 1
        
        .PivotFields("Reporting Date").Orientation = xlRowField
        .PivotFields("Reporting Date").Position = 2
        
        .PivotFields("Month").Orientation = xlColumnField
        
        ' Example measure fields in the Values area
        .AddDataField .PivotFields("Total # of Funds"), "Sum of Total # of Funds", xlSum
        .AddDataField .PivotFields("Total NAV"), "Sum of Total NAV", xlSum
        
        ' (Optional) Another measure
        '.AddDataField .PivotFields("Something"), "Sum of Something", xlSum
        
        ' (Optional) If you have a field to use as a slicer or filter
        '.PivotFields("Fund Manager").Orientation = xlPageField
        
        ' Format the Pivot Table or set Subtotals, etc.
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
    End With
    
    ' (Optional) You can create a Pivot Chart from the same pivot cache or from the Pivot Table
    ' Example of creating a Pivot Chart
    'Dim chtObj As ChartObject
    'Set chtObj = wsDashboard.ChartObjects.Add(Left:=400, Top:=20, Width:=400, Height:=300)
    'chtObj.Chart.SetSourceData Source:=Range("Dashboard!A3")
    'chtObj.Chart.ChartType = xlColumnClustered
    
    ' You can further customize the Pivot Table style, row/column labels, and more here
    Application.ScreenUpdating = True
    MsgBox "Dashboard created successfully!", vbInformation
End Sub