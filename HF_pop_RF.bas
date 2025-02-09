Option Explicit

' Constants for file paths - Update these with actual paths
Const HF_FILE_PATH As String = "C:\Path\To\HF_File.xlsx"
Const SHAREPOINT_FILE_PATH As String = "C:\Path\To\SharePoint_File.xlsx"

Sub Phase1_NewFundsIdentification()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Step 1 & 2: Import and convert files to tables
    ImportAndConvertToTables
    
    ' Step 3: Create required sheets if they don't exist and paste tables
    CreateRequiredSheets
    
    ' Step 4: Apply filters to HFTable
    ApplyHFTableFilters
    
    ' Step 5 & 6: Compare and create upload sheet
    CreateUploadSheet
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Phase 1 completed successfully!", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub ImportAndConvertToTables()
    Dim wbHF As Workbook, wbSP As Workbook
    
    ' Open source workbooks
    Set wbHF = Workbooks.Open(HF_FILE_PATH)
    Set wbSP = Workbooks.Open(SHAREPOINT_FILE_PATH)
    
    ' Convert HF data to table if not already
    If Not IsTable(wbHF.Sheets(1).UsedRange) Then
        wbHF.Sheets(1).UsedRange.Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "HFTable"
    End If
    
    ' Convert SharePoint data to table if not already
    If Not IsTable(wbSP.Sheets(1).UsedRange) Then
        wbSP.Sheets(1).UsedRange.Select
        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "SharePoint"
    End If
    
    ' Copy tables
    wbHF.Sheets(1).ListObjects("HFTable").Range.Copy
    wbSP.Sheets(1).ListObjects("SharePoint").Range.Copy
    
    ' Close source workbooks
    wbHF.Close False
    wbSP.Close False
End Sub

Private Sub CreateRequiredSheets()
    Dim ws As Worksheet
    
    ' Create Source Population sheet if it doesn't exist
    If Not SheetExists("Source Population") Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Source Population"
    End If
    
    ' Create SharePoint sheet if it doesn't exist
    If Not SheetExists("SharePoint") Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "SharePoint"
    End If
    
    ' Paste HFTable in Source Population sheet
    ThisWorkbook.Sheets("Source Population").Range("A1").PasteSpecial xlPasteAll
    
    ' Paste SharePoint table in SharePoint sheet
    ThisWorkbook.Sheets("SharePoint").Range("A1").PasteSpecial xlPasteAll
End Sub

Private Sub ApplyHFTableFilters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("Source Population")
    Set tbl = ws.ListObjects("HFTable")
    
    With tbl.Range
        ' Clear any existing filters
        If .AutoFilter Then .AutoFilter
        
        ' Apply all filters
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_Transparency_Tier"), Criteria1:=Array("1", "2"), Operator:=xlFilterValues
        
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Strategy"), _
            Criteria1:=Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"), _
            Operator:=xlFilterValues
        
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
            Criteria1:=Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                           "Managed Account", "Managed Account - No AF", "Loan Monitoring", _
                           "Loan FiF - No tracking", "Sleeve/share class/sub-account"), _
            Operator:=xlFilterValues
        
        ' Filter for dates 2023 and beyond
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_last_update_date"), _
            Criteria1:=">=" & DateSerial(2023, 1, 1)
    End With
End Sub

Private Sub CreateUploadSheet()
    Dim wsSource As Worksheet, wsSharePoint As Worksheet, wsUpload As Worksheet
    Dim tblSource As ListObject, tblSharePoint As ListObject
    Dim dictSharePoint As Object
    Dim cell As Range
    Dim row As Long
    
    ' Setup sheets
    Set wsSource = ThisWorkbook.Sheets("Source Population")
    Set wsSharePoint = ThisWorkbook.Sheets("SharePoint")
    
    ' Create Upload to SP sheet
    If Not SheetExists("Upload to SP") Then
        Set wsUpload = ThisWorkbook.Sheets.Add
        wsUpload.Name = "Upload to SP"
    Else
        Set wsUpload = ThisWorkbook.Sheets("Upload to SP")
        wsUpload.Cells.Clear
    End If
    
    ' Create headers for UploadHF table
    With wsUpload
        .Range("A1").Value = "HFAD_Fund_CoperID"
        .Range("B1").Value = "HFAD_Fund_Name"
        .Range("C1").Value = "HFAD_IM_CoperID"
        .Range("D1").Value = "HFAD_IM_Name"
        .Range("E1").Value = "HFAD_Credit_Officer"
        .Range("F1").Value = "Tier"
        .Range("G1").Value = "Status"
    End With
    
    ' Create dictionary of SharePoint CoperIDs
    Set dictSharePoint = CreateObject("Scripting.Dictionary")
    Set tblSharePoint = wsSharePoint.ListObjects("SharePoint")
    
    For Each cell In tblSharePoint.ListColumns("HFAD_Fund_CoperID").DataBodyRange
        If Not IsEmpty(cell) Then dictSharePoint(cell.Value) = True
    Next cell
    
    ' Populate upload table with new funds
    row = 2
    Set tblSource = wsSource.ListObjects("HFTable")
    
    For Each cell In tblSource.ListColumns("HFAD_Fund_CoperID").DataBodyRange
        If Not IsEmpty(cell) And Not dictSharePoint.Exists(cell.Value) Then
            With wsUpload
                .Cells(row, 1).Value = cell.Value  ' CoperID
                .Cells(row, 2).Value = GetValueFromSource(tblSource, cell.Row, "HFAD_Fund_Name")
                .Cells(row, 3).Value = GetValueFromSource(tblSource, cell.Row, "HFAD_IM_CoperID")
                .Cells(row, 4).Value = GetValueFromSource(tblSource, cell.Row, "HFAD_IM_Name")
                .Cells(row, 5).Value = GetValueFromSource(tblSource, cell.Row, "HFAD_Credit_Officer")
                .Cells(row, 6).Value = GetValueFromSource(tblSource, cell.Row, "IRR_Transparency_Tier")
                .Cells(row, 7).Value = "Active"
            End With
            row = row + 1
        End If
    Next cell
    
    ' Convert range to table
    If row > 2 Then
        wsUpload.Range("A1").CurrentRegion.Select
        wsUpload.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "UploadHF"
    End If
End Sub

' Helper Functions
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Private Function IsTable(rng As Range) As Boolean
    On Error Resume Next
    IsTable = rng.ListObject.Name <> ""
    On Error GoTo 0
End Function

Private Function GetColumnNumber(tbl As ListObject, columnName As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = columnName Then
            GetColumnNumber = col.Index
            Exit Function
        End If
    Next col
    GetColumnNumber = 0
End Function

Private Function GetValueFromSource(tbl As ListObject, rowNum As Long, columnName As String) As Variant
    On Error Resume Next
    GetValueFromSource = tbl.ListColumns(columnName).DataBodyRange.Cells(rowNum - tbl.HeaderRowRange.Row, 1).Value
    On Error GoTo 0
End Function
