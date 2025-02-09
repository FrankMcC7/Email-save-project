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
    Dim wsHF As Worksheet, wsSP As Worksheet
    Dim rngHF As Range, rngSP As Range
    
    ' Open source workbooks
    Set wbHF = Workbooks.Open(HF_FILE_PATH)
    Set wbSP = Workbooks.Open(SHAREPOINT_FILE_PATH)
    
    ' Get first worksheet from each workbook
    Set wsHF = wbHF.Sheets(1)
    Set wsSP = wbSP.Sheets(1)
    
    ' Set ranges
    Set rngHF = wsHF.UsedRange
    Set rngSP = wsSP.UsedRange
    
    ' Convert HF data to table if not already
    If Not IsTable(rngHF) Then
        wsHF.ListObjects.Add(xlSrcRange, rngHF, , xlYes).Name = "HFTable"
    ElseIf wsHF.ListObjects(1).Name <> "HFTable" Then
        wsHF.ListObjects(1).Name = "HFTable"
    End If
    
    ' Convert SharePoint data to table if not already
    If Not IsTable(rngSP) Then
        wsSP.ListObjects.Add(xlSrcRange, rngSP, , xlYes).Name = "SharePoint"
    ElseIf wsSP.ListObjects(1).Name <> "SharePoint" Then
        wsSP.ListObjects(1).Name = "SharePoint"
    End If
    
    ' Copy tables to clipboard
    wsHF.ListObjects("HFTable").Range.Copy
    
    ' Create Source Population sheet and paste
    If Not SheetExists("Source Population") Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "Source Population"
    End If
    ThisWorkbook.Sheets("Source Population").Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False  ' Clear clipboard
    
    ' Copy SharePoint table
    wsSP.ListObjects("SharePoint").Range.Copy
    
    ' Create SharePoint sheet and paste
    If Not SheetExists("SharePoint") Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "SharePoint"
    End If
    ThisWorkbook.Sheets("SharePoint").Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False  ' Clear clipboard
    
    ' Close source workbooks
    wbHF.Close False
    wbSP.Close False
End Sub

Private Sub CreateRequiredSheets()
    ' Sheets are now created in ImportAndConvertToTables
    ' This sub is kept for potential future use
End Sub

Private Sub ApplyHFTableFilters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("Source Population")
    On Error Resume Next
    Set tbl = ws.ListObjects("HFTable")
    If tbl Is Nothing Then Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    With tbl.Range
        ' Clear any existing filters
        If .AutoFilter Then .AutoFilter
        
        ' Apply all filters
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_Transparency_Tier"), Criteria1:=Array("1", "2"), Operator:=xlFilterValues
        
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Strategy"), _
            Criteria1:="<>FIF", Operator:=xlAnd, _
            Criteria2:="<>Fund of Funds", _
            Criteria2:="<>Sub/Sleeve- No Benchmark"
        
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
            Criteria1:="<>Guaranteed subsidiary", Operator:=xlAnd, _
            Criteria2:="<>Investment Manager as Agent"
        
        ' Apply second set of filters for HFAD_Entity_type since we have more exclusions
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
            Criteria1:="<>Managed Account", Operator:=xlAnd, _
            Criteria2:="<>Managed Account - No AF"
            
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
            Criteria1:="<>Loan Monitoring", Operator:=xlAnd, _
            Criteria2:="<>Loan FiF - No tracking"
            
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
            Criteria1:="<>Sleeve/share class/sub-account"
        
        ' Filter for dates 2023 and beyond
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_last_update_date"), _
            Criteria1:=">=" & DateSerial(2023, 1, 1)
    End With
End Sub

Private Sub CreateUploadSheet()
    Dim wsSource As Worksheet, wsSharePoint As Worksheet, wsUpload As Worksheet
    Dim tblSource As ListObject, tblSharePoint As ListObject, tblUpload As ListObject
    Dim dictSharePoint As Object
    Dim cell As Range
    Dim row As Long
    Dim rngUpload As Range
    
    ' Setup sheets
    Set wsSource = ThisWorkbook.Sheets("Source Population")
    Set wsSharePoint = ThisWorkbook.Sheets("SharePoint")
    
    ' Create Upload to SP sheet
    If Not SheetExists("Upload to SP") Then
        Set wsUpload = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
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
    Set tblSharePoint = wsSharePoint.ListObjects(1)
    
    For Each cell In tblSharePoint.ListColumns("HFAD_Fund_CoperID").DataBodyRange
        If Not IsEmpty(cell) Then dictSharePoint(cell.Value) = True
    Next cell
    
    ' Populate upload table with new funds
    row = 2
    Set tblSource = wsSource.ListObjects(1)
    
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
    
    ' Convert range to table if we have data
    If row > 2 Then
        Set rngUpload = wsUpload.Range("A1").Resize(row - 1, 7)
        Set tblUpload = wsUpload.ListObjects.Add(xlSrcRange, rngUpload, , xlYes)
        tblUpload.Name = "UploadHF"
    End If
End Sub

' Helper Functions
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
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
