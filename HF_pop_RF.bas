Option Explicit

' Constants for file paths - Update these with actual paths
Const HF_FILE_PATH As String = "C:\Path\To\HF_File.xlsx"
Const SHAREPOINT_FILE_PATH As String = "C:\Path\To\SharePoint_File.xlsx"

Sub Phase1_NewFundsIdentification()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ImportAndProcessTables
    ApplyHFTableFilters
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

Private Sub ImportAndProcessTables()
    Dim wbHF As Workbook, wbSP As Workbook
    Dim wsHF As Worksheet, wsSP As Worksheet
    Dim wsTarget As Worksheet
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
    
    ' Process HF Table
    If Not IsTable(rngHF) Then
        wsHF.ListObjects.Add(xlSrcRange, rngHF, , xlYes).Name = "HFTable"
    ElseIf wsHF.ListObjects(1).Name <> "HFTable" Then
        wsHF.ListObjects(1).Name = "HFTable"
    End If
    
    ' Process SharePoint Table
    If Not IsTable(rngSP) Then
        wsSP.ListObjects.Add(xlSrcRange, rngSP, , xlYes).Name = "SharePoint"
    ElseIf wsSP.ListObjects(1).Name <> "SharePoint" Then
        wsSP.ListObjects(1).Name = "SharePoint"
    End If
    
    ' Create and populate Source Population sheet
    If Not SheetExists("Source Population") Then
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTarget.Name = "Source Population"
    Else
        Set wsTarget = ThisWorkbook.Sheets("Source Population")
    End If
    
    wsHF.ListObjects("HFTable").Range.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Create and populate SharePoint sheet
    If Not SheetExists("SharePoint") Then
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTarget.Name = "SharePoint"
    Else
        Set wsTarget = ThisWorkbook.Sheets("SharePoint")
    End If
    
    wsSP.ListObjects("SharePoint").Range.Copy
    wsTarget.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Close source workbooks
    wbHF.Close False
    wbSP.Close False
End Sub

Private Sub ApplyHFTableFilters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim strategyExclusions As Variant
    Dim entityExclusions As Variant
    
    Set ws = ThisWorkbook.Sheets("Source Population")
    On Error Resume Next
    Set tbl = ws.ListObjects("HFTable")
    If tbl Is Nothing Then Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    ' Define exclusion arrays
    strategyExclusions = Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark")
    entityExclusions = Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                           "Managed Account", "Managed Account - No AF", "Loan Monitoring", _
                           "Loan FiF - No tracking", "Sleeve/share class/sub-account")
    
    With tbl.Range
        ' Clear any existing filters
        If .AutoFilter Then .AutoFilter
        
        ' Apply all filters
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_Transparency_Tier"), _
                   Criteria1:=Array("1", "2"), _
                   Operator:=xlFilterValues
        
        ' Filter out strategies using NOT IN logic
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Strategy"), _
                   Criteria1:=strategyExclusions, _
                   Operator:=xlFilterValues, _
                   Criteria2:="="
        
        ' Filter out entity types using NOT IN logic
        .AutoFilter Field:=GetColumnNumber(tbl, "HFAD_Entity_type"), _
                   Criteria1:=entityExclusions, _
                   Operator:=xlFilterValues, _
                   Criteria2:="="
        
        ' Filter for dates 2023 and beyond
        .AutoFilter Field:=GetColumnNumber(tbl, "IRR_last_update_date"), _
                   Criteria1:=">=" & DateSerial(2023, 1, 1)
    End With
End Sub

Private Sub CreateUploadSheet()
    Dim wsSource As Worksheet, wsSharePoint As Worksheet, wsUpload As Worksheet
    Dim tblSource As ListObject, tblSharePoint As ListObject, tblUpload As ListObject
    Dim dictSharePoint As Object
    Dim visibleRange As Range
    Dim dataRow As Range
    Dim row As Long
    Dim rngUpload As Range
    
    ' Setup sheets
    Set wsSource = ThisWorkbook.Sheets("Source Population")
    Set wsSharePoint = ThisWorkbook.Sheets("SharePoint")
    Set tblSource = wsSource.ListObjects(1)
    
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
    
    For Each dataRow In tblSharePoint.ListColumns("HFAD_Fund_CoperID").DataBodyRange
        If Not IsEmpty(dataRow) Then dictSharePoint(dataRow.Value) = True
    Next dataRow
    
    ' Get visible rows after filtering
    On Error Resume Next
    Set visibleRange = tblSource.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If visibleRange Is Nothing Then
        MsgBox "No visible rows found after filtering.", vbInformation
        Exit Sub
    End If
    
    ' Populate upload table with new funds
    row = 2
    
    ' Loop through visible rows only
    For Each dataRow In visibleRange.Rows
        Dim coperID As Variant
        coperID = dataRow.Cells(GetColumnNumber(tblSource, "HFAD_Fund_CoperID"), 1).Value
        
        If Not IsEmpty(coperID) And Not dictSharePoint.Exists(coperID) Then
            With wsUpload
                .Cells(row, 1).Value = coperID
                .Cells(row, 2).Value = dataRow.Cells(GetColumnNumber(tblSource, "HFAD_Fund_Name"), 1).Value
                .Cells(row, 3).Value = dataRow.Cells(GetColumnNumber(tblSource, "HFAD_IM_CoperID"), 1).Value
                .Cells(row, 4).Value = dataRow.Cells(GetColumnNumber(tblSource, "HFAD_IM_Name"), 1).Value
                .Cells(row, 5).Value = dataRow.Cells(GetColumnNumber(tblSource, "HFAD_Credit_Officer"), 1).Value
                .Cells(row, 6).Value = dataRow.Cells(GetColumnNumber(tblSource, "IRR_Transparency_Tier"), 1).Value
                .Cells(row, 7).Value = "Active"
            End With
            row = row + 1
        End If
    Next dataRow
    
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
