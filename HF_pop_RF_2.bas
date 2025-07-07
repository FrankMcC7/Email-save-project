Option Explicit

Sub NewFundsIdentificationMacro()
    '=======================
    ' Main variable declarations
    Dim HFFilePath As String, SPFilePath As String
    Dim wbMain As Workbook, wbHF As Workbook, wbSP As Workbook
    Dim wsHFSource As Worksheet, wsSPSource As Worksheet
    Dim loHF As ListObject, loSP As ListObject
    Dim wsSourcePop As Worksheet, wsSPMain As Worksheet
    Dim loMainHF As ListObject, loMainSP As ListObject
    Dim visData As Range, r As Range
    Dim colIndex As Long
    Dim dictSP As Object
    Dim iSP As Long, key As String
    Dim newFunds As Collection, rec As Variant

    ' Paths and workbooks
    HFFilePath = "C:\YourFolder\HFFile.xlsx"
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"
    Set wbMain = ThisWorkbook

    ' Open HF file and table
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, wsHFSource.UsedRange, , xlYes)
    End If
    loHF.Name = "HFTable"

    ' Open SP file and table
    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, wsSPSource.UsedRange, , xlYes)
    End If
    loSP.Name = "SharePoint"

    ' Copy to main workbook
    On Error Resume Next
    Set wsSourcePop = wbMain.Sheets("Source Population")
    If wsSourcePop Is Nothing Then
        Set wsSourcePop = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSourcePop.Name = "Source Population"
    Else
        wsSourcePop.Cells.Clear
    End If
    Set wsSPMain = wbMain.Sheets("SharePoint")
    If wsSPMain Is Nothing Then
        Set wsSPMain = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSPMain.Name = "SharePoint"
    Else
        wsSPMain.Cells.Clear
    End If
    On Error GoTo 0
    loHF.Range.Copy wsSourcePop.Range("A1")
    loSP.Range.Copy wsSPMain.Range("A1")
    wbHF.Close False
    wbSP.Close False

    ' Ensure tables
    On Error Resume Next
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    If loMainHF Is Nothing Then Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes): loMainHF.Name = "HFTable"
    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    If loMainSP Is Nothing Then Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes): loMainSP.Name = "SharePoint"
    On Error GoTo 0

    ' Clear any existing filters
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData

    ' 1. Filter: only Transparency factor
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"

    ' 2. Filter: IRR_last_update_date from Jan 1, 2023 onwards
    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), Operator:=xlAnd

    ' 3. Filter: tier values 1 & 2 in IRR_Scorecard_factor_value
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues

    ' 4. Filter: Strategy include blank
    ApplyStrategyFilter ByVal loMainHF

    ' 5. Filter: Entity type include blank
    ApplyEntityFilter ByVal loMainHF

    ' Identify new funds not in SharePoint
    Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        Dim spData As Variant
        spData = loMainSP.DataBodyRange.Value
        For iSP = 1 To UBound(spData, 1)
            key = Trim(CStr(spData(iSP, colIndex)))
            If Not dictSP.Exists(key) Then dictSP.Add key, True
        Next iSP
    End If
    Set newFunds = New Collection
    colIndex = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        On Error Resume Next
        Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not visData Is Nothing Then
            Dim idxVal As Long, idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long
            idxVal = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
            idxName = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
            idxIMID = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
            idxIMName = GetColumnIndex(loMainHF, "HFAD_IM_Name")
            idxCred = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
            For Each r In visData.Rows
                If Not r.EntireRow.Hidden Then
                    Dim fundID As String
                    fundID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundID) Then
                        rec = Array(fundID, r.Cells(1, idxName).Value, r.Cells(1, idxIMID).Value, r.Cells(1, idxIMName).Value, r.Cells(1, idxCred).Value, r.Cells(1, idxVal).Value, "Active")
                        newFunds.Add rec
                    End If
                End If
            Next r
        End If
    End If

    ' Continue with upload logic...
    MsgBox "Macro completed successfully.", vbInformation
End Sub

' ======================= Helpers =======================

Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim(lo.HeaderRowRange.Cells(1, i).Value) = headerName Then
            GetColumnIndex = i: Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Sub ApplyStrategyFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Strategy")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Strategy", Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
    If UBound(allowed) >= LBound(allowed) Then loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Sub ApplyEntityFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Entity_type")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary", "Investment Manager as Agent", "Managed Account", "Managed Account - No AF", "Loan Monitoring", "Loan FiF - No tracking", "Sleeve/share class/sub-account"))
    If UBound(allowed) >= LBound(allowed) Then loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIdx As Long: colIdx = GetColumnIndex(lo, fieldName)
    If colIdx = 0 Then GetAllowedValues = Array(): Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, skipVal As Boolean, i As Long
    For Each cell In lo.ListColumns(fieldName).DataBodyRange
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If Trim(CStr(cell.Value)) = excludeArr(i) Then skipVal = True: Exit For
        Next i
        If Not skipVal Then dict(cell.Value) = cell.Value
    Next cell
    If dict.Count > 0 Then GetAllowedValues = dict.Keys Else GetAllowedValues = Array()
End Function

Function AppendToArray(arr As Variant, valueToAppend As Variant) As Variant
    Dim newArr() As Variant, n As Long, i As Long
    If Not IsArray(arr) Then
        newArr = Array(arr, valueToAppend)
    Else
        n = UBound(arr) - LBound(arr) + 1
        ReDim newArr(LBound(arr) To UBound(arr) + 1)
        For i = LBound(arr) To UBound(arr)
            newArr(i) = arr(i)
        Next i
        newArr(UBound(arr) + 1) = valueToAppend
    End If
    AppendToArray = newArr
End Function

Function ColumnExists(lo As ListObject, colName As String) As Boolean
    Dim cl As ListColumn
    For Each cl In lo.ListColumns
        If Trim(cl.Name) = colName Then ColumnExists = True: Exit Function
    Next cl
    ColumnExists = False
End Function
