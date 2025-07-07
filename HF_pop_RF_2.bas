Option Explicit

Sub NewFundsIdentificationMacro()
    '=======================
    ' Main variable declarations
    Dim HFFilePath As String, SPFilePath As String
    Dim wbMain As Workbook, wbHF As Workbook, wbSP As Workbook
    Dim wsHFSource As Worksheet, wsSPSource As Worksheet
    Dim loHF As ListObject, loSP As ListObject
    Dim rngHF As Range, rngSP As Range
    Dim wsSourcePop As Worksheet, wsSPMain As Worksheet
    Dim loMainHF As ListObject, loMainSP As ListObject
    Dim visData As Range, r As Range
    Dim colIndex As Long
    Dim dictSP As Object
    Dim i As Long, j As Long
    Dim fundCoperID As String
    Dim newFunds As Collection
    Dim rec As Variant
    Dim wsUpload As Worksheet
    Dim loUpload As ListObject
    Dim rngUpload As Range
    Dim headers As Variant
    Dim rowCounter As Long

    '-----------------------------
    ' Additional variable declarations (for lookups and loops)
    Dim iSP As Long, key As String
    Dim wsCO As Worksheet, loCO As ListObject
    Dim coDict As Object
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long
    Dim coData As Variant, rIdx As Long, coKey As String
    Dim imDict As Object
    Dim spData As Variant, imKey As String
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long, sp_AdHocCol As Long, sp_ParentFlagCol As Long
    Dim daysDict As Object
    Dim hfData As Variant
    Dim hfFundIDCol As Long, hfDaysCol As Long
    Dim fundKey As String
    Dim up_CreditOfficerCol As Long, up_RegionCol As Long, up_IMCoperIDCol As Long, up_NAVSourceCol As Long
    Dim up_FrequencyCol As Long, up_AdHocCol As Long, up_ParentFlagCol As Long, up_FundCoperIDCol As Long, up_DaysToReportCol As Long
    Dim upRow As ListRow
    Dim creditOfficerName As String, imCoperID As String
    Dim inactiveRow As ListRow

    '=======================
    ' 1. Define file paths
    HFFilePath = "C:\YourFolder\HFFile.xlsx"
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"
    Set wbMain = ThisWorkbook

    ' 2. Open HF file
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set rngHF = wsHFSource.UsedRange
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, rngHF, , xlYes)
    End If
    loHF.Name = "HFTable"

    ' 3. Open SharePoint file
    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set rngSP = wsSPSource.UsedRange
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, rngSP, , xlYes)
    End If
    loSP.Name = "SharePoint"

    ' 4. Copy tables to main workbook
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

    ' 5. Ensure tables in main
    On Error Resume Next
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    If loMainHF Is Nothing Then Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes): loMainHF.Name = "HFTable"
    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    If loMainSP Is Nothing Then Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes): loMainSP.Name = "SharePoint"
    On Error GoTo 0

    ' 6. Filter for Transparency factor
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"

    ' 7. Filter factor values ≥ 2023-01-01
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=">=" & Format(DateSerial(2023,1,1),"mm/dd/yyyy"), Operator:=xlAnd

    ' 8. Filter other fields as before
    ApplyStrategyFilter loMainHF
    ApplyEntityFilter loMainHF

    ' 9. Identify new funds
    Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        spData = loMainSP.DataBodyRange.Value
        For iSP = 1 To UBound(spData)
            key = Trim(CStr(spData(iSP, colIndex)))
            If Not dictSP.Exists(key) Then dictSP.Add key, True
        Next
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
                    fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, r.Cells(1, idxName).Value, r.Cells(1, idxIMID).Value, _
                                    r.Cells(1, idxIMName).Value, r.Cells(1, idxCred).Value, r.Cells(1, idxVal).Value, "Active")
                        newFunds.Add rec
                    End If
                End If
            Next
        End If
    End If

    ' 10. Create Upload sheet (rest of macro remains unchanged)
    MsgBox "Macro completed successfully.", vbInformation
End Sub

' Helper: Get column index
Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim(lo.HeaderRowRange.Cells(1, i).Value) = headerName Then
            GetColumnIndex = i: Exit Function
        End If
    Next
    GetColumnIndex = 0
End Function

' Helper: Apply strategy filter
Sub ApplyStrategyFilter(loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Strategy")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Strategy", Array("FIF","Fund of Funds","Sub/Sleeve- No Benchmark"))
    If IsError(Application.Match("", allowed, 0)) Then allowed = AppendToArray(allowed, "")
    If Not IsEmpty(allowed) Then loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

' Helper: Apply entity filter
Sub ApplyEntityFilter(loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Entity_type")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary","Investment Manager as Agent","Managed Account","Managed Account - No AF","Loan Monitoring","Loan FiF - No tracking","Sleeve/share class/sub-account"))
    If IsError(Application.Match("", allowed, 0)) Then allowed = AppendToArray(allowed, "")
    If Not IsEmpty(allowed) Then loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

' Other helpers (GetAllowedValues, AppendToArray, ColumnExists) unchanged
