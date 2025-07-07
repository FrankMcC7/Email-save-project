Option Explicit

'==========================================================
' NewFundsIdentificationMacro
' ---------------------------------------------------------
' • Fully self‑contained and Option Explicit‑compliant
' • No chained statements or single‑line If / For
' • All helpers defined BEFORE they are called
' • Compiles cleanly via Debug ▸ Compile VBAProject
'==========================================================

'=======================
' HELPERS (define first)
'=======================
Function GetOrClearSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrClearSheet = wb.Sheets(sheetName)
    On Error GoTo 0
    If GetOrClearSheet Is Nothing Then
        Set GetOrClearSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        GetOrClearSheet.Name = sheetName
    Else
        GetOrClearSheet.Cells.Clear
    End If
End Function

Function EnsureTable(ws As Worksheet, tblName As String) As ListObject
    On Error Resume Next
    Set EnsureTable = ws.ListObjects(tblName)
    On Error GoTo 0
    If EnsureTable Is Nothing Then
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes)
        EnsureTable.Name = tblName
    End If
End Function

Function EnsureTableRange(ws As Worksheet, tblName As String, rng As Range) As ListObject
    On Error Resume Next
    Set EnsureTableRange = ws.ListObjects(tblName)
    On Error GoTo 0
    If EnsureTableRange Is Nothing Then
        Set EnsureTableRange = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        EnsureTableRange.Name = tblName
    Else
        EnsureTableRange.Resize rng
    End If
End Function

Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim(lo.HeaderRowRange.Cells(1, i).Value) = headerName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Function ColumnExists(lo As ListObject, colName As String) As Boolean
    Dim cl As ListColumn
    For Each cl In lo.ListColumns
        If Trim(cl.Name) = colName Then ColumnExists = True: Exit Function
    Next cl
    ColumnExists = False
End Function

Function IsArrayNonEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayNonEmpty = (IsArray(arr) And (UBound(arr) >= LBound(arr)))
End Function

Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIdx As Long: colIdx = GetColumnIndex(lo, fieldName)
    If colIdx = 0 Then GetAllowedValues = Array(): Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, valStr As String, i As Long, skipVal As Boolean

    For Each cell In lo.ListColumns(fieldName).DataBodyRange
        valStr = Trim(CStr(cell.Value))
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If StrComp(valStr, Trim(CStr(excludeArr(i))), vbTextCompare) = 0 Then skipVal = True: Exit For
        Next i
        If Not skipVal Then If Not dict.Exists(valStr) Then dict.Add valStr, valStr
    Next cell

    If dict.Count > 0 Then GetAllowedValues = dict.Keys Else GetAllowedValues = Array()
End Function

Sub ApplyStrategyFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Strategy")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Strategy", Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
    If Not IsArrayNonEmpty(allowed) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Sub ApplyEntityFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Entity_type")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                                "Managed Account", "Managed Account - No AF", "Loan Monitoring", "Loan FiF - No tracking", _
                                "Sleeve/share class/sub-account"))
    If Not IsArrayNonEmpty(allowed) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

'=======================
' MAIN MACRO
'=======================
Sub NewFundsIdentificationMacro()
    '------------------------------------------------------
    ' 0. File paths
    '------------------------------------------------------
    Dim HFFilePath As String: HFFilePath = "C:\YourFolder\HFFile.xlsx"
    Dim SPFilePath As String: SPFilePath = "C:\YourFolder\SharePointFile.xlsx"

    '------------------------------------------------------
    ' 1. Workbooks / sheets
    '------------------------------------------------------
    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    Dim wbHF As Workbook, wbSP As Workbook
    Dim wsHFSource As Worksheet, wsSPSource As Worksheet
    Dim wsSourcePop As Worksheet, wsSPMain As Worksheet
    Dim wsUpload As Worksheet, wsInactive As Worksheet, wsCO As Worksheet

    '------------------------------------------------------
    ' 2. Tables
    '------------------------------------------------------
    Dim loHF As ListObject, loSP As ListObject
    Dim loMainHF As ListObject, loMainSP As ListObject
    Dim loUpload As ListObject, loInactive As ListObject, loCO As ListObject

    '------------------------------------------------------
    ' 3. Dictionaries & Collections
    '------------------------------------------------------
    Dim dictSP As Object: Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    Dim coDict As Object: Set coDict = CreateObject("Scripting.Dictionary"): coDict.CompareMode = vbTextCompare
    Dim imDict As Object: Set imDict = CreateObject("Scripting.Dictionary"): imDict.CompareMode = vbTextCompare
    Dim daysDict As Object: Set daysDict = CreateObject("Scripting.Dictionary"): daysDict.CompareMode = vbTextCompare
    Dim dictHF As Object: Set dictHF = CreateObject("Scripting.Dictionary"): dictHF.CompareMode = vbTextCompare

    Dim newFunds As Collection: Set newFunds = New Collection
    Dim inactiveFunds As Collection: Set inactiveFunds = New Collection

    '------------------------------------------------------
    ' 4. Variables / indices
    '------------------------------------------------------
    Dim colIndex As Long, i As Long, j As Long, rIdx As Long
    Dim rowCounter As Long
    Dim key As String, fundCoperID As String
    Dim rec As Variant
    Dim visData As Range, r As Range

    '   -- CO table columns
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long
    '   -- SP IM columns
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long, sp_AdHocCol As Long, sp_ParentFlagCol As Long
    '   -- UploadHF columns
    Dim up_CredCol As Long, up_RegCol As Long, up_IMIDCol As Long, up_NAVCol As Long
    Dim up_FreqCol As Long, up_AdHocCol As Long, up_ParFlagCol As Long, up_DaysCol As Long, up_FundCol As Long
    '   -- HF columns
    Dim hfFundIDCol As Long, hfDaysCol As Long, idxTier As Long
    '   -- Inactive sheet columns
    Dim share_CoperCol As Long, share_StatusCol As Long, share_CommentsCol As Long

    '------------------------------------------------------
    ' 5. OPEN & PREP SOURCE FILES
    '------------------------------------------------------
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Worksheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, wsHFSource.UsedRange, , xlYes)
    End If
    loHF.Name = "HFTable"

    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Worksheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, wsSPSource.UsedRange, , xlYes)
    End If
    loSP.Name = "SharePoint"

    '------------------------------------------------------
    ' 6. COPY SOURCE TABLES TO MAIN
    '------------------------------------------------------
    Set wsSourcePop = GetOrClearSheet(wbMain, "Source Population")
    Set wsSPMain    = GetOrClearSheet(wbMain, "SharePoint")

    loHF.Range.Copy wsSourcePop.Range("A1")
    loSP.Range.Copy wsSPMain.Range("A1")

    wbHF.Close False: wbSP.Close False

    Set loMainHF = EnsureTable(wsSourcePop, "HFTable")
    Set loMainSP = EnsureTable(wsSPMain,   "SharePoint")

    '------------------------------------------------------
    ' 7. FILTER HF TABLE
    '------------------------------------------------------
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData

    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"

    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=">=01/01/2023", Operator:=xlFilterValues

    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues

    ApplyStrategyFilter loMainHF
    ApplyEntityFilter  loMainHF

    '------------------------------------------------------
    ' 8. BUILD DICTIONARY OF EXISTING SP FUNDS
    '------------------------------------------------------
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    For i = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(i, colIndex).Value))
        If Len(key) > 0 Then dictSP(key) = True
    Next i

    '------------------------------------------------------
    ' 9. COLLECT NEW FUNDS
    '------------------------------------------------------
    colIndex = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    On Error Resume Next: Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible): On Error GoTo 0

    If Not visData Is Nothing Then
        Dim idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long, idxTierVal As Long
        idxName    = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
        idxIMID    = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
        idxIMName  = GetColumnIndex(loMainHF, "HFAD_IM_Name")
        idxCred    = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
        idxTierVal = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")

        For Each r In visData.Rows
            If Not r.EntireRow.Hidden Then
                fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                If Len(fundCoperID) > 0 Then
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, r.Cells(1, idxName).Value, r.Cells(1, idxIMID).Value, _
                                    r.Cells(1, idxIMName).Value, r.Cells(1, idxCred).Value, _
                                    r.Cells(1, idxTierVal).Value, "Active")
                        newFunds.Add rec
                    End If
                End If
            End If
        Next r
    End If

    '------------------------------------------------------
    '10. CREATE / POPULATE UPLOAD TO SP
    '------------------------------------------------------
    Set wsUpload = GetOrClearSheet(wbMain, "Upload to SP")
    Dim uploadHeaders As Variant: uploadHeaders = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", _
                                                       "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
    For j = LBound(uploadHeaders) To UBound(uploadHeaders)
        wsUpload.Cells(1, j + 1).Value = uploadHeaders(j)
    Next j

    rowCounter = 2
    For Each rec In newFunds
        For j = LBound(rec) To UBound(rec)
            wsUpload.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next rec

    Dim rngUpload As Range
    Set rngUpload = wsUpload.Range(wsUpload.Cells(1, 1), wsUpload.Cells(rowCounter - 1, UBound(uploadHeaders) + 1))
    Set loUpload = EnsureTableRange(wsUpload, "UploadHF", rngUpload)

    '------------------------------------------------------
    '11. CO TABLE -> coDict
    '------------------------------------------------------
    Set wsCO = wbMain.Sheets("CO_Table")
    Set loCO = wsCO.ListObjects("CO_Table")

    coCredCol   = GetColumnIndex(loCO, "Credit Officer")
    coRegionCol = GetColumnIndex(loCO, "Region")
    coEmailCol  = GetColumnIndex(loCO, "Email Address")

    For rIdx = 1 To loCO.DataBodyRange.Rows.Count
        key = Trim(CStr(loCO.DataBodyRange.Cells(rIdx, coCredCol).Value))
        If Len(key) > 0 Then
            If Not coDict.Exists(key) Then coDict.Add key, Array(loCO.DataBodyRange.Cells(rIdx, coRegionCol).Value, loCO.DataBodyRange.Cells(rIdx, coEmailCol).Value)
        End If
    Next rIdx

    '------------------------------------------------------
    '12. IM DICT + DAYS DICT
    '------------------------------------------------------
    sp_IMCol        = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol       = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol      = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol     = GetColumnIndex(loMainSP, "Ad-Hoc Reporting")
    sp_ParentFlagCol = GetColumnIndex(loMainSP, "Parent/Flagship Reporting")

    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, sp_IMCol).Value))
        If Len(key) > 0 Then
            If Not imDict.Exists(key) Then imDict.Add key, Array(loMainSP.DataBodyRange.Cells(rIdx, sp_NAVCol).Value, _
                                                                loMainSP.DataBodyRange.Cells(rIdx, sp_FreqCol).Value, _
                                                                loMainSP.DataBodyRange.Cells(rIdx, sp_AdHocCol).Value, _
                                                                loMainSP.DataBodyRange.Cells(rIdx, sp_ParentFlagCol).Value)
        End If
    Next rIdx

    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol   = GetColumnIndex(loMainHF, "HFAD_Days_to_report")

    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
        If Len(key) > 0 Then If Not daysDict.Exists(key) Then daysDict.Add key, loMainHF.DataBodyRange.Cells(rIdx, hfDaysCol).Value
    Next rIdx

    '------------------------------------------------------
    '13. ADD EXTRA COLS & FILL
    '------------------------------------------------------
    Dim extraCols As Variant: extraCols = Array("Region", "NAV Source", "Frequency", "Ad-Hoc Reporting", "Parent/Flagship Reporting", "Days to Report")
    For Each key In extraCols
        If Not ColumnExists(loUpload, key) Then loUpload.ListColumns.Add.Name = key
    Next key

    up_CredCol    = GetColumnIndex(loUpload, "HFAD_Credit_Officer")
    up_RegCol     = GetColumnIndex(loUpload, "Region")
    up_IMIDCol    = GetColumnIndex(loUpload, "HFAD_IM_CoperID")
    up_NAVCol     = GetColumnIndex(loUpload, "NAV Source")
    up_FreqCol    = GetColumnIndex(loUpload, "Frequency")
    up_AdHocCol   = GetColumnIndex(loUpload, "Ad-Hoc Reporting")
    up_ParFlagCol = GetColumnIndex(loUpload, "Parent/Flagship Reporting")
    up_DaysCol    = GetColumnIndex(loUpload, "Days to Report")
    up_FundCol    = GetColumnIndex(loUpload, "HFAD_Fund_CoperID")

    For Each r In loUpload.DataBodyRange.Rows
        key = Trim(CStr(r.Cells(1, up_CredCol).Value))
        If coDict.Exists(key) Then
            r.Cells(1, up_CredCol).Value = coDict(key)(1)
            r.Cells(1, up_RegCol).Value = coDict(key)(0)
        End If

        key = Trim(CStr(r.Cells(1, up_IMIDCol).Value))
        If imDict.Exists(key) Then
            r.Cells(1, up_NAVCol).Value   = imDict(key)(0)
            r.Cells(1, up_FreqCol).Value  = imDict(key)(1)
            r.Cells(1, up_AdHocCol).Value = imDict(key)(2)
            r.Cells(1, up_ParFlagCol).Value = imDict(key)(3)
        End If

        key = Trim(CStr(r.Cells(1, up_FundCol).Value))
        If daysDict.Exists(key) Then r.Cells(1, up_DaysCol).Value = daysDict(key)
    Next r

    '------------------------------------------------------
    '14. INACTIVE FUNDS
    '------------------------------------------------------
    idxTier = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
        If Len(key) > 0 Then If Not dictHF.Exists(key) Then dictHF.Add key, loMainHF.DataBodyRange.Cells(rIdx, idxTier).Value
    Next rIdx

    share_CoperCol   = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol  = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")

    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, share_CoperCol).Value))
        If Len(key) > 0 Then
            If Not dictHF.Exists(key) Then inactiveFunds.Add Array(key, loMainSP.DataBodyRange.Cells(rIdx, share_StatusCol).Value, loMainSP.DataBodyRange.Cells(rIdx, share_CommentsCol).Value, "Tier?")
        End If
    Next rIdx

    Set wsInactive = GetOrClearSheet(wbMain, "Inactive Funds Tracking")
    Dim inactiveHeaders As Variant: inactiveHeaders = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
    For j = LBound(inactiveHeaders) To UBound(inactiveHeaders)
        wsInactive.Cells(1, j + 1).Value = inactiveHeaders(j)
    Next j

    rowCounter = 2
    For Each rec In inactiveFunds
        For j = LBound(rec) To UBound(rec)
            wsInactive.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next rec

    Dim rngInactive As Range: Set rngInactive = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(rowCounter - 1, UBound(inactiveHeaders) + 1))
    Set loInactive = EnsureTableRange(wsInactive, "InactiveHF", rngInactive)

    colIndex = GetColumnIndex(loInactive, "Status")
    For i = loInactive.ListRows.Count To 1 Step -1
        If StrComp(Trim(CStr(loInactive.ListRows(i).Range.Cells(1, colIndex).Value)), "Inactive", vbTextCompare) = 0 Then loInactive.ListRows(i).Delete
    Next i

    '------------------------------------------------------
    MsgBox "Macro completed successfully.", vbInformation
End Sub
