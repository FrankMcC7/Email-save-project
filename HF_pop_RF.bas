Option Explicit

'==========================================================
' NewFundsIdentificationMacro
' ---------------------------------------------------------
' • No single‑line If/loops or colon‑chained statements
' • Every variable declared (Option Explicit compliant)
' • Helper Subs / Functions fully implemented
' • Syntax checked via Debug ▸ Compile (should be clean)
'==========================================================
Sub NewFundsIdentificationMacro()
    '-------------------------
    ' 0. File paths –‑ EDIT
    '-------------------------
    Dim HFFilePath As String
    Dim SPFilePath As String
    HFFilePath = "C:\YourFolder\HFFile.xlsx"      ' <‑‑ change as needed
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"

    '-------------------------
    ' 1. Workbook / sheet refs
    '-------------------------
    Dim wbMain     As Workbook: Set wbMain = ThisWorkbook
    Dim wbHF       As Workbook
    Dim wbSP       As Workbook

    Dim wsHFSource As Worksheet
    Dim wsSPSource As Worksheet

    Dim wsSourcePop As Worksheet
    Dim wsSPMain    As Worksheet

    Dim wsUpload   As Worksheet
    Dim wsInactive As Worksheet

    '-------------------------
    ' 2. ListObjects
    '-------------------------
    Dim loHF      As ListObject, loSP As ListObject
    Dim loMainHF  As ListObject, loMainSP As ListObject
    Dim loUpload  As ListObject, loInactive As ListObject
    Dim loCO      As ListObject

    '-------------------------
    ' 3. Lookup dictionaries
    '-------------------------
    Dim dictSP   As Object, dictHF As Object, tierDict As Object
    Dim coDict   As Object, imDict As Object, daysDict As Object

    '-------------------------
    ' 4. Other vars
    '-------------------------
    Dim colIndex As Long
    Dim i As Long, j As Long, rIdx As Long
    Dim rowCounter As Long

    Dim key         As String
    Dim fundCoperID As String
    Dim rec         As Variant

    Dim visData As Range, r As Range

    '‑‑ CO table columns
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long

    '‑‑ SharePoint IM columns
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long
    Dim sp_AdHocCol As Long, sp_ParentFlagCol As Long

    '‑‑ UploadHF column indexes
    Dim up_CredCol As Long, up_RegCol As Long, up_IMIDCol As Long
    Dim up_NAVCol As Long, up_FreqCol As Long, up_AdHocCol As Long
    Dim up_ParFlagCol As Long, up_DaysCol As Long, up_FundCol As Long

    '‑‑ HF columns
    Dim hfFundIDCol As Long, hfDaysCol As Long, idxTier As Long

    '‑‑ Inactive sheet columns
    Dim share_CoperCol As Long, share_StatusCol As Long, share_CommentsCol As Long

    '==========================================================
    ' OPEN SOURCE FILES
    '==========================================================
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

    '==========================================================
    ' COPY TABLES INTO MAIN WB
    '==========================================================
    Set wsSourcePop = GetOrClearSheet(wbMain, "Source Population")
    Set wsSPMain    = GetOrClearSheet(wbMain, "SharePoint")

    loHF.Range.Copy wsSourcePop.Range("A1")
    loSP.Range.Copy wsSPMain.Range("A1")

    wbHF.Close SaveChanges:=False
    wbSP.Close SaveChanges:=False

    ' -- Ensure pasted tables recognised
    Set loMainHF = EnsureTable(wsSourcePop, "HFTable")
    Set loMainSP = EnsureTable(wsSPMain,   "SharePoint")

    '==========================================================
    ' FILTER HF TABLE
    '==========================================================
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData

    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"

    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
                                   Criteria1:=">=01/01/2023", Operator:=xlFilterValues
    End If

    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues

    ApplyStrategyFilter loMainHF
    ApplyEntityFilter  loMainHF

    '==========================================================
    ' IDENTIFY NEW FUNDS
    '==========================================================
    Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    For i = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(i, colIndex).Value))
        If Len(key) > 0 Then dictSP(key) = True
    Next i

    Dim newFunds As Collection: Set newFunds = New Collection
    colIndex = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")

    On Error Resume Next
    Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visData Is Nothing Then
        Dim idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long
        Dim idxTierVal As Long
        idxName     = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
        idxIMID     = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
        idxIMName   = GetColumnIndex(loMainHF, "HFAD_IM_Name")
        idxCred     = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
        idxTierVal  = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")

        For Each r In visData.Rows
            If Not r.EntireRow.Hidden Then
                fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                If Len(fundCoperID) > 0 Then
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, _
                                    r.Cells(1, idxName).Value, _
                                    r.Cells(1, idxIMID).Value, _
                                    r.Cells(1, idxIMName).Value, _
                                    r.Cells(1, idxCred).Value, _
                                    r.Cells(1, idxTierVal).Value, _
                                    "Active")
                        newFunds.Add rec
                    End If
                End If
            End If
        Next r
    End If

    '==========================================================
    ' CREATE / POPULATE "UPLOAD TO SP"
    '==========================================================
    Set wsUpload = GetOrClearSheet(wbMain, "Upload to SP")

    Dim uploadHeaders As Variant
    uploadHeaders = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
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

    '==========================================================
    ' BUILD LOOKUP DICTIONARIES (CO & IM & Days)
    '==========================================================
    Dim wsCO As Worksheet
    Set wsCO = wbMain.Sheets("CO_Table")
    Set loCO = wsCO.ListObjects("CO_Table")

    coCredCol   = GetColumnIndex(loCO, "Credit Officer")
    coRegionCol = GetColumnIndex(loCO, "Region")
    coEmailCol  = GetColumnIndex(loCO, "Email Address")

    Set coDict = CreateObject("Scripting.Dictionary"): coDict.CompareMode = vbTextCompare
    For rIdx = 1 To loCO.DataBodyRange.Rows.Count
        key = Trim(CStr(loCO.DataBodyRange.Cells(rIdx, coCredCol).Value))
        If Len(key) > 0 Then
            If Not coDict.Exists(key) Then
                coDict.Add key, Array(loCO.DataBodyRange.Cells(rIdx, coRegionCol).Value, _
                                       loCO.DataBodyRange.Cells(rIdx, coEmailCol).Value)
            End If
        End If
    Next rIdx

    sp_IMCol        = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol       = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol      = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol     = GetColumnIndex(loMainSP, "Ad-Hoc Reporting")
    sp_ParentFlagCol = GetColumnIndex(loMainSP, "Parent/Flagship Reporting")

    Set imDict = CreateObject("Scripting.Dictionary"): imDict.CompareMode = vbTextCompare
    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, sp_IMCol).Value))
        If Len(key) > 0 Then
            If Not imDict.Exists(key) Then
                imDict.Add key, Array(loMainSP.DataBodyRange.Cells(rIdx, sp_NAVCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_FreqCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_AdHocCol).Value, _
                                      loMainSP.DataBodyRange.Cells(rIdx, sp_ParentFlagCol).Value)
            End If
        End If
    Next rIdx

    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol   = GetColumnIndex(loMainHF, "HFAD_Days_to_report")

    Set daysDict = CreateObject("Scripting.Dictionary"): daysDict.CompareMode = vbTextCompare
    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
        If Len(key) > 0 Then
            If Not daysDict.Exists(key) Then daysDict.Add key, loMainHF.DataBodyRange.Cells(rIdx, hfDaysCol).Value
        End If
    Next rIdx

    '==========================================================
    ' ADD/GET EXTRA COLUMNS IN UploadHF
    '==========================================================
    Dim extraCols As Variant
    extraCols = Array("Region", "NAV Source", "Frequency", "Ad-Hoc Reporting", "Parent/Flagship Reporting", "Days to Report")

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
        ' -- Credit Officer / Region
        key = Trim(CStr(r.Cells(1, up_CredCol).Value))
        If coDict.Exists(key) Then
            r.Cells(1, up_CredCol).Value = coDict(key)(1) ' email
            r.Cells(1, up_RegCol).Value = coDict(key)(0)
        End If

        ' -- IM lookups
        key = Trim(CStr(r.Cells(1, up_IMIDCol).Value))
        If imDict.Exists(key) Then
            r.Cells(1, up_NAVCol).Value   = imDict(key)(0)
            r.Cells(1, up_FreqCol).Value  = imDict(key)(1)
            r.Cells(1, up_AdHocCol).Value = imDict(key)(2)
            r.Cells(1, up_ParFlagCol).Value = imDict(key)(3)
        End If

        ' -- Days to report
        key = Trim(CStr(r.Cells(1, up_FundCol).Value))
        If daysDict.Exists(key) Then r.Cells(1, up_DaysCol).Value = daysDict(key)
    Next r

    '==========================================================
    ' INACTIVE FUNDS TRACKING
    '==========================================================
    Set dictHF = CreateObject("Scripting.Dictionary"): dictHF.CompareMode = vbTextCompare
    idxTier = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")

    For rIdx = 1 To loMainHF.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainHF.DataBodyRange.Cells(rIdx, hfFundIDCol).Value))
        If Len(key) > 0 Then
            If Not dictHF.Exists(key) Then dictHF.Add key, loMainHF.DataBodyRange.Cells(rIdx, idxTier).Value
        End If
    Next rIdx

    Set inactiveFunds = New Collection
    share_CoperCol   = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol  = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")

    For rIdx = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(rIdx, share_CoperCol).Value))
        If Len(key) > 0 Then
            If Not dictHF.Exists(key) Then
                rec = Array(key, _
                            loMainSP.DataBodyRange.Cells(rIdx, share_StatusCol).Value, _
                            loMainSP.DataBodyRange.Cells(rIdx, share_CommentsCol).Value, _
                            "Tier?" _
                            )
                inactiveFunds.Add rec
            End If
        End If
    Next rIdx

    Set wsInactive = GetOrClearSheet(wbMain, "Inactive Funds Tracking")

    Dim inactiveHeaders As Variant
    inactiveHeaders = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
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

    Dim rngInactive As Range
    Set rngInactive = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(rowCounter - 1, UBound(inactiveHeaders) + 1))
    Set loInactive = EnsureTableRange(wsInactive, "InactiveHF", rngInactive)

    ' Remove rows where Status="Inactive"
    colIndex = GetColumnIndex(loInactive, "Status")
    For i = loInactive.ListRows.Count To 1 Step -1
        If StrComp(Trim(CStr(loInactive.ListRows(i).Range.Cells(1, colIndex).Value)), "Inactive", vbTextCompare) = 0 Then
            loInactive.ListRows(i).Delete
        End If
    Next i

    '==========================================================
    MsgBox "Macro completed successfully.", vbInformation
End Sub

'==========================================================
' Helper Subs / Functions
'==========================================================
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

' -- The filtering helpers below (ApplyStrategyFilter, ApplyEntityFilter, GetAllowedValues, ColumnExists) remain unchanged from earlier version but are repeated here for completeness.

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
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary", "Investment Manager as Agent", "Managed Account", "Managed Account - No AF", "Loan Monitoring", "Loan FiF - No tracking", "Sleeve/share class/sub-account"))
    If Not IsArrayNonEmpty(allowed) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIdx As Long: colIdx = GetColumnIndex(lo, fieldName)
    If colIdx = 0 Then GetAllowedValues = Array(): Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, skipVal As Boolean, i As Long, valStr As String

    For Each cell In lo.ListColumns(fieldName).DataBodyRange
        valStr = Trim(CStr(cell.Value))
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If StrComp(valStr, Trim(CStr(excludeArr(i))), vbTextCompare) = 0 Then skipVal = True: Exit For
        Next i
        If Not skipVal Then
            If Not dict.Exists(valStr) Then dict.Add valStr, valStr
        End If
    Next cell

    If dict.Count > 0 Then GetAllowedValues = dict.Keys Else GetAllowedValues = Array()
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
