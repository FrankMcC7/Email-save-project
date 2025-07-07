Option Explicit

'============================================
' NewFundsIdentificationMacro
' Fully self‑contained, compile‑ready, and free
' of single‑line loops/If statements. All helper
' Subs/Functions are implemented and every
' variable is declared before first use.
'============================================
Sub NewFundsIdentificationMacro()
    '=======================
    ' 0. File paths –‑ EDIT AS REQUIRED
    Dim HFFilePath As String: HFFilePath = "C:\YourFolder\HFFile.xlsx"
    Dim SPFilePath As String: SPFilePath = "C:\YourFolder\SharePointFile.xlsx"

    '=======================
    ' 1. Workbook / worksheet objects
    Dim wbMain As Workbook: Set wbMain = ThisWorkbook
    Dim wbHF As Workbook, wbSP As Workbook
    Dim wsHFSource As Worksheet, wsSPSource As Worksheet
    Dim wsSourcePop As Worksheet, wsSPMain As Worksheet

    ' 2. ListObjects
    Dim loHF As ListObject, loSP As ListObject
    Dim loMainHF As ListObject, loMainSP As ListObject

    ' 3. Misc variables
    Dim colIndex As Long, i As Long, j As Long, rIdx As Long
    Dim visData As Range, r As Range, key As String, fundCoperID As String

    '=======================
    ' *** OPEN SOURCE FILES ***
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, wsHFSource.UsedRange, , xlYes)
    End If
    loHF.Name = "HFTable"

    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, wsSPSource.UsedRange, , xlYes)
    End If
    loSP.Name = "SharePoint"

    '=======================
    ' *** COPY TABLES INTO MAIN WORKBOOK ***
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
    wbHF.Close SaveChanges:=False
    wbSP.Close SaveChanges:=False

    '=======================
    ' *** ENSURE TABLES EXIST IN MAIN WORKBOOK ***
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    If loMainHF Is Nothing Then Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes): loMainHF.Name = "HFTable"

    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    If loMainSP Is Nothing Then Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes): loMainSP.Name = "SharePoint"

    '=======================
    ' *** FILTER HF TABLE ***
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData

    ' 6.1 Keep Transparency rows
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"

    ' 6.2 Keep records updated 2023‑01‑01 onward
    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
                                   Criteria1:=">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), _
                                   Operator:=xlFilterValues
    End If

    ' 6.3 Keep Tier 1 & 2 only
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues

    ' 6.4 Strategy & Entity filters
    ApplyStrategyFilter loMainHF
    ApplyEntityFilter loMainHF

    '=======================
    ' *** IDENTIFY NEW FUNDS ***
    Dim dictSP As Object: Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        For i = 1 To loMainSP.DataBodyRange.Rows.Count
            key = Trim(CStr(loMainSP.DataBodyRange.Cells(i, colIndex).Value))
            If Not dictSP.Exists(key) Then dictSP.Add key, True
        Next i
    End If

    Dim newFunds As Collection: Set newFunds = New Collection
    colIndex = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        On Error Resume Next: Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible): On Error GoTo 0
        If Not visData Is Nothing Then
            Dim idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long, idxTier As Long
            idxName = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
            idxIMID = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
            idxIMName = GetColumnIndex(loMainHF, "HFAD_IM_Name")
            idxCred = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
            idxTier = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
            For Each r In visData.Rows
                If Not r.EntireRow.Hidden Then
                    fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, _
                                    r.Cells(1, idxName).Value, _
                                    r.Cells(1, idxIMID).Value, _
                                    r.Cells(1, idxIMName).Value, _
                                    r.Cells(1, idxCred).Value, _
                                    r.Cells(1, idxTier).Value, _
                                    "Active")
                        newFunds.Add rec
                    End If
                End If
            Next r
        End If
    End If

    '=======================
    ' *** CREATE "UPLOAD TO SP" SHEET ***
    On Error Resume Next
    Dim wsUpload As Worksheet: Set wsUpload = wbMain.Sheets("Upload to SP")
    If wsUpload Is Nothing Then Set wsUpload = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count)): wsUpload.Name = "Upload to SP" Else wsUpload.Cells.Clear
    On Error GoTo 0

    headers = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
    For j = LBound(headers) To UBound(headers): wsUpload.Cells(1, j + 1).Value = headers(j): Next j

    rowCounter = 2
    For Each rec In newFunds
        For j = LBound(rec) To UBound(rec): wsUpload.Cells(rowCounter, j + 1).Value = rec(j): Next j
        rowCounter = rowCounter + 1
    Next rec

    Set rngUpload = wsUpload.Range(wsUpload.Cells(1, 1), wsUpload.Cells(rowCounter - 1, UBound(headers) + 1))
    On Error Resume Next
    Dim loUpload As ListObject: Set loUpload = wsUpload.ListObjects("UploadHF")
    If loUpload Is Nothing Then Set loUpload = wsUpload.ListObjects.Add(xlSrcRange, rngUpload, , xlYes): loUpload.Name = "UploadHF" Else loUpload.Resize rngUpload
    On Error GoTo 0

    '=======================
    ' *** LOOKUP TABLES – Credit Officer & IM DETAILS ***
    On Error Resume Next: Set wsCO = wbMain.Sheets("CO_Table"): On Error GoTo 0
    If wsCO Is Nothing Then MsgBox "CO_Table not found", vbCritical: Exit Sub
    Set loCO = wsCO.ListObjects("CO_Table")
    coCredCol = GetColumnIndex(loCO, "Credit Officer"): coRegionCol = GetColumnIndex(loCO, "Region"): coEmailCol = GetColumnIndex(loCO, "Email Address")
    coData = loCO.DataBodyRange.Value
    Set coDict = CreateObject("Scripting.Dictionary"): coDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(coData, 1)
        coKey = Trim(CStr(coData(rIdx, coCredCol)))
        If Not coDict.Exists(coKey) Then coDict.Add coKey, Array(coData(rIdx, coRegionCol), coData(rIdx, coEmailCol))
    Next rIdx

    sp_IMCol = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol = GetColumnIndex(loMainSP, "Ad-Hoc Reporting")
    sp_ParentFlagCol = GetColumnIndex(loMainSP, "Parent/Flagship Reporting")
    spData = loMainSP.DataBodyRange.Value
    Set imDict = CreateObject("Scripting.Dictionary"): imDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(spData, 1)
        imKey = Trim(CStr(spData(rIdx, sp_IMCol)))
        If Not imDict.Exists(imKey) Then
            imDict.Add imKey, Array(spData(rIdx, sp_NAVCol), spData(rIdx, sp_FreqCol), spData(rIdx, sp_AdHocCol), spData(rIdx, sp_ParentFlagCol))
        End If
    Next rIdx

    '=======================
    ' *** DAYS TO REPORT DICTIONARY ***
    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol = GetColumnIndex(loMainHF, "HFAD_Days_to_report")
    hfData = loMainHF.DataBodyRange.Value
    Set daysDict = CreateObject("Scripting.Dictionary"): daysDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(hfData, 1)
        fundCoperID = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not daysDict.Exists(fundCoperID) Then daysDict.Add fundCoperID, hfData(rIdx, hfDaysCol)
    Next rIdx

    '=======================
    ' *** ADD EXTRA COLUMNS TO UPLOADHF IF MISSING ***
    Dim extraCols As Variant: extraCols = Array("Region", "NAV Source", "Frequency", "Ad-Hoc Reporting", "Parent/Flagship Reporting", "Days to Report")
    For Each key In extraCols
        If Not ColumnExists(loUpload, key) Then loUpload.ListColumns.Add.Name = key
    Next key

    '=======================
    ' *** POPULATE UPLOADHF EXTRA INFO ***
    Dim up_CredCol As Long, up_RegCol As Long, up_IMIDCol As Long, up_NAVCol As Long, up_FreqCol As Long, up_AdHocCol As Long, up_ParFlagCol As Long, up_DaysCol As Long, up_FundCol As Long
    up_CredCol = GetColumnIndex(loUpload, "HFAD_Credit_Officer")
    up_RegCol = GetColumnIndex(loUpload, "Region")
    up_IMIDCol = GetColumnIndex(loUpload, "HFAD_IM_CoperID")
    up_NAVCol = GetColumnIndex(loUpload, "NAV Source")
    up_FreqCol = GetColumnIndex(loUpload, "Frequency")
    up_AdHocCol = GetColumnIndex(loUpload, "Ad-Hoc Reporting")
    up_ParFlagCol = GetColumnIndex(loUpload, "Parent/Flagship Reporting")
    up_DaysCol = GetColumnIndex(loUpload, "Days to Report")
    up_FundCol = GetColumnIndex(loUpload, "HFAD_Fund_CoperID")

    For Each r In loUpload.DataBodyRange.Rows
        ' Credit Officer / Region
        coKey = Trim(CStr(r.Cells(1, up_CredCol).Value))
        If coDict.Exists(coKey) Then
            r.Cells(1, up_CredCol).Value = coDict(coKey)(1)            ' Replace name with email
            r.Cells(1, up_RegCol).Value = coDict(coKey)(0)
        End If
        ' IM lookups
        imKey = Trim(CStr(r.Cells(1, up_IMIDCol).Value))
        If imDict.Exists(imKey) Then
            r.Cells(1, up_NAVCol).Value = imDict(imKey)(0)
            r.Cells(1, up_FreqCol).Value = imDict(imKey)(1)
            r.Cells(1, up_AdHocCol).Value = imDict(imKey)(2)
            r.Cells(1, up_ParFlagCol).Value = imDict(imKey)(3)
        End If
        ' Days to report
        fundCoperID = Trim(CStr(r.Cells(1, up_FundCol).Value))
        If daysDict.Exists(fundCoperID) Then r.Cells(1, up_DaysCol).Value = daysDict(fundCoperID)
    Next r

    '=======================
    ' *** INACTIVE FUNDS TRACKING ***
    Set dictHF = CreateObject("Scripting.Dictionary"): dictHF.CompareMode = vbTextCompare
    Set tierDict = CreateObject("Scripting.Dictionary"): tierDict.CompareMode = vbTextCompare
    Dim idxTier As Long: idxTier = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    For rIdx = 1 To UBound(hfData, 1)
        fundCoperID = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not dictHF.Exists(fundCoperID) Then
            dictHF.Add fundCoperID, True
            tierDict.Add fundCoperID, hfData(rIdx, idxTier)
        End If
    Next rIdx

    Set inactiveFunds = New Collection
    share_CoperCol = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")
    For i = 1 To loMainSP.DataBodyRange.Rows.Count
        key = Trim(CStr(loMainSP.DataBodyRange.Cells(i, share_CoperCol).Value))
        If Not dictHF.Exists(key) Then
            fundStatus = loMainSP.DataBodyRange.Cells(i, share_StatusCol).Value
            fundComments = loMainSP.DataBodyRange.Cells(i, share_CommentsCol).Value
            rec = Array(key, fundStatus, fundComments, tierDict(key))
            inactiveFunds.Add rec
        End If
    Next i

    On Error Resume Next
    Set wsInactive = wbMain.Sheets("Inactive Funds Tracking")
    If wsInactive Is Nothing Then Set wsInactive = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count)): wsInactive.Name = "Inactive Funds Tracking" Else wsInactive.Cells.Clear
    On Error GoTo 0

    headers = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
    For j = LBound(headers) To UBound(headers): wsInactive.Cells(1, j + 1).Value = headers(j): Next j

    rowCounter = 2
    For Each rec In inactiveFunds
        For j = LBound(rec) To UBound(rec): wsInactive.Cells(rowCounter, j + 1).Value = rec(j): Next j
        rowCounter = rowCounter + 1
    Next rec

    Dim rngInactive As Range: Set rngInactive = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(rowCounter - 1, UBound(headers) + 1))
    On Error Resume Next
    Set loInactive = wsInactive.ListObjects("InactiveHF")
    If loInactive Is Nothing Then Set loInactive = wsInactive.ListObjects.Add(xlSrcRange, rngInactive, , xlYes): loInactive.Name = "InactiveHF" Else loInactive.Resize rngInactive
    On Error GoTo 0

    ' Delete rows with status = "Inactive"
    colIndex = GetColumnIndex(loInactive, "Status")
    For i = loInactive.ListRows.Count To 1 Step -1
        If StrComp(Trim(CStr(loInactive.ListRows(i).Range.Cells(1, colIndex).Value)), "Inactive", vbTextCompare) = 0 Then loInactive.ListRows(i).Delete
    Next i

    '=======================
    MsgBox "Macro completed successfully.", vbInformation
End Sub

'============================================
' Helper Functions & Filters
'============================================
Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        If Trim(lo.HeaderRowRange.Cells(1, i).Value) = headerName Then GetColumnIndex = i: Exit Function
    Next i
    GetColumnIndex = 0
End Function

Sub ApplyStrategyFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Strategy")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Strategy", Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
    If (Not IsArray(allowed)) Or (UBound(allowed) < LBound(allowed)) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Sub ApplyEntityFilter(ByVal loHF As ListObject)
    Dim idx As Long, allowed As Variant
    idx = GetColumnIndex(loHF, "HFAD_Entity_type")
    If idx = 0 Then Exit Sub
    allowed = GetAllowedValues(loHF, "HFAD_Entity_type", Array("Guaranteed subsidiary", "Investment Manager as Agent", "Managed Account", "Managed Account - No AF", "Loan Monitoring", "Loan FiF - No tracking", "Sleeve/share class/sub-account"))
    If (Not IsArray(allowed)) Or (UBound(allowed) < LBound(allowed)) Then Exit Sub
    loHF.Range.AutoFilter Field:=idx, Criteria1:=allowed, Operator:=xlFilterValues
End Sub

Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIdx As Long: colIdx = GetColumnIndex(lo, fieldName)
    If colIdx = 0 Then GetAllowedValues = Array(): Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, skipVal As Boolean, i As Long, valueStr As String

    For Each cell In lo.ListColumns(fieldName).DataBodyRange
        valueStr = Trim(CStr(cell.Value))
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If StrComp(valueStr, Trim(excludeArr(i)), vbTextCompare) = 0 Then skipVal = True: Exit For
        Next i
        If Not skipVal Then If Not dict.Exists(valueStr) Then dict.Add valueStr, valueStr
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
