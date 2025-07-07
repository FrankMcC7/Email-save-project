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
    Dim dictSP As Object, dictHF As Object, tierDict As Object
    Dim newFunds As Collection, inactiveFunds As Collection
    Dim rec As Variant
    Dim i As Long, j As Long, iSP As Long, rIdx As Long
    Dim key As String, fundCoperID As String

    ' Variables for Upload sheet
    Dim wsUpload As Worksheet, loUpload As ListObject, rngUpload As Range
    Dim headers As Variant, rowCounter As Long
    ' Variables for CO_Table lookup
    Dim wsCO As Worksheet, loCO As ListObject, coDict As Object
    Dim coData As Variant, coKey As String
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long
    ' Variables for SharePoint IM lookup
    Dim imDict As Object, spData As Variant, imKey As String
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long, sp_AdHocCol As Long, sp_ParentFlagCol As Long
    ' Variables for Days to Report
    Dim daysDict As Object, hfData As Variant
    Dim hfFundIDCol As Long, hfDaysCol As Long
    ' Variables for Inactive funds sheet
    Dim wsInactive As Worksheet, loInactive As ListObject
    Dim share_CoperCol As Long, share_StatusCol As Long, share_CommentsCol As Long
    Dim arrSPInactive As Variant
    Dim fundStatus As String, fundComments As String
    Dim dbDict As Object, wbDB As Workbook, wsDB As Worksheet, loDB As ListObject, dbData As Variant
    Dim dbKey As String, dbValue As String
    Dim hfStatusCol As Long, reviewCol As Long

    '=======================
    ' 1. Define file paths
    HFFilePath = "C:\YourFolder\HFFile.xlsx"  ' update path
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"
    Set wbMain = ThisWorkbook

    ' 2. Open HF file and table
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1)
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, wsHFSource.UsedRange, , xlYes)
    End If
    loHF.Name = "HFTable"

    ' 3. Open SharePoint file and table
    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, wsSPSource.UsedRange, , xlYes)
    End If
    loSP.Name = "SharePoint"

    ' 4. Copy tables into main workbook
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

    ' 5. Ensure ListObjects exist in main
    On Error Resume Next
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    If loMainHF Is Nothing Then
        Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes)
        loMainHF.Name = "HFTable"
    End If
    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    If loMainSP Is Nothing Then
        Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes)
        loMainSP.Name = "SharePoint"
    End If
    On Error GoTo 0

    ' 6. Apply filters to HFTable
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData
    ' 6.1 Keep only Transparency factor
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"
    ' 6.2 Date filter on IRR_last_update_date
    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1: ">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), Operator:=xlAnd
    ' 6.3 Tier values 1 & 2
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:=Array("1", "2"), Operator:=xlFilterValues
    ' 6.4 Strategy and Entity filters
    ApplyStrategyFilter loMainHF
    ApplyEntityFilter loMainHF

    ' 7. Identify new funds not in SharePoint
    Set dictSP = CreateObject("Scripting.Dictionary"): dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
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
            Dim idxName As Long, idxIMID As Long, idxIMName As Long, idxCred As Long, idxVal As Long
            idxName = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
            idxIMID = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
            idxIMName = GetColumnIndex(loMainHF, "HFAD_IM_Name")
            idxCred = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
            idxVal = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
            For Each r In visData.Rows
                If Not r.EntireRow.Hidden Then
                    fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, r.Cells(1, idxName).Value, r.Cells(1, idxIMID).Value, r.Cells(1, idxIMName).Value, r.Cells(1, idxCred).Value, r.Cells(1, idxVal).Value, "Active")
                        newFunds.Add rec
                    End If
                End If
            Next r
        End If
    End If

    ' 8. Create "Upload to SP" sheet and table
    On Error Resume Next
    Set wsUpload = wbMain.Sheets("Upload to SP")
    If wsUpload Is Nothing Then
        Set wsUpload = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsUpload.Name = "Upload to SP"
    Else
        wsUpload.Cells.Clear
    End If
    On Error GoTo 0
    headers = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
    For j = LBound(headers) To UBound(headers)
        wsUpload.Cells(1, j + 1).Value = headers(j)
    Next j
    rowCounter = 2
    For i = 1 To newFunds.Count
        rec = newFunds(i)
        For j = LBound(rec) To UBound(rec)
            wsUpload.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next i
    Set rngUpload = wsUpload.Range(wsUpload.Cells(1, 1), wsUpload.Cells(rowCounter - 1, UBound(headers) + 1))
    On Error Resume Next
    Set loUpload = wsUpload.ListObjects("UploadHF")
    If loUpload Is Nothing Then
        Set loUpload = wsUpload.ListObjects.Add(xlSrcRange, rngUpload, , xlYes)
        loUpload.Name = "UploadHF"
    Else
        loUpload.Resize rngUpload
    End If
    On Error GoTo 0

    ' 9. Enhance "UploadHF"
    ' (a) CO_Table lookup
    On Error Resume Next
    Set wsCO = wbMain.Sheets("CO_Table")
    On Error GoTo 0
    If wsCO Is Nothing Then MsgBox "CO_Table sheet not found", vbCritical: Exit Sub
    Set loCO = wsCO.ListObjects("CO_Table")
    coCredCol = GetColumnIndex(loCO, "Credit Officer")
    coRegionCol = GetColumnIndex(loCO, "Region")
    coEmailCol = GetColumnIndex(loCO, "Email Address")
    coData = loCO.DataBodyRange.Value
    Set coDict = CreateObject("Scripting.Dictionary"): coDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(coData, 1)
        coKey = Trim(CStr(coData(rIdx, coCredCol)))
        If Not coDict.Exists(coKey) Then
            coDict.Add coKey, Array(coData(rIdx, coRegionCol), coData(rIdx, coEmailCol))
        End If
    Next rIdx
    ' (b) SharePoint IM lookup
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
    ' (c) Days to Report lookup
    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol = GetColumnIndex(loMainHF, "HFAD_Days_to_report")
    hfData = loMainHF.DataBodyRange.Value
    Set daysDict = CreateObject("Scripting.Dictionary"): daysDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(hfData, 1)
        fundCoperID = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not daysDict.Exists(fundCoperID) Then daysDict.Add fundCoperID, hfData(rIdx, hfDaysCol)
    Next rIdx
    ' Populate UploadHF enhancements
    For Each upRow In loUpload.ListRows
        creditOfficerName = Trim(upRow.Range.Cells(1, GetColumnIndex(loUpload, "HFAD_Credit_Officer")).Value)
        imCoperID = Trim(upRow.Range.Cells(1, GetColumnIndex(loUpload, "HFAD_IM_CoperID")).Value)
        fundCoperID = Trim(upRow.Range.Cells(1, GetColumnIndex(loUpload, "HFAD_Fund_CoperID")).Value)
        If coDict.Exists(creditOfficerName) Then
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "HFAD_Credit_Officer")).Value = coDict(creditOfficerName)(1)
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "Region")).Value = coDict(creditOfficerName)(0)
        End If
        If imDict.Exists(imCoperID) Then
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "NAV Source")).Value = imDict(imCoperID)(0)
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "Frequency")).Value = imDict(imCoperID)(1)
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "Ad-Hoc Reporting")).Value = imDict(imCoperID)(2)
            upRow.Range.Cells(1, GetColumnIndex(loUpload, "Parent/Flagship Reporting")).Value = imDict(imCoperID)(3)
        End If
        If daysDict.Exists(fundCoperID) Then upRow.Range.Cells(1, GetColumnIndex(loUpload, "Days to Report")).Value = daysDict(fundCoperID)
    Next upRow

    ' 10. Inactive funds tracking
    Set dictHF = CreateObject("Scripting.Dictionary"): dictHF.CompareMode = vbTextCompare
    Set tierDict = CreateObject("Scripting.Dictionary"): tierDict.CompareMode = vbTextCompare
    For rIdx = 1 To UBound(hfData, 1)
        fundCoperID = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not dictHF.Exists(fundCoperID) Then
            dictHF.Add fundCoperID, True
            tierDict.Add fundCoperID, hfData(rIdx, GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value"))
        End If
    Next rIdx
    Set inactiveFunds = New Collection
    share_CoperCol = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")
    arrSPInactive = loMainSP.DataBodyRange.Value
    For rIdx = 1 To UBound(arrSPInactive, 1)
        key = Trim(CStr(arrSPInactive(rIdx, share_CoperCol)))
        If Not dictHF.Exists(key) Then
            fundStatus = arrSPInactive(rIdx, share_StatusCol)
            fundComments = arrSPInactive(rIdx, share_CommentsCol)
            inactRec = Array(key, fundStatus, fundComments, tierDict(key))
            inactiveFunds.Add inactRec
        End If
    Next rIdx
    On Error Resume Next
    Set wsInactive = wbMain.Sheets("Inactive Funds Tracking")
    If wsInactive Is Nothing Then
        Set wsInactive = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsInactive.Name = "Inactive Funds Tracking"
    Else
        wsInactive.Cells.Clear
    End If
    On Error GoTo 0
    headers = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
    For j = LBound(headers) To UBound(headers)
        wsInactive.Cells(1, j + 1).Value = headers(j)
    Next j
    rowCounter = 2
    For i = 1 To inactiveFunds.Count
        inactRec = inactiveFunds(i)
        For j = LBound(inactRec) To UBound(inactRec)
            wsInactive.Cells(rowCounter, j + 1).Value = inactRec(j)
        Next j
        rowCounter = rowCounter + 1
    Next i
    Set rngUpload = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(rowCounter - 1, UBound(headers) + 1))
    On Error Resume Next
    Set loInactive = wsInactive.ListObjects("InactiveHF")
    If loInactive Is Nothing Then
        Set loInactive = wsInactive.ListObjects.Add(xlSrcRange, rngUpload, , xlYes)
        loInactive.Name = "InactiveHF"
    Else
        loInactive.Resize rngUpload
    End If
    On Error GoTo 0
    ' Remove truly inactive
    colIndex = GetColumnIndex(loInactive, "Status")
    For i = loInactive.ListRows.Count To 1 Step -1
        If StrComp(Trim(loInactive.ListRows(i).Range.Cells(1, colIndex).Value), "Inactive", vbTextCompare) = 0 Then
            loInactive.ListRows(i).Delete
        End If
    Next i

    ' 11. Optional DB status lookup
    Dim dbFilePath As String
    dbFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select Database File")
    If dbFilePath <> "False" Then
        Set wbDB = Workbooks.Open(dbFilePath)
        Set wsDB = wbDB.Sheets(1)
        On Error Resume Next
        Set loDB = wsDB.ListObjects("DBTable")
        If loDB Is Nothing Then Set loDB = wsDB.ListObjects.Add(xlSrcRange, wsDB.UsedRange, , xlYes)
        On Error GoTo 0
        dbData = loDB.DataBodyRange.Value
        Set dbDict = CreateObject("Scripting.Dictionary"): dbDict.CompareMode = vbTextCompare
        For rIdx = 1 To UBound(dbData, 1)
            dbKey = Trim(CStr(dbData(rIdx, GetColumnIndex(loDB, "Fund CoPer"))))
            dbValue = CStr(dbData(rIdx, GetColumnIndex(loDB, "Active")))
            If Not dbDict.Exists(dbKey) Then dbDict.Add dbKey, dbValue
        Next rIdx
        wbDB.Close False
        hfStatusCol = GetColumnIndex(loInactive, "HF Status")
        If hfStatusCol = 0 Then hfStatusCol = loInactive.ListColumns.Add.Name: hfStatusCol = GetColumnIndex(loInactive, "HF Status")
        For Each r In loInactive.DataBodyRange.Rows
            fundCoperID = Trim(CStr(r.Cells(1, GetColumnIndex(loInactive, "HFAD_Fund_CoperID")).Value))
            If dbDict.Exists(fundCoperID) Then r.Cells(1, hfStatusCol).Value = dbDict(fundCoperID)
            If Trim(r.Cells(1, GetColumnIndex(loInactive, "Status")).Value) <> Trim(r.Cells(1, hfStatusCol).Value) Then _
                r.Cells(1, GetColumnIndex(loInactive, "Review")).Value = "Check"
        Next r
    End If

    MsgBox "Macro completed successfully.", vbInformation
End Sub

'=======================
' Helper Functions

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
        If Trim(cl.Name) = colName Then
            ColumnExists = True
            Exit Function
        End If
    Next cl
    ColumnExists = False
End Function
