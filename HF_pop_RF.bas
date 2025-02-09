Option Explicit

Sub NewFundsIdentificationMacro()
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
    
    ' === 1. Define file paths (hardcoded) ===
    HFFilePath = "C:\YourFolder\HFFile.xlsx"           ' <<< Change to your actual path
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"     ' <<< Change to your actual path
    
    Set wbMain = ThisWorkbook
    
    ' === 2. Open the HF file and convert its data to table "HFTable" ===
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1) ' Assumes data is in the first sheet
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set rngHF = wsHFSource.UsedRange
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, rngHF, , xlYes)
    End If
    loHF.Name = "HFTable"
    
    ' === 2. Open the SharePoint file and convert its data to table "SharePoint" ===
    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set rngSP = wsSPSource.UsedRange
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, rngSP, , xlYes)
    End If
    loSP.Name = "SharePoint"
    
    ' === 3. Paste the tables into the main workbook ============
    ' Create (or clear) sheet "Source Population" for HF data
    On Error Resume Next
    Set wsSourcePop = wbMain.Sheets("Source Population")
    On Error GoTo 0
    If wsSourcePop Is Nothing Then
        Set wsSourcePop = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSourcePop.Name = "Source Population"
    Else
        wsSourcePop.Cells.Clear
    End If
    
    ' Create (or clear) sheet "SharePoint" for SP data
    On Error Resume Next
    Set wsSPMain = wbMain.Sheets("SharePoint")
    On Error GoTo 0
    If wsSPMain Is Nothing Then
        Set wsSPMain = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSPMain.Name = "SharePoint"
    Else
        wsSPMain.Cells.Clear
    End If
    
    ' Copy HF table into "Source Population"
    loHF.Range.Copy Destination:=wsSourcePop.Range("A1")
    ' Copy SharePoint table into "SharePoint"
    loSP.Range.Copy Destination:=wsSPMain.Range("A1")
    
    ' Close the source workbooks without saving changes
    wbHF.Close SaveChanges:=False
    wbSP.Close SaveChanges:=False
    
    ' Convert the pasted HF data into a ListObject (if not already)
    On Error Resume Next
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    On Error GoTo 0
    If loMainHF Is Nothing Then
        Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes)
        loMainHF.Name = "HFTable"
    End If
    
    ' Ensure the pasted SharePoint data is a ListObject
    On Error Resume Next
    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    On Error GoTo 0
    If loMainSP Is Nothing Then
        Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes)
        loMainSP.Name = "SharePoint"
    End If
    
    ' === 4. Apply filters on the HFTable in "Source Population" ============
    ' Clear any existing filters
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData
    
    ' 4.1 Filter IRR_Transparency_Tier to keep only "1" and "2"
    colIndex = GetColumnIndex(loMainHF, "IRR_Transparency_Tier")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
            Criteria1:=Array("1", "2"), Operator:=xlFilterValues
    End If
    
    ' 4.2 Filter HFAD_Strategy to remove "FIF", "Fund of Funds" and "Sub/Sleeve- No Benchmark"
    '      but include blanks if they exist.
    colIndex = GetColumnIndex(loMainHF, "HFAD_Strategy")
    If colIndex > 0 Then
        Dim allowedStrategy As Variant
        allowedStrategy = GetAllowedValues(loMainHF, "HFAD_Strategy", _
                            Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
        ' Include blank values if not already present.
        If IsError(Application.Match("", allowedStrategy, 0)) Then
            allowedStrategy = AppendToArray(allowedStrategy, "")
        End If
        If Not IsEmpty(allowedStrategy) Then
            loMainHF.Range.AutoFilter Field:=colIndex, _
                Criteria1:=allowedStrategy, Operator:=xlFilterValues
        End If
    End If
    
    ' 4.3 Filter HFAD_Entity_type to remove unwanted values
    '      but include blanks if they exist.
    colIndex = GetColumnIndex(loMainHF, "HFAD_Entity_type")
    If colIndex > 0 Then
        Dim allowedEntity As Variant
        allowedEntity = GetAllowedValues(loMainHF, "HFAD_Entity_type", _
                            Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                                  "Managed Account", "Managed Account - No AF", _
                                  "Loan Monitoring", "Loan FiF - No tracking", _
                                  "Sleeve/share class/sub-account"))
        ' Include blank values if not already present.
        If IsError(Application.Match("", allowedEntity, 0)) Then
            allowedEntity = AppendToArray(allowedEntity, "")
        End If
        If Not IsEmpty(allowedEntity) Then
            loMainHF.Range.AutoFilter Field:=colIndex, _
                Criteria1:=allowedEntity, Operator:=xlFilterValues
        End If
    End If
    
    ' 4.4 Filter IRR_last_update_date to keep only dates from 2023 and later.
    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
            Criteria1:=">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), Operator:=xlAnd
    End If
    
    ' === 5. Identify new funds (present in HFTable but missing in SharePoint) ============
    ' Build a dictionary of SharePoint HFAD_Fund_CoperID values, standardizing the keys.
    Set dictSP = CreateObject("Scripting.Dictionary")
    dictSP.CompareMode = vbTextCompare  ' Ensure case-insensitive comparisons.
    
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        Dim dataSP As Variant, rIndex As Long, key As String
        dataSP = loMainSP.DataBodyRange.Value
        For rIndex = 1 To UBound(dataSP, 1)
            key = Trim(CStr(dataSP(rIndex, colIndex)))
            If Not dictSP.Exists(key) Then
                dictSP.Add key, True
            End If
        Next rIndex
    End If
    
    ' Loop through the visible rows in HFTable and collect records where the HFAD_Fund_CoperID
    ' is not found in the SharePoint dictionary.
    Set newFunds = New Collection
    colIndex = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        On Error Resume Next
        Set visData = loMainHF.DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not visData Is Nothing Then
            Dim idxFundName As Long, idxIMCoperID As Long, idxIMName As Long, idxCreditOfficer As Long, idxTier As Long
            idxFundName = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
            idxIMCoperID = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
            idxIMName = GetColumnIndex(loMainHF, "HFAD_IM_Name")
            idxCreditOfficer = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
            idxTier = GetColumnIndex(loMainHF, "IRR_Transparency_Tier")
            
            For Each r In visData.Rows
                If r.EntireRow.Hidden = False Then
                    fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array( _
                            fundCoperID, _
                            r.Cells(1, idxFundName).Value, _
                            r.Cells(1, idxIMCoperID).Value, _
                            r.Cells(1, idxIMName).Value, _
                            r.Cells(1, idxCreditOfficer).Value, _
                            r.Cells(1, idxTier).Value, _
                            "Active")
                        newFunds.Add rec
                    End If
                End If
            Next r
        End If
    End If
    
    ' === 6. Create new sheet "Upload to SP" with table "UploadHF" ============
    On Error Resume Next
    Set wsUpload = wbMain.Sheets("Upload to SP")
    On Error GoTo 0
    If wsUpload Is Nothing Then
        Set wsUpload = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsUpload.Name = "Upload to SP"
    Else
        wsUpload.Cells.Clear
    End If
    
    ' Create initial headers for UploadHF table.
    headers = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", _
                    "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
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
    On Error GoTo 0
    If loUpload Is Nothing Then
        Set loUpload = wsUpload.ListObjects.Add(xlSrcRange, rngUpload, , xlYes)
        loUpload.Name = "UploadHF"
    Else
        loUpload.Resize rngUpload
    End If
    
    ' === 7. Additional Enhancements: Update UploadHF with extra lookup columns ============
    ' (a) & (b): Using CO_Table from the "CO_Table" sheet to add "Region" and update "HFAD_Credit_Officer"
    Dim wsCO As Worksheet, loCO As ListObject
    Dim coDict As Object ' Dictionary: key = Credit Officer, value = Array(Region, Email Address)
    Set coDict = CreateObject("Scripting.Dictionary")
    coDict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Set wsCO = wbMain.Sheets("CO_Table")
    On Error GoTo 0
    If wsCO Is Nothing Then
        MsgBox "CO_Table sheet not found.", vbCritical
        Exit Sub
    End If
    On Error Resume Next
    Set loCO = wsCO.ListObjects("CO_Table")
    On Error GoTo 0
    If loCO Is Nothing Then
        MsgBox "CO_Table table not found on CO_Table sheet.", vbCritical
        Exit Sub
    End If
    
    Dim coCredCol As Long, coRegionCol As Long, coEmailCol As Long
    coCredCol = GetColumnIndex(loCO, "Credit Officer")
    coRegionCol = GetColumnIndex(loCO, "Region")
    coEmailCol = GetColumnIndex(loCO, "Email Address")
    If coCredCol = 0 Or coRegionCol = 0 Or coEmailCol = 0 Then
        MsgBox "Required columns not found in CO_Table.", vbCritical
        Exit Sub
    End If
    
    Dim coData As Variant, rIdx As Long, coKey As String
    coData = loCO.DataBodyRange.Value
    For rIdx = 1 To UBound(coData, 1)
        coKey = Trim(CStr(coData(rIdx, coCredCol)))
        If Not coDict.Exists(coKey) Then
            coDict.Add coKey, Array(coData(rIdx, coRegionCol), coData(rIdx, coEmailCol))
        End If
    Next rIdx
    
    ' (c) Build a dictionary from the SharePoint table for IM info using HFAD_IM_CoperID.
    Dim imDict As Object
    Set imDict = CreateObject("Scripting.Dictionary")
    imDict.CompareMode = vbTextCompare
    
    Dim spData As Variant, imKey As String
    Dim sp_IMCol As Long, sp_NAVCol As Long, sp_FreqCol As Long, sp_AdHocCol As Long, sp_ParentFlagCol As Long
    sp_IMCol = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol = GetColumnIndex(loMainSP, "Ad Hoc Reporting")
    sp_ParentFlagCol = GetColumnIndex(loMainSP, "Parent/Flagship Reporting")
    If sp_IMCol = 0 Or sp_NAVCol = 0 Or sp_FreqCol = 0 Or sp_AdHocCol = 0 Or sp_ParentFlagCol = 0 Then
        MsgBox "One or more required columns not found in SharePoint table.", vbCritical
        Exit Sub
    End If
    
    spData = loMainSP.DataBodyRange.Value
    For rIdx = 1 To UBound(spData, 1)
        imKey = Trim(CStr(spData(rIdx, sp_IMCol)))
        If Not imDict.Exists(imKey) Then
            imDict.Add imKey, Array(spData(rIdx, sp_NAVCol), spData(rIdx, sp_FreqCol), spData(rIdx, sp_AdHocCol), spData(rIdx, sp_ParentFlagCol))
        End If
    Next rIdx
    
    ' (e) Build a dictionary from the HFTable for "Days to Report" using HFAD_Fund_CoperID.
    Dim daysDict As Object
    Set daysDict = CreateObject("Scripting.Dictionary")
    daysDict.CompareMode = vbTextCompare
    
    Dim hfData As Variant
    Dim hfFundIDCol As Long, hfDaysCol As Long
    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol = GetColumnIndex(loMainHF, "HFAD_days_to_report")
    If hfFundIDCol = 0 Or hfDaysCol = 0 Then
        MsgBox "Required columns not found in HFTable for Days to Report.", vbCritical
        Exit Sub
    End If
    
    hfData = loMainHF.DataBodyRange.Value
    For rIdx = 1 To UBound(hfData, 1)
        Dim fundKey As String
        fundKey = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not daysDict.Exists(fundKey) Then
            daysDict.Add fundKey, hfData(rIdx, hfDaysCol)
        End If
    Next rIdx
    
    ' Add new columns to UploadHF if they do not already exist.
    If Not ColumnExists(loUpload, "Region") Then loUpload.ListColumns.Add.Name = "Region"
    If Not ColumnExists(loUpload, "NAV Source") Then loUpload.ListColumns.Add.Name = "NAV Source"
    If Not ColumnExists(loUpload, "Frequency") Then loUpload.ListColumns.Add.Name = "Frequency"
    If Not ColumnExists(loUpload, "Ad Hoc Reporting") Then loUpload.ListColumns.Add.Name = "Ad Hoc Reporting"
    If Not ColumnExists(loUpload, "Parent/Flagship Reporting") Then loUpload.ListColumns.Add.Name = "Parent/Flagship Reporting"
    If Not ColumnExists(loUpload, "Days to Report") Then loUpload.ListColumns.Add.Name = "Days to Report"
    
    ' Get the column indices in UploadHF for the lookup updates.
    Dim up_CreditOfficerCol As Long, up_RegionCol As Long, up_IMCoperIDCol As Long, up_NAVSourceCol As Long
    Dim up_FrequencyCol As Long, up_AdHocCol As Long, up_ParentFlagCol As Long, up_FundCoperIDCol As Long, up_DaysToReportCol As Long
    up_CreditOfficerCol = GetColumnIndex(loUpload, "HFAD_Credit_Officer")
    up_RegionCol = GetColumnIndex(loUpload, "Region")
    up_IMCoperIDCol = GetColumnIndex(loUpload, "HFAD_IM_CoperID")
    up_NAVSourceCol = GetColumnIndex(loUpload, "NAV Source")
    up_FrequencyCol = GetColumnIndex(loUpload, "Frequency")
    up_AdHocCol = GetColumnIndex(loUpload, "Ad Hoc Reporting")
    up_ParentFlagCol = GetColumnIndex(loUpload, "Parent/Flagship Reporting")
    up_FundCoperIDCol = GetColumnIndex(loUpload, "HFAD_Fund_CoperID")
    up_DaysToReportCol = GetColumnIndex(loUpload, "Days to Report")
    
    ' Loop through each row in UploadHF and update with lookup values.
    Dim upRow As ListRow
    For Each upRow In loUpload.ListRows
        Dim creditOfficerName As String, imCoperID As String, fundCoperID As String
        creditOfficerName = Trim(CStr(upRow.Range.Cells(1, up_CreditOfficerCol).Value))
        imCoperID = Trim(CStr(upRow.Range.Cells(1, up_IMCoperIDCol).Value))
        fundCoperID = Trim(CStr(upRow.Range.Cells(1, up_FundCoperIDCol).Value))
        
        ' (a) & (b): Lookup in CO_Table dictionary using the credit officer name.
        If coDict.Exists(creditOfficerName) Then
            Dim coInfo As Variant
            coInfo = coDict(creditOfficerName)
            ' coInfo(0) = Region, coInfo(1) = Email Address.
            upRow.Range.Cells(1, up_CreditOfficerCol).Value = coInfo(1)
            upRow.Range.Cells(1, up_RegionCol).Value = coInfo(0)
        Else
            upRow.Range.Cells(1, up_RegionCol).Value = ""
        End If
        
        ' (c) & (d): Lookup in SharePoint dictionary using HFAD_IM_CoperID.
        If imDict.Exists(imCoperID) Then
            Dim imInfo As Variant
            imInfo = imDict(imCoperID)
            upRow.Range.Cells(1, up_NAVSourceCol).Value = imInfo(0)
            upRow.Range.Cells(1, up_FrequencyCol).Value = imInfo(1)
            upRow.Range.Cells(1, up_AdHocCol).Value = imInfo(2)
            upRow.Range.Cells(1, up_ParentFlagCol).Value = imInfo(3)
        Else
            upRow.Range.Cells(1, up_NAVSourceCol).Value = "Client Email"
            upRow.Range.Cells(1, up_FrequencyCol).Value = "Monthly"
            upRow.Range.Cells(1, up_AdHocCol).Value = "No"
            upRow.Range.Cells(1, up_ParentFlagCol).Value = "No"
        End If
        
        ' (e) Lookup in HFTable dictionary for Days to Report using HFAD_Fund_CoperID.
        If daysDict.Exists(fundCoperID) Then
            upRow.Range.Cells(1, up_DaysToReportCol).Value = daysDict(fundCoperID)
        Else
            upRow.Range.Cells(1, up_DaysToReportCol).Value = ""
        End If
    Next upRow
    
    MsgBox "Macro completed successfully.", vbInformation
End Sub

' -------------------------------
' Helper function: Returns the relative column index (within the ListObject)
' for a given header name. Returns 0 if the header is not found.
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

' -------------------------------
' Helper function: Returns an array of unique allowed values from a given field
' in the ListObject that are NOT in the provided exclusion list.
Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIndex As Long
    colIndex = GetColumnIndex(lo, fieldName)
    If colIndex = 0 Then
        GetAllowedValues = Array() ' Return an empty array if the header is not found.
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim cell As Range, rng As Range
    Dim cellVal As Variant, i As Long, skipVal As Boolean
    
    Set rng = lo.ListColumns(fieldName).DataBodyRange
    For Each cell In rng
        cellVal = cell.Value
        skipVal = False
        For i = LBound(excludeArr) To UBound(excludeArr)
            If Trim(CStr(cellVal)) = Trim(CStr(excludeArr(i))) Then
                skipVal = True
                Exit For
            End If
        Next i
        If Not skipVal Then
            If Not dict.Exists(cellVal) Then
                dict.Add cellVal, cellVal
            End If
        End If
    Next cell
    
    If dict.Count > 0 Then
        GetAllowedValues = dict.Keys
    Else
        GetAllowedValues = Array()
    End If
End Function

' -------------------------------
' Helper function: Appends a value to an existing array.
Function AppendToArray(arr As Variant, valueToAppend As Variant) As Variant
    Dim newArr() As Variant
    Dim i As Long, n As Long
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

' -------------------------------
' Helper function: Checks if a column with the given name exists in the ListObject.
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
