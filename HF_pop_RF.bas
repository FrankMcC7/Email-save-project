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
    Dim inactiveRow As ListRow ' used for InactiveHF table loops
    
    '=======================
    ' 1. Define file paths (hardcoded)
    HFFilePath = "C:\YourFolder\HFFile.xlsx"           ' <<< Change to your actual path
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"     ' <<< Change to your actual path
    Set wbMain = ThisWorkbook
    
    '=======================
    ' 2. Open the HF file and convert its data to table "HFTable"
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1)  ' Assumes data is in the first sheet
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set rngHF = wsHFSource.UsedRange
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, rngHF, , xlYes)
    End If
    loHF.Name = "HFTable"
    
    '------------------------------------------------------------
    ' 2. Open the SharePoint file and convert its data to table "SharePoint"
    Set wbSP = Workbooks.Open(SPFilePath)
    Set wsSPSource = wbSP.Sheets(1)
    If wsSPSource.ListObjects.Count > 0 Then
        Set loSP = wsSPSource.ListObjects(1)
    Else
        Set rngSP = wsSPSource.UsedRange
        Set loSP = wsSPSource.ListObjects.Add(xlSrcRange, rngSP, , xlYes)
    End If
    loSP.Name = "SharePoint"
    
    '------------------------------------------------------------
    ' 3. Paste the tables into the main workbook
    On Error Resume Next
    Set wsSourcePop = wbMain.Sheets("Source Population")
    On Error GoTo 0
    If wsSourcePop Is Nothing Then
        Set wsSourcePop = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSourcePop.Name = "Source Population"
    Else
        wsSourcePop.Cells.Clear
    End If
    
    On Error Resume Next
    Set wsSPMain = wbMain.Sheets("SharePoint")
    On Error GoTo 0
    If wsSPMain Is Nothing Then
        Set wsSPMain = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsSPMain.Name = "SharePoint"
    Else
        wsSPMain.Cells.Clear
    End If
    
    ' Copy the tables into the respective sheets
    loHF.Range.Copy Destination:=wsSourcePop.Range("A1")
    loSP.Range.Copy Destination:=wsSPMain.Range("A1")
    
    wbHF.Close SaveChanges:=False
    wbSP.Close SaveChanges:=False
    
    ' Ensure the pasted HF data is a table named "HFTable"
    On Error Resume Next
    Set loMainHF = wsSourcePop.ListObjects("HFTable")
    On Error GoTo 0
    If loMainHF Is Nothing Then
        Set loMainHF = wsSourcePop.ListObjects.Add(xlSrcRange, wsSourcePop.UsedRange, , xlYes)
        loMainHF.Name = "HFTable"
    End If
    
    ' Ensure the pasted SharePoint data is a table named "SharePoint"
    On Error Resume Next
    Set loMainSP = wsSPMain.ListObjects("SharePoint")
    On Error GoTo 0
    If loMainSP Is Nothing Then
        Set loMainSP = wsSPMain.ListObjects.Add(xlSrcRange, wsSPMain.UsedRange, , xlYes)
        loMainSP.Name = "SharePoint"
    End If
    
    '=======================
    ' 4. Apply filters on the HFTable in "Source Population"
    If loMainHF.AutoFilter.FilterMode Then loMainHF.AutoFilter.ShowAllData
    
    ' 4.1 Filter IRR_Transparency_Tier to keep only "1" and "2"
    colIndex = GetColumnIndex(loMainHF, "IRR_Transparency_Tier")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
            Criteria1:=Array("1", "2"), Operator:=xlFilterValues
    End If
    
    ' 4.2 Filter HFAD_Strategy to remove unwanted values but include blanks
    colIndex = GetColumnIndex(loMainHF, "HFAD_Strategy")
    If colIndex > 0 Then
        Dim allowedStrategy As Variant
        allowedStrategy = GetAllowedValues(loMainHF, "HFAD_Strategy", _
                           Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
        If IsError(Application.Match("", allowedStrategy, 0)) Then
            allowedStrategy = AppendToArray(allowedStrategy, "")
        End If
        If Not IsEmpty(allowedStrategy) Then
            loMainHF.Range.AutoFilter Field:=colIndex, _
                Criteria1:=allowedStrategy, Operator:=xlFilterValues
        End If
    End If
    
    ' 4.3 Filter HFAD_Entity_type to remove unwanted values but include blanks
    colIndex = GetColumnIndex(loMainHF, "HFAD_Entity_type")
    If colIndex > 0 Then
        Dim allowedEntity As Variant
        allowedEntity = GetAllowedValues(loMainHF, "HFAD_Entity_type", _
                           Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                           "Managed Account", "Managed Account - No AF", "Loan Monitoring", _
                           "Loan FiF - No tracking", "Sleeve/share class/sub-account"))
        If IsError(Application.Match("", allowedEntity, 0)) Then
            allowedEntity = AppendToArray(allowedEntity, "")
        End If
        If Not IsEmpty(allowedEntity) Then
            loMainHF.Range.AutoFilter Field:=colIndex, _
                Criteria1:=allowedEntity, Operator:=xlFilterValues
        End If
    End If
    
    ' 4.4 Filter IRR_last_update_date to keep only dates from 2023 and later
    colIndex = GetColumnIndex(loMainHF, "IRR_last_update_date")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
            Criteria1:=">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), Operator:=xlAnd
    End If
    
    '=======================
    ' 5. Identify new funds: HFAD_Fund_CoperID values in HFTable not present in SharePoint
    Set dictSP = CreateObject("Scripting.Dictionary")
    dictSP.CompareMode = vbTextCompare
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        Dim dataSP As Variant
        dataSP = loMainSP.DataBodyRange.Value
        For iSP = 1 To UBound(dataSP, 1)
            key = Trim(CStr(dataSP(iSP, colIndex)))
            If Not dictSP.Exists(key) Then
                dictSP.Add key, True
            End If
        Next iSP
    End If
    
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
                        rec = Array(fundCoperID, _
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
    
    '=======================
    ' 6. Create new sheet "Upload to SP" with table "UploadHF"
    On Error Resume Next
    Set wsUpload = wbMain.Sheets("Upload to SP")
    On Error GoTo 0
    If wsUpload Is Nothing Then
        Set wsUpload = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsUpload.Name = "Upload to SP"
    Else
        wsUpload.Cells.Clear
    End If
    
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
    
    '=======================
    ' 7. Additional Enhancements for UploadHF (Credit Officer, IM info, Days to Report)
    ' (a) & (b): Update HFAD_Credit_Officer and add Region from CO_Table
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
    
    coCredCol = GetColumnIndex(loCO, "Credit Officer")
    coRegionCol = GetColumnIndex(loCO, "Region")
    coEmailCol = GetColumnIndex(loCO, "Email Address")
    If coCredCol = 0 Or coRegionCol = 0 Or coEmailCol = 0 Then
        MsgBox "Required columns not found in CO_Table.", vbCritical
        Exit Sub
    End If
    
    coData = loCO.DataBodyRange.Value
    For rIdx = 1 To UBound(coData, 1)
        coKey = Trim(CStr(coData(rIdx, coCredCol)))
        If Not coDict.Exists(coKey) Then
            coDict.Add coKey, Array(coData(rIdx, coRegionCol), coData(rIdx, coEmailCol))
        End If
    Next rIdx
    
    ' (c) & (d): Update NAV Source, Frequency, Ad-Hoc Reporting, Parent/Flagship Reporting from SharePoint
    Set imDict = CreateObject("Scripting.Dictionary")
    imDict.CompareMode = vbTextCompare
    sp_IMCol = GetColumnIndex(loMainSP, "HFAD_IM_CoperID")
    sp_NAVCol = GetColumnIndex(loMainSP, "NAV Source")
    sp_FreqCol = GetColumnIndex(loMainSP, "Frequency")
    sp_AdHocCol = GetColumnIndex(loMainSP, "Ad-Hoc Reporting")
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
    
    ' (e): Update Days to Report from HFTable using HFAD_Days_to_report column
    Set daysDict = CreateObject("Scripting.Dictionary")
    daysDict.CompareMode = vbTextCompare
    hfFundIDCol = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    hfDaysCol = GetColumnIndex(loMainHF, "HFAD_Days_to_report")
    If hfFundIDCol = 0 Or hfDaysCol = 0 Then
        MsgBox "Required columns not found in HFTable for Days to Report.", vbCritical
        Exit Sub
    End If
    
    hfData = loMainHF.DataBodyRange.Value
    For rIdx = 1 To UBound(hfData, 1)
        fundKey = Trim(CStr(hfData(rIdx, hfFundIDCol)))
        If Not daysDict.Exists(fundKey) Then
            daysDict.Add fundKey, hfData(rIdx, hfDaysCol)
        End If
    Next rIdx
    
    ' Add new columns to UploadHF if they do not already exist.
    If Not ColumnExists(loUpload, "Region") Then loUpload.ListColumns.Add.Name = "Region"
    If Not ColumnExists(loUpload, "NAV Source") Then loUpload.ListColumns.Add.Name = "NAV Source"
    If Not ColumnExists(loUpload, "Frequency") Then loUpload.ListColumns.Add.Name = "Frequency"
    If Not ColumnExists(loUpload, "Ad-Hoc Reporting") Then loUpload.ListColumns.Add.Name = "Ad-Hoc Reporting"
    If Not ColumnExists(loUpload, "Parent/Flagship Reporting") Then loUpload.ListColumns.Add.Name = "Parent/Flagship Reporting"
    If Not ColumnExists(loUpload, "Days to Report") Then loUpload.ListColumns.Add.Name = "Days to Report"
    
    up_CreditOfficerCol = GetColumnIndex(loUpload, "HFAD_Credit_Officer")
    up_RegionCol = GetColumnIndex(loUpload, "Region")
    up_IMCoperIDCol = GetColumnIndex(loUpload, "HFAD_IM_CoperID")
    up_NAVSourceCol = GetColumnIndex(loUpload, "NAV Source")
    up_FrequencyCol = GetColumnIndex(loUpload, "Frequency")
    up_AdHocCol = GetColumnIndex(loUpload, "Ad-Hoc Reporting")
    up_ParentFlagCol = GetColumnIndex(loUpload, "Parent/Flagship Reporting")
    up_FundCoperIDCol = GetColumnIndex(loUpload, "HFAD_Fund_CoperID")
    up_DaysToReportCol = GetColumnIndex(loUpload, "Days to Report")
    
    For Each upRow In loUpload.ListRows
        creditOfficerName = Trim(CStr(upRow.Range.Cells(1, up_CreditOfficerCol).Value))
        imCoperID = Trim(CStr(upRow.Range.Cells(1, up_IMCoperIDCol).Value))
        fundCoperID = Trim(CStr(upRow.Range.Cells(1, up_FundCoperIDCol).Value))
        
        ' (a) & (b): Lookup in CO_Table dictionary using the Credit Officer name.
        If coDict.Exists(creditOfficerName) Then
            Dim coInfo As Variant
            coInfo = coDict(creditOfficerName)
            upRow.Range.Cells(1, up_CreditOfficerCol).Value = coInfo(1)  ' Replace with Email Address
            upRow.Range.Cells(1, up_RegionCol).Value = coInfo(0)          ' Populate Region
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
        
        ' (e): Lookup Days to Report in HFTable dictionary using HFAD_Fund_CoperID.
        If daysDict.Exists(fundCoperID) Then
            upRow.Range.Cells(1, up_DaysToReportCol).Value = daysDict(fundCoperID)
        Else
            upRow.Range.Cells(1, up_DaysToReportCol).Value = ""
        End If
    Next upRow
    
    '=======================
    ' 8. Inactive Funds Identification and Processing
    ' (1) Build dictionaries from HFTable: one for all fund IDs and one for Tier info.
    Dim dictHF As Object, tierDict As Object
    Set dictHF = CreateObject("Scripting.Dictionary")
    dictHF.CompareMode = vbTextCompare
    Set tierDict = CreateObject("Scripting.Dictionary")
    tierDict.CompareMode = vbTextCompare
    
    Dim hfIDColForDict As Long
    hfIDColForDict = GetColumnIndex(loMainHF, "HFAD_Fund_CoperID")
    Dim tierCol As Long
    ' Use "IRR_Tranparency_Tier" as per instructions; if not found, try "IRR_Transparency_Tier"
    tierCol = GetColumnIndex(loMainHF, "IRR_Tranparency_Tier")
    If tierCol = 0 Then
        tierCol = GetColumnIndex(loMainHF, "IRR_Transparency_Tier")
    End If
    
    If hfIDColForDict > 0 Then
        Dim arrHF As Variant
        arrHF = loMainHF.DataBodyRange.Value
        Dim rHF As Long
        For rHF = 1 To UBound(arrHF, 1)
            Dim idVal As String
            idVal = Trim(CStr(arrHF(rHF, hfIDColForDict)))
            If Not dictHF.Exists(idVal) Then
                dictHF.Add idVal, True
            End If
            If tierCol > 0 Then
                Dim tierVal As Variant
                tierVal = arrHF(rHF, tierCol)
                If Not tierDict.Exists(idVal) Then
                    tierDict.Add idVal, tierVal
                End If
            End If
        Next rHF
    End If
    
    ' (2) Build a collection of inactive funds from SharePoint: funds present in SharePoint but not in HFTable.
    Dim inactiveFunds As Collection
    Set inactiveFunds = New Collection
    Dim share_CoperCol As Long, share_StatusCol As Long, share_CommentsCol As Long
    share_CoperCol = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    share_StatusCol = GetColumnIndex(loMainSP, "Status")
    share_CommentsCol = GetColumnIndex(loMainSP, "Comments")
    
    Dim arrSPInactive As Variant
    arrSPInactive = loMainSP.DataBodyRange.Value
    Dim rSP As Long
    For rSP = 1 To UBound(arrSPInactive, 1)
        Dim spFundID As String
        spFundID = Trim(CStr(arrSPInactive(rSP, share_CoperCol)))
        If Not dictHF.Exists(spFundID) Then
            Dim fundStatus As String, fundComments As String, fundTier As Variant
            fundStatus = arrSPInactive(rSP, share_StatusCol)
            fundComments = arrSPInactive(rSP, share_CommentsCol)
            If tierDict.Exists(spFundID) Then
                fundTier = tierDict(spFundID)
            Else
                fundTier = ""
            End If
            Dim inactRec As Variant
            inactRec = Array(spFundID, fundStatus, fundComments, fundTier)
            inactiveFunds.Add inactRec
        End If
    Next rSP
    
    ' (2b) Create "Inactive Funds Tracking" sheet and table "InactiveHF"
    Dim wsInactive As Worksheet
    Dim loInactive As ListObject
    On Error Resume Next
    Set wsInactive = wbMain.Sheets("Inactive Funds Tracking")
    On Error GoTo 0
    If wsInactive Is Nothing Then
        Set wsInactive = wbMain.Sheets.Add(After:=wbMain.Sheets(wbMain.Sheets.Count))
        wsInactive.Name = "Inactive Funds Tracking"
    Else
        wsInactive.Cells.Clear
    End If
    
    Dim inactHeaders As Variant
    inactHeaders = Array("HFAD_Fund_CoperID", "Status", "Comments", "Tier")
    Dim colCounter As Long
    For colCounter = LBound(inactHeaders) To UBound(inactHeaders)
        wsInactive.Cells(1, colCounter + 1).Value = inactHeaders(colCounter)
    Next colCounter
    
    Dim inactRowCounter As Long
    inactRowCounter = 2
    For i = 1 To inactiveFunds.Count
        Dim recInactive As Variant
        recInactive = inactiveFunds(i)
        For j = LBound(recInactive) To UBound(recInactive)
            wsInactive.Cells(inactRowCounter, j + 1).Value = recInactive(j)
        Next j
        inactRowCounter = inactRowCounter + 1
    Next i
    
    Dim rngInactive As Range
    Set rngInactive = wsInactive.Range(wsInactive.Cells(1, 1), wsInactive.Cells(inactRowCounter - 1, UBound(inactHeaders) + 1))
    On Error Resume Next
    Set loInactive = wsInactive.ListObjects("InactiveHF")
    On Error GoTo 0
    If loInactive Is Nothing Then
        Set loInactive = wsInactive.ListObjects.Add(xlSrcRange, rngInactive, , xlYes)
        loInactive.Name = "InactiveHF"
    Else
        loInactive.Resize rngInactive
    End If
    
    ' (3) Delete all rows of InactiveHF where the value of "Status" is "Inactive"
    Dim statusColInactive As Long
    statusColInactive = GetColumnIndex(loInactive, "Status")
    If statusColInactive > 0 Then
        Dim k As Long
        For k = loInactive.ListRows.Count To 1 Step -1
            Dim rowStatus As String
            rowStatus = Trim(CStr(loInactive.ListRows(k).Range.Cells(1, statusColInactive).Value))
            If StrComp(rowStatus, "Inactive", vbTextCompare) = 0 Then
                loInactive.ListRows(k).Delete
            End If
        Next k
    End If
    
    ' (4) Prompt the user to locate a Database file; open it, convert its data to table "DBTable",
    ' then add a new column "HF Status" in InactiveHF by looking up the DBTable.
    Dim dbFilePath As String
    dbFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select Database File")
    If dbFilePath = "False" Then
        MsgBox "No database file selected. Skipping HF Status update.", vbExclamation
    Else
        Dim wbDB As Workbook, wsDB As Worksheet, loDB As ListObject
        Set wbDB = Workbooks.Open(dbFilePath)
        Set wsDB = wbDB.Sheets(1) ' assume data is in first sheet
        On Error Resume Next
        Set loDB = wsDB.ListObjects("DBTable")
        On Error GoTo 0
        If loDB Is Nothing Then
            Dim rngDB As Range
            Set rngDB = wsDB.UsedRange
            Set loDB = wsDB.ListObjects.Add(xlSrcRange, rngDB, , xlYes)
            loDB.Name = "DBTable"
        End If
        
        ' Build dictionary from DBTable: key = Fund CoPer, value = value from "Active" column.
        Dim dbDict As Object
        Set dbDict = CreateObject("Scripting.Dictionary")
        dbDict.CompareMode = vbTextCompare
        Dim dbFundCol As Long, dbActiveCol As Long
        dbFundCol = GetColumnIndex(loDB, "Fund CoPer")
        dbActiveCol = GetColumnIndex(loDB, "Active")
        If dbFundCol = 0 Or dbActiveCol = 0 Then
            MsgBox "Required columns not found in DBTable.", vbCritical
            wbDB.Close SaveChanges:=False
            GoTo SkipDBProcessing
        End If
        
        Dim arrDB As Variant
        arrDB = loDB.DataBodyRange.Value
        Dim rDB As Long
        For rDB = 1 To UBound(arrDB, 1)
            Dim dbKey As String
            dbKey = Trim(CStr(arrDB(rDB, dbFundCol)))
            If Not dbDict.Exists(dbKey) Then
                dbDict.Add dbKey, arrDB(rDB, dbActiveCol)
            End If
        Next rDB
        
        wbDB.Close SaveChanges:=False
        
        ' Add column "HF Status" to InactiveHF if not exists.
        If Not ColumnExists(loInactive, "HF Status") Then loInactive.ListColumns.Add.Name = "HF Status"
        Dim hfStatusCol As Long
        hfStatusCol = GetColumnIndex(loInactive, "HF Status")
        
        For Each inactiveRow In loInactive.ListRows
            Dim inactiveFundID As String
            inactiveFundID = Trim(CStr(inactiveRow.Range.Cells(1, GetColumnIndex(loInactive, "HFAD_Fund_CoperID")).Value))
            If dbDict.Exists(inactiveFundID) Then
                Dim dbValue As String
                dbValue = Trim(CStr(dbDict(inactiveFundID)))
                If StrComp(dbValue, "N", vbTextCompare) = 0 Then
                    inactiveRow.Range.Cells(1, hfStatusCol).Value = "Inactive"
                ElseIf StrComp(dbValue, "Y", vbTextCompare) = 0 Then
                    inactiveRow.Range.Cells(1, hfStatusCol).Value = "Active"
                Else
                    inactiveRow.Range.Cells(1, hfStatusCol).Value = dbValue
                End If
            Else
                inactiveRow.Range.Cells(1, hfStatusCol).Value = ""
            End If
        Next inactiveRow
    End If
SkipDBProcessing:
    
    ' (5) Create column "Review" in InactiveHF: if "Status" and "HF Status" differ then "Check", else blank.
    If Not ColumnExists(loInactive, "Review") Then loInactive.ListColumns.Add.Name = "Review"
    Dim reviewCol As Long
    reviewCol = GetColumnIndex(loInactive, "Review")
    Dim statusCol As Long
    statusCol = GetColumnIndex(loInactive, "Status")
    For Each inactiveRow In loInactive.ListRows
        Dim statVal As String, hfStatVal As String
        statVal = Trim(CStr(inactiveRow.Range.Cells(1, statusCol).Value))
        hfStatVal = Trim(CStr(inactiveRow.Range.Cells(1, hfStatusCol).Value))
        If StrComp(statVal, hfStatVal, vbTextCompare) = 0 Then
            inactiveRow.Range.Cells(1, reviewCol).Value = ""
        Else
            inactiveRow.Range.Cells(1, reviewCol).Value = "Check"
        End If
    Next inactiveRow
    
    '=======================
    MsgBox "Macro completed successfully.", vbInformation
End Sub

'------------------------------------------------
' Helper function: Returns the column index (within the ListObject) for a given header.
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

'------------------------------------------------
' Helper function: Returns an array of unique allowed values from a table column excluding specified items.
Function GetAllowedValues(lo As ListObject, fieldName As String, excludeArr As Variant) As Variant
    Dim colIndex As Long
    colIndex = GetColumnIndex(lo, fieldName)
    If colIndex = 0 Then
        GetAllowedValues = Array()
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

'------------------------------------------------
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

'------------------------------------------------
' Helper function: Checks if a column with the given name exists in a ListObject.
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
