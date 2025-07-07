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
    
    ' 4.1 Filter IRR_Scorecard_factor to keep only "Transparency"
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, Criteria1:="Transparency"
    End If
    
    ' 4.2 Filter IRR_Scorecard_factor_value to keep only values from 2023 and later
    colIndex = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
    If colIndex > 0 Then
        loMainHF.Range.AutoFilter Field:=colIndex, _
            Criteria1:= ">=" & Format(DateSerial(2023, 1, 1), "mm/dd/yyyy"), _
            Operator:=xlAnd
    End If
    
    ' 4.3 Filter HFAD_Strategy to remove unwanted values but include blanks
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
    
    ' 4.4 Filter HFAD_Entity_type to remove unwanted values but include blanks
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
            Dim idxFactorVal As Long
            idxFactorVal = GetColumnIndex(loMainHF, "IRR_Scorecard_factor_value")
            Dim idxFundName As Long, idxIMCoperID As Long, idxIMName As Long, idxCreditOfficer As Long
            idxFundName = GetColumnIndex(loMainHF, "HFAD_Fund_Name")
            idxIMCoperID = GetColumnIndex(loMainHF, "HFAD_IM_CoperID")
            idxIMName = GetColumnIndex(loMainHF, "HFAD_IM_Name")
            idxCreditOfficer = GetColumnIndex(loMainHF, "HFAD_Credit_Officer")
            For Each r In visData.Rows
                If Not r.EntireRow.Hidden Then
                    fundCoperID = Trim(CStr(r.Cells(1, colIndex).Value))
                    If Not dictSP.Exists(fundCoperID) Then
                        rec = Array(fundCoperID, _
                                    r.Cells(1, idxFundName).Value, _
                                    r.Cells(1, idxIMCoperID).Value, _
                                    r.Cells(1, idxIMName).
