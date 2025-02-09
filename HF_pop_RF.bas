Option Explicit

' Main procedure for “Phase 1 – New funds identification”
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
    Dim i As Long, j As Integer
    Dim fundCoperID As Variant
    Dim newFunds As Collection
    Dim rec As Variant
    Dim wsUpload As Worksheet
    Dim loUpload As ListObject
    Dim rngUpload As Range
    Dim headers As Variant
    Dim rowCounter As Long
    
    ' === 1. Define the file paths (hardcoded) ===
    HFFilePath = "C:\YourFolder\HFFile.xlsx"           ' <<< CHANGE to your actual path
    SPFilePath = "C:\YourFolder\SharePointFile.xlsx"     ' <<< CHANGE to your actual path
    
    ' Set main workbook (this workbook where the macro resides)
    Set wbMain = ThisWorkbook
    
    ' === 2. Open the HF file and convert data to table "HFTable" ===
    Set wbHF = Workbooks.Open(HFFilePath)
    Set wsHFSource = wbHF.Sheets(1) ' Assumes data is in the first sheet
    If wsHFSource.ListObjects.Count > 0 Then
        Set loHF = wsHFSource.ListObjects(1)
    Else
        Set rngHF = wsHFSource.UsedRange
        Set loHF = wsHFSource.ListObjects.Add(xlSrcRange, rngHF, , xlYes)
    End If
    loHF.Name = "HFTable"
    
    ' === 2. Open the SharePoint file and convert data to table "SharePoint" ===
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
    
    ' Copy HF data table from source file into "Source Population"
    loHF.Range.Copy Destination:=wsSourcePop.Range("A1")
    ' Copy SharePoint data table from source file into "SharePoint"
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
    
    ' (Optional) Ensure the SharePoint pasted data is a ListObject
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
    colIndex = GetColumnIndex(loMainHF, "HFAD_Strategy")
    If colIndex > 0 Then
        Dim allowedStrategy As Variant
        allowedStrategy = GetAllowedValues(loMainHF, "HFAD_Strategy", _
                            Array("FIF", "Fund of Funds", "Sub/Sleeve- No Benchmark"))
        ' Include blank values if not already present
        If IsError(Application.Match("", allowedStrategy, 0)) Then
            allowedStrategy = AppendToArray(allowedStrategy, "")
        End If
        If Not IsEmpty(allowedStrategy) Then
            loMainHF.Range.AutoFilter Field:=colIndex, _
                Criteria1:=allowedStrategy, Operator:=xlFilterValues
        End If
    End If
    
    ' 4.3 Filter HFAD_Entity_type to remove specific unwanted values
    colIndex = GetColumnIndex(loMainHF, "HFAD_Entity_type")
    If colIndex > 0 Then
        Dim allowedEntity As Variant
        allowedEntity = GetAllowedValues(loMainHF, "HFAD_Entity_type", _
                            Array("Guaranteed subsidiary", "Investment Manager as Agent", _
                                  "Managed Account", "Managed Account - No AF", _
                                  "Loan Monitoring", "Loan FiF - No tracking", _
                                  "Sleeve/share class/sub-account"))
        ' Include blank values if not already present
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
    
    ' === 5. Identify new funds (present in HFTable but missing in SharePoint) ============
    ' Build a dictionary of SharePoint HFAD_Fund_CoperID values
    Set dictSP = CreateObject("Scripting.Dictionary")
    colIndex = GetColumnIndex(loMainSP, "HFAD_Fund_CoperID")
    If colIndex > 0 Then
        Dim dataSP As Variant, rIndex As Long
        dataSP = loMainSP.DataBodyRange.Value
        For rIndex = 1 To UBound(dataSP, 1)
            If Not dictSP.Exists(dataSP(rIndex, colIndex)) Then
                dictSP.Add dataSP(rIndex, colIndex), True
            End If
        Next rIndex
    End If
    
    ' Loop through visible rows in HFTable and check HFAD_Fund_CoperID
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
                ' Check if the entire row is visible (it might be hidden by the filter)
                If r.EntireRow.Hidden = False Then
                    fundCoperID = r.Cells(1, colIndex).Value
                    If Not dictSP.Exists(fundCoperID) Then
                        ' Store the record as an array:
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
    
    ' Define headers in the order required
    headers = Array("HFAD_Fund_CoperID", "HFAD_Fund_Name", "HFAD_IM_CoperID", _
                    "HFAD_IM_Name", "HFAD_Credit_Officer", "Tier", "Status")
    For j = LBound(headers) To UBound(headers)
        wsUpload.Cells(1, j + 1).Value = headers(j)
    Next j
    
    ' Write new funds data (each record from the collection)
    rowCounter = 2
    For i = 1 To newFunds.Count
        rec = newFunds(i)
        For j = LBound(rec) To UBound(rec)
            wsUpload.Cells(rowCounter, j + 1).Value = rec(j)
        Next j
        rowCounter = rowCounter + 1
    Next i
    
    ' Convert the Upload data into a table named "UploadHF"
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
    
    MsgBox "Macro completed successfully.", vbInformation
End Sub

' -------------------------------
' Helper function: Returns the relative column index (within the ListObject)
' for a given header name. Returns 0 if not found.
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
        GetAllowedValues = Array() ' return empty array if header not found
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
        GetAllowedValues = Array() ' return empty array if no allowed values found
    End If
End Function

' -------------------------------
' Helper function: Appends a value to an existing array.
Function AppendToArray(arr As Variant, valueToAppend As Variant) As Variant
    Dim newArr() As Variant
    Dim i As Long, n As Long
    ' If arr is not an array, create a new array with two elements.
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
