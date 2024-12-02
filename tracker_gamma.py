Sub MacroGamma()
    Dim wbMaster As Workbook
    Dim wbPrevious As Workbook
    Dim previousFile As String

    Dim wsPortfolio As Worksheet
    Dim wsEMEA As Worksheet, wsAMRS As Worksheet, wsAPAC As Worksheet
    Dim wsEMEAPrev As Worksheet, wsAMRSPrev As Worksheet, wsAPACPrev As Worksheet

    Dim portfolioTable As ListObject
    Dim tblEMEA As ListObject, tblAMRS As ListObject, tblAPAC As ListObject
    Dim tblEMEAPrev As ListObject, tblAMRSPrev As ListObject, tblAPACPrev As ListObject

    Dim colsToClear As Variant
    Dim colName As Variant
    Dim colIndex As Long
    Dim arrRegions As Variant
    Dim arrTables As Variant
    Dim arrPrevTables As Variant
    Dim region As String
    Dim i As Long, j As Long, k As Long
    Dim regionField As Long
    Dim wksMissingField As Long
    Dim familyField As Long
    Dim rngFamily As Range
    Dim dictFamilies As Object
    Dim arrFamilies As Variant
    Dim numFamilies As Long
    Dim tbl As ListObject
    Dim tblPrev As ListObject
    Dim dictPrevData As Object
    Dim prevFamilyField As Long
    Dim prevDataRow As Range
    Dim prevFamily As String
    Dim currFamily As String
    Dim dataValues() As Variant
    Dim colNames(1 To 6) As String
    Dim prevColIndices(1 To 6) As Long
    Dim familyCell As Range

    ' Variables for Data Validation
    Dim dvType As Long
    Dim dvAlertStyle As Long
    Dim dvFormula1 As String
    Dim dvIgnoreBlank As Boolean
    Dim dvInCellDropdown As Boolean

    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Set up workbooks and worksheets
    Set wbMaster = ThisWorkbook
    Set wsPortfolio = wbMaster.Sheets("Portfolio")
    Set wsEMEA = wbMaster.Sheets("EMEA")
    Set wsAMRS = wbMaster.Sheets("AMRS")
    Set wsAPAC = wbMaster.Sheets("APAC")

    ' Ensure PortfolioTable exists
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    If portfolioTable Is Nothing Then
        MsgBox "PortfolioTable does not exist on the Portfolio sheet.", vbCritical
        GoTo CleanUp
    End If

    ' Ensure region tables exist
    On Error Resume Next
    Set tblEMEA = wsEMEA.ListObjects("EMEA")
    Set tblAMRS = wsAMRS.ListObjects("AMRS")
    Set tblAPAC = wsAPAC.ListObjects("APAC")
    On Error GoTo 0
    If tblEMEA Is Nothing Or tblAMRS Is Nothing Or tblAPAC Is Nothing Then
        MsgBox "One or more tables (EMEA, AMRS, APAC) are missing in the current workbook.", vbCritical
        GoTo CleanUp
    End If

    ' Step 3: Delete specified columns in all three tables
    colsToClear = Array("Family", "Last Action", "Last Action Date", "Action Taker", "CPO/ECA", "Remediation Action Holder", "Comments")
    arrTables = Array(tblEMEA, tblAMRS, tblAPAC)
    For i = LBound(arrTables) To UBound(arrTables)
        Set tbl = arrTables(i)
        For Each colName In colsToClear
            On Error Resume Next
            colIndex = tbl.ListColumns(colName).Index
            On Error GoTo 0
            If colIndex > 0 Then
                If Not tbl.ListColumns(colName).DataBodyRange Is Nothing Then
                    tbl.ListColumns(colName).DataBodyRange.ClearContents
                End If
            Else
                MsgBox "Column '" & colName & "' not found in table '" & tbl.Name & "'.", vbExclamation
            End If
        Next colName
    Next i

    ' Step 4: Ask user for the previous version of tracker file
    previousFile = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , "Select the previous version of the tracker file")
    If previousFile = "False" Then
        MsgBox "No file selected.", vbExclamation
        GoTo CleanUp
    End If
    Set wbPrevious = Workbooks.Open(previousFile)

    ' Ensure previous version's tables exist
    Set wsEMEAPrev = wbPrevious.Sheets("EMEA")
    Set wsAMRSPrev = wbPrevious.Sheets("AMRS")
    Set wsAPACPrev = wbPrevious.Sheets("APAC")

    On Error Resume Next
    Set tblEMEAPrev = wsEMEAPrev.ListObjects("EMEA")
    Set tblAMRSPrev = wsAMRSPrev.ListObjects("AMRS")
    Set tblAPACPrev = wsAPACPrev.ListObjects("APAC")
    On Error GoTo 0
    If tblEMEAPrev Is Nothing Or tblAMRSPrev Is Nothing Or tblAPACPrev Is Nothing Then
        MsgBox "One or more tables (EMEA, AMRS, APAC) are missing in the previous version.", vbCritical
        wbPrevious.Close SaveChanges:=False
        GoTo CleanUp
    End If

    ' Prepare arrays for regions and tables
    arrRegions = Array("EMEA", "AMRS", "APAC")
    arrTables = Array(tblEMEA, tblAMRS, tblAPAC)
    arrPrevTables = Array(tblEMEAPrev, tblAMRSPrev, tblAPACPrev)

    ' Get field indices in PortfolioTable
    regionField = portfolioTable.ListColumns("Region").Index
    wksMissingField = portfolioTable.ListColumns("Wks Missing").Index
    familyField = portfolioTable.ListColumns("Family").Index

    ' Prepare column names and indices
    colNames(1) = "Last Action"
    colNames(2) = "Last Action Date"
    colNames(3) = "Action Taker"
    colNames(4) = "CPO/ECA"
    colNames(5) = "Remediation Action Holder"
    colNames(6) = "Comments"

    ' Define Data Validation settings
    dvType = xlValidateList
    dvAlertStyle = xlValidAlertStop
    dvFormula1 = "Pending,1st Outreach,2nd Outreach,1st Onshore Review,1st Escalation,2nd Escalation,3rd Escalation,Post-Update,In-Closing"
    dvIgnoreBlank = True
    dvInCellDropdown = True

    ' Step 5 & 6: Process each region
    For i = LBound(arrRegions) To UBound(arrRegions)
        region = arrRegions(i)
        Set tbl = arrTables(i)
        Set tblPrev = arrPrevTables(i)

        ' Remove any existing filters
        If portfolioTable.AutoFilter.FilterMode Then
            portfolioTable.AutoFilter.ShowAllData
        End If

        ' Apply filters to PortfolioTable
        portfolioTable.Range.AutoFilter Field:=regionField, Criteria1:=region
        portfolioTable.Range.AutoFilter Field:=wksMissingField, Criteria1:="<>"

        ' Get 'Family' column visible cells
        On Error Resume Next
        Set rngFamily = portfolioTable.ListColumns("Family").DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        ' Collect unique 'Family' values
        Set dictFamilies = CreateObject("Scripting.Dictionary")
        If Not rngFamily Is Nothing Then
            For Each familyCell In rngFamily.Cells
                If Not IsEmpty(familyCell.Value) And Not dictFamilies.Exists(familyCell.Value) Then
                    dictFamilies.Add familyCell.Value, Nothing
                End If
            Next familyCell
        End If

        ' Get array of unique families
        If dictFamilies.Count > 0 Then
            arrFamilies = dictFamilies.Keys
            numFamilies = dictFamilies.Count
        Else
            arrFamilies = Array()
            numFamilies = 0
        End If

        ' Adjust the current table rows
        ' Clear existing data in 'Family' column
        If Not tbl.ListColumns("Family").DataBodyRange Is Nothing Then
            tbl.ListColumns("Family").DataBodyRange.ClearContents
        End If

        ' Remove extra rows
        If tbl.ListRows.Count > numFamilies Then
            For j = tbl.ListRows.Count To numFamilies + 1 Step -1
                tbl.ListRows(j).Delete
            Next j
        End If

        ' Add missing rows
        If tbl.ListRows.Count < numFamilies Then
            For j = tbl.ListRows.Count + 1 To numFamilies
                tbl.ListRows.Add
            Next j
        End If

        ' Write 'Family' values into the current table
        For j = 1 To numFamilies
            tbl.ListColumns("Family").DataBodyRange.Cells(j, 1).Value = arrFamilies(j - 1)
        Next j

        ' Apply Data Validation to 'Last Action' column
        If Not tbl.ListColumns("Last Action").DataBodyRange Is Nothing Then
            With tbl.ListColumns("Last Action").DataBodyRange.Validation
                .Delete
                .Add Type:=dvType, AlertStyle:=dvAlertStyle, Formula1:=dvFormula1
                .IgnoreBlank = dvIgnoreBlank
                .InCellDropdown = dvInCellDropdown
            End With
        End If

        ' Create a dictionary from previous table based on 'Family'
        Set dictPrevData = CreateObject("Scripting.Dictionary")
        ' Get indices of the columns in previous table
        For k = 1 To 6
            On Error Resume Next
            prevColIndices(k) = tblPrev.ListColumns(colNames(k)).Index
            On Error GoTo 0
            If prevColIndices(k) = 0 Then
                MsgBox "Column '" & colNames(k) & "' not found in table '" & tblPrev.Name & "' in previous version.", vbCritical
                wbPrevious.Close SaveChanges:=False
                GoTo CleanUp
            End If
        Next k
        prevFamilyField = tblPrev.ListColumns("Family").Index

        ' Build the dictionary from previous table
        For Each prevDataRow In tblPrev.DataBodyRange.Rows
            prevFamily = prevDataRow.Cells(1, prevFamilyField).Value
            If Not IsEmpty(prevFamily) And Not dictPrevData.Exists(prevFamily) Then
                ReDim dataValues(1 To 6)
                For k = 1 To 6
                    dataValues(k) = prevDataRow.Cells(1, prevColIndices(k)).Value
                Next k
                dictPrevData.Add prevFamily, dataValues
            End If
        Next prevDataRow

        ' Copy data from previous table to current table
        For j = 1 To numFamilies
            currFamily = tbl.ListColumns("Family").DataBodyRange.Cells(j, 1).Value
            If dictPrevData.Exists(currFamily) Then
                dataValues = dictPrevData(currFamily)
                For k = 1 To 6
                    tbl.ListColumns(colNames(k)).DataBodyRange.Cells(j, 1).Value = dataValues(k)
                Next k
            End If
        Next j

    Next i

    ' Remove filters from PortfolioTable
    If portfolioTable.AutoFilter.FilterMode Then
        portfolioTable.AutoFilter.ShowAllData
    End If

    ' Close the previous workbook without saving
    wbPrevious.Close SaveChanges:=False

    ' Completion message
    MsgBox "Macro Gamma has successfully updated the tables.", vbInformation

CleanUp:
    ' Reset application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in Macro Gamma: " & Err.Description, vbCritical
    If Not wbPrevious Is Nothing Then
        wbPrevious.Close SaveChanges:=False
    End If
    Resume CleanUp
End Sub
