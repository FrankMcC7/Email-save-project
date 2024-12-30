
' ModTypes.bas
Option Explicit

' Public type definition for use across modules
Public Type IALevelData
    GCI As String
    Region As String
    Manager As String
    TriggerStatus As String
    NavSources As Collection
    ClientContacts As String
    TriggerCount As Long
    NonTriggerCount As Long
    MissingTriggerCount As Long
    MissingNonTriggerCount As Long
    ManualData As Variant
End Type


































' ModCoffee.bas
Option Explicit

' Constants for column names to improve maintainability
Private Const COL_FUND_MANAGER_GCI As String = "Fund Manager GCI"
Private Const COL_REGION As String = "Region"
Private Const COL_FUND_MANAGER As String = "Fund Manager"
Private Const COL_TRIGGER_NON_TRIGGER As String = "Trigger/Non-Trigger"
Private Const COL_NAV_SOURCE As String = "NAV Source"
Private Const COL_PRIMARY_CONTACT As String = "Primary Client Contact"
Private Const COL_SECONDARY_CONTACT As String = "Secondary Client Contact"
Private Const COL_WKS_MISSING As String = "Wks Missing"

'--------------------------------------------------------------------------------
' Main Procedure
'--------------------------------------------------------------------------------
Public Sub MacroEpsilon()
    On Error GoTo ErrorHandler
    
    ' Application settings optimization
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' Initialize core worksheet references
    Dim wsPortfolio As Worksheet: Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Dim wsIA As Worksheet: Set wsIA = ThisWorkbook.Sheets("IA_Level")
    
    ' Initialize and validate tables
    Dim portfolioTable As ListObject: Set portfolioTable = GetTableSafely(wsPortfolio, "PortfolioTable")
    Dim iaTable As ListObject: Set iaTable = InitializeIATable(wsIA)
    
    If portfolioTable Is Nothing Or iaTable Is Nothing Then GoTo CleanUp
    
    ' Validate required columns exist
    If Not ValidateRequiredColumns(portfolioTable) Then GoTo CleanUp
    
    ' Process manual data import if user wants it
    Dim manualData As New Collection
    Dim wantManualData As Boolean
    wantManualData = GetUserManualDataPreference()
    
    If wantManualData Then
        If Not ImportManualData(manualData) Then
            wantManualData = False ' Failed to import, proceed without manual data
        End If
    End If
    
    ' Process portfolio data
    Dim iaLevelCollection As Collection
    Set iaLevelCollection = ProcessPortfolioData(portfolioTable, manualData, wantManualData)
    
    ' Write processed data to IA table
    WriteIATableData iaTable, iaLevelCollection
    
    MsgBox "IA_Table has been updated successfully.", vbInformation

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

'--------------------------------------------------------------------------------
' Core Processing Functions
'--------------------------------------------------------------------------------
Private Function ProcessPortfolioData(ByVal portfolioTable As ListObject, _
                                    ByVal manualData As Collection, _
                                    ByVal hasManualData As Boolean) As Collection
    Set ProcessPortfolioData = New Collection
    
    ' Get portfolio data
    Dim data As Variant
    data = portfolioTable.DataBodyRange.Value
    
    ' Get column indices
    Dim colIndices As Object
    Set colIndices = GetColumnIndices(portfolioTable)
    
    ' Process each row
    Dim i As Long
    Dim iaData As IALevelData
    
    For i = 1 To UBound(data, 1)
        Dim gci As String
        gci = SafeString(data(i, colIndices(COL_FUND_MANAGER_GCI)))
        
        If Len(gci) > 0 Then
            ' Get or create IALevelData for this GCI
            If Not CollectionHasKey(ProcessPortfolioData, gci) Then
                InitializeIALevelData iaData, gci, _
                    data(i, colIndices(COL_REGION)), _
                    data(i, colIndices(COL_FUND_MANAGER))
                    
                If hasManualData Then
                    iaData.ManualData = GetManualDataFromCollection(manualData, gci)
                End If
                
                ProcessPortfolioData.Add iaData, gci
            End If
            
            ' Update counts and data
            UpdateIALevelData ProcessPortfolioData.Item(gci), _
                data(i, colIndices(COL_TRIGGER_NON_TRIGGER)), _
                SafeString(data(i, colIndices(COL_NAV_SOURCE))), _
                SafeString(data(i, colIndices(COL_PRIMARY_CONTACT))), _
                SafeString(data(i, colIndices(COL_SECONDARY_CONTACT))), _
                SafeString(data(i, colIndices(COL_WKS_MISSING)))
        End If
    Next i
End Function

Private Function ImportManualData(ByRef manualData As Collection) As Boolean
    Dim previousFile As String
    previousFile = Application.GetOpenFilename( _
        "Excel Files (*.xls*), *.xls*", , _
        "Select previous version")
    
    If previousFile = "False" Then
        MsgBox "No file selected. Manual columns will be empty.", vbExclamation
        ImportManualData = False
        Exit Function
    End If
    
    Dim wbPrev As Workbook
    Set wbPrev = Workbooks.Open(previousFile)
    
    On Error Resume Next
    Dim wsIAPrev As Worksheet
    Set wsIAPrev = wbPrev.Sheets("IA_Level")
    
    If wsIAPrev Is Nothing Then
        MsgBox "IA_Level sheet not found in previous file.", vbCritical
        wbPrev.Close False
        ImportManualData = False
        Exit Function
    End If
    
    Dim iaPrevTable As ListObject
    Set iaPrevTable = wsIAPrev.ListObjects("IA_Table")
    
    If iaPrevTable Is Nothing Or iaPrevTable.DataBodyRange Is Nothing Then
        MsgBox "No data found in previous IA_Table.", vbExclamation
        wbPrev.Close False
        ImportManualData = False
        Exit Function
    End If
    
    ' Import manual data
    Dim data As Variant
    data = iaPrevTable.DataBodyRange.Value
    
    Dim manualIndices As Variant
    manualIndices = Array(13, 14, 15, 16, 17, 18, 19, 20) ' Manual column indices
    
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim gci As String
        gci = data(r, GetColumnIndex(iaPrevTable, COL_FUND_MANAGER_GCI))
        
        If Len(gci) > 0 Then
            Dim manualVals(1 To 8) As Variant
            Dim i As Long
            For i = LBound(manualIndices) To UBound(manualIndices)
                manualVals(i + 1) = data(r, manualIndices(i))
            Next i
            If Not CollectionHasKey(manualData, gci) Then
                manualData.Add manualVals, gci
            End If
        End If
    Next r
    
    wbPrev.Close False
    ImportManualData = True
End Function

Private Sub WriteIATableData(ByVal iaTable As ListObject, ByVal data As Collection)
    ' Prepare array for writing
    Dim result() As Variant
    ReDim result(1 To data.Count, 1 To 20)
    
    Dim i As Long: i = 1
    Dim item As IALevelData
    
    For Each item In data
        With item
            result(i, 1) = .GCI
            result(i, 2) = .Region
            result(i, 3) = .Manager
            result(i, 4) = GetTriggerStatus(.TriggerCount, .NonTriggerCount)
            result(i, 5) = GetNavSourcesString(.NavSources)
            result(i, 6) = .ClientContacts
            result(i, 7) = .TriggerCount
            result(i, 8) = .NonTriggerCount
            result(i, 9) = .TriggerCount + .NonTriggerCount
            result(i, 10) = .MissingTriggerCount
            result(i, 11) = .MissingNonTriggerCount
            result(i, 12) = .MissingTriggerCount + .MissingNonTriggerCount
            
            ' Manual data columns (13-20)
            If Not IsEmpty(.ManualData) Then
                Dim j As Long
                For j = 1 To 8
                    result(i, j + 12) = .ManualData(j)
                Next j
            End If
        End With
        i = i + 1
    Next item
    
    ' Write to table
    Dim targetRange As Range
    Set targetRange = iaTable.HeaderRowRange.Offset(1, 0).Resize(data.Count, 20)
    targetRange.Value = result
    
    ' Resize table
    iaTable.Resize iaTable.Range.Resize(data.Count + 1, 20)
End Sub

'--------------------------------------------------------------------------------
' Table Management Functions
'--------------------------------------------------------------------------------
Private Function GetTableSafely(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetTableSafely = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If GetTableSafely Is Nothing Then
        MsgBox "Table '" & tableName & "' not found on sheet '" & ws.Name & "'", vbCritical
    ElseIf GetTableSafely.DataBodyRange Is Nothing Then
        MsgBox "No data found in '" & tableName & "'", vbExclamation
        Set GetTableSafely = Nothing
    End If
End Function

Private Function InitializeIATable(ws As Worksheet) As ListObject
    On Error Resume Next
    Set InitializeIATable = ws.ListObjects("IA_Table")
    On Error GoTo 0
    
    If InitializeIATable Is Nothing Then
        ' Define required columns for IA table
        Dim requiredCols As Variant
        requiredCols = Array( _
            COL_FUND_MANAGER_GCI, COL_REGION, COL_FUND_MANAGER, _
            COL_TRIGGER_NON_TRIGGER, COL_NAV_SOURCE, "Client Contact(s)", _
            "Trigger", "Non-Trigger", "Total Funds", "Missing Trigger", _
            "Missing Non-Trigger", "Total Missing", "Days to Report", _
            "1st Client Outreach Date", "2nd Client Outreach Date", _
            "OA Escalation Date", "NOA Escalation Date", "Escalation Name", _
            "Final Status", "Comments")
        
        ' Create headers
        Dim col As Long
        For col = LBound(requiredCols) To UBound(requiredCols)
            ws.Cells(1, col + 1).Value = requiredCols(col)
        Next col
        
        ' Create table
        Set InitializeIATable = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:T1"), _
            XlListObjectHasHeaders:=xlYes)
        InitializeIATable.Name = "IA_Table"
    End If
    
    ' Clear existing data
    If Not InitializeIATable.DataBodyRange Is Nothing Then
        InitializeIATable.DataBodyRange.Delete
    End If
End Function

'--------------------------------------------------------------------------------
' Helper Functions
'--------------------------------------------------------------------------------
Private Function ValidateRequiredColumns(tbl As ListObject) As Boolean
    Dim requiredCols As Variant
    requiredCols = Array( _
        COL_FUND_MANAGER_GCI, COL_REGION, COL_FUND_MANAGER, _
        COL_TRIGGER_NON_TRIGGER, COL_NAV_SOURCE, COL_PRIMARY_CONTACT, _
        COL_SECONDARY_CONTACT, COL_WKS_MISSING)
    
    Dim col As Variant
    For Each col In requiredCols
        If GetColumnIndex(tbl, CStr(col)) = 0 Then
            MsgBox "Required column '" & col & "' missing in " & tbl.Name, vbCritical
            ValidateRequiredColumns = False
            Exit Function
        End If
    Next col
    
    ValidateRequiredColumns = True
End Function

Private Function GetUserManualDataPreference() As Boolean
    GetUserManualDataPreference = (MsgBox("Load manual data from previous version?", _
        vbYesNo + vbQuestion, "Manual Data Import") = vbYes)
End Function

Private Function GetColumnIndex(tbl As ListObject, colName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
End Function

Private Function GetColumnIndices(tbl As ListObject) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Add required column indices to dictionary
    result.Add COL_FUND_MANAGER_GCI, GetColumnIndex(tbl, COL_FUND_MANAGER_GCI)
    result.Add COL_REGION, GetColumnIndex(tbl, COL_REGION)
    result.Add COL_FUND_MANAGER, GetColumnIndex(tbl, COL_FUND_MANAGER)
    result.Add COL_TRIGGER_NON_TRIGGER, GetColumnIndex(tbl, COL_TRIGGER_NON_TRIGGER)
    result.Add COL_NAV_SOURCE, GetColumnIndex(tbl, COL_NAV_SOURCE)
    result.Add COL_PRIMARY_CONTACT, GetColumnIndex(tbl, COL_PRIMARY_CONTACT)
    result.Add COL_SECONDARY_CONTACT, GetColumnIndex(tbl, COL_SECONDARY_CONTACT)
    result.Add COL_WKS_MISSING, GetColumnIndex(tbl, COL_WKS_MISSING)
    
    Set GetColumnIndices = result
End Function

Private Function SafeString(val As Variant) As String
    If IsError(val) Then
        SafeString = ""
    Else
        SafeString = CStr(val)
    End If
End Function

Private Function CollectionHasKey(col As Collection, key As String) As Boolean
    On Error Resume Next
    col.Item key
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function CollectionHasValue(col As Collection, value As String) As Boolean
    Dim item As Variant
    For Each item In col
        If item = value Then
            CollectionHasValue = True
            Exit Function
        End If
    Next item
    CollectionHasValue = False
End Function

Private Function GetManualDataFromCollection(ByVal manualData As Collection, ByVal gci As String) As Variant
    On Error Resume Next
    GetManualDataFromCollection = manualData.Item(gci)
    On Error GoTo 0
End Function

Private Sub InitializeIALevelData(ByRef iaData As IALevelData, _
                                ByVal gci As String, _
                                ByVal region As String, _
                                ByVal manager As String)
    With iaData
        .GCI = gci
        .Region = region
        .Manager = manager
        .TriggerStatus = ""
        Set .NavSources = New Collection
        .ClientContacts = ""
        .TriggerCount = 0
        .NonTriggerCount = 0
        .MissingTriggerCount = 0
        .MissingNonTriggerCount = 0
        .ManualData = Empty
    End With
End Sub

Private Sub UpdateIALevelData(ByRef iaData As IALevelData, _
                            ByVal triggerStatus As String, _
                            ByVal navSource As String, _
                            ByVal primaryContact As String, _
                            ByVal secondaryContact As String, _
                            ByVal weeksMissing As String)
    With iaData
        ' Update trigger counts
        If triggerStatus = "Trigger" Then
            .TriggerCount = .TriggerCount + 1
            If Len(weeksMissing) > 0 Then
                .MissingTriggerCount = .MissingTriggerCount + 1
            End If
        ElseIf triggerStatus = "Non-Trigger" Then
            .NonTriggerCount = .NonTriggerCount + 1
            If Len(weeksMissing) > 0 Then
                .MissingNonTriggerCount = .MissingNonTriggerCount + 1
            End If
        End If
        
        ' Update NAV Sources
        If Len(navSource) > 0 Then
            If Not CollectionHasValue(.NavSources, navSource) Then
                .NavSources.Add navSource
            End If
        End If
        
        ' Update Client Contacts
        Dim contacts As String
        contacts = ""
        If Len(primaryContact) > 0 Then contacts = primaryContact
        If Len(secondaryContact) > 0 Then
            If Len(contacts) > 0 Then
                contacts = contacts & ";" & secondaryContact
            Else
                contacts = secondaryContact
            End If
        End If
        If Len(contacts) > 0 Then .ClientContacts = contacts
    End With
End Sub

Private Function GetTriggerStatus(triggerCount As Long, nonTriggerCount As Long) As String
    If triggerCount > 0 And nonTriggerCount > 0 Then
        GetTriggerStatus = "Both"
    ElseIf triggerCount > 0 Then
        GetTriggerStatus = "Trigger"
    ElseIf nonTriggerCount > 0 Then
        GetTriggerStatus = "Non-Trigger"
    Else
        GetTriggerStatus = ""
    End If
End Function

Private Function GetNavSourcesString(navSources As Collection) As String
    If navSources.Count = 0 Then
        GetNavSourcesString = "[No NAV Source]"
        Exit Function
    End If
    
    Dim result As String
    Dim src As Variant
    For Each src In navSources
        If Len(result) > 0 Then result = result & ";"
        result = result & src
    Next src
    
    GetNavSourcesString = result
End Function
