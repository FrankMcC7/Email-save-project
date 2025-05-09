' === IALevelData Class Module (Save as IALevelData.cls) ===
Option Explicit

' Private member variables
Private m_GCI As String
Private m_Region As String
Private m_Manager As String
Private m_ECAIndiaAnalyst As String
Private m_TriggerStatus As String
Private m_NavSources As Collection
Private m_ClientContacts As String
Private m_TriggerCount As Long
Private m_NonTriggerCount As Long
Private m_MissingTriggerCount As Long
Private m_MissingNonTriggerCount As Long
Private m_ManualData As Variant

' Initialize the class
Private Sub Class_Initialize()
    Set m_NavSources = New Collection
End Sub

' Clean up
Private Sub Class_Terminate()
    Set m_NavSources = Nothing
End Sub

' Property Get/Let methods
Public Property Get GCI() As String
    GCI = m_GCI
End Property

Public Property Let GCI(value As String)
    m_GCI = value
End Property

Public Property Get Region() As String
    Region = m_Region
End Property

Public Property Let Region(value As String)
    m_Region = value
End Property

Public Property Get Manager() As String
    Manager = m_Manager
End Property

Public Property Let Manager(value As String)
    m_Manager = value
End Property

Public Property Get ECAIndiaAnalyst() As String
    ECAIndiaAnalyst = m_ECAIndiaAnalyst
End Property

Public Property Let ECAIndiaAnalyst(value As String)
    m_ECAIndiaAnalyst = value
End Property

Public Property Get TriggerStatus() As String
    TriggerStatus = m_TriggerStatus
End Property

Public Property Let TriggerStatus(value As String)
    m_TriggerStatus = value
End Property

Public Property Get NavSources() As Collection
    Set NavSources = m_NavSources
End Property

Public Property Set NavSources(value As Collection)
    Set m_NavSources = value
End Property

Public Property Get ClientContacts() As String
    ClientContacts = m_ClientContacts
End Property

Public Property Let ClientContacts(value As String)
    m_ClientContacts = value
End Property

Public Property Get TriggerCount() As Long
    TriggerCount = m_TriggerCount
End Property

Public Property Let TriggerCount(value As Long)
    m_TriggerCount = value
End Property

Public Property Get NonTriggerCount() As Long
    NonTriggerCount = m_NonTriggerCount
End Property

Public Property Let NonTriggerCount(value As Long)
    m_NonTriggerCount = value
End Property

Public Property Get MissingTriggerCount() As Long
    MissingTriggerCount = m_MissingTriggerCount
End Property

Public Property Let MissingTriggerCount(value As Long)
    m_MissingTriggerCount = value
End Property

Public Property Get MissingNonTriggerCount() As Long
    MissingNonTriggerCount = m_MissingNonTriggerCount
End Property

Public Property Let MissingNonTriggerCount(value As Long)
    m_MissingNonTriggerCount = value
End Property

Public Property Get ManualData() As Variant
    If IsObject(m_ManualData) Then
        Set ManualData = m_ManualData
    Else
        ManualData = m_ManualData
    End If
End Property

Public Property Let ManualData(value As Variant)
    m_ManualData = value
End Property

Public Property Set ManualData(value As Variant)
    Set m_ManualData = value
End Property

' === Main Module (Save as EpsilonModule.bas) ===
Option Explicit

' Constants for column names
Private Const COL_FUND_MANAGER_GCI As String = "Fund Manager GCI"
Private Const COL_REGION As String = "Region"
Private Const COL_FUND_MANAGER As String = "Fund Manager"
Private Const COL_ECA_INDIA_ANALYST As String = "ECA India Analyst"
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
        .DisplayAlerts = False
    End With
    
    ' Initialize core worksheet references
    Dim wsPortfolio As Worksheet: Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    Dim wsIA As Worksheet
    
    ' Check if IA_Level sheet exists, if not create it
    On Error Resume Next
    Set wsIA = ThisWorkbook.Sheets("IA_Level")
    On Error GoTo 0
    
    If wsIA Is Nothing Then
        Set wsIA = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsIA.Name = "IA_Level"
    End If
    
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
    
    ' Format the IA table
    FormatIATable iaTable
    
    MsgBox "IA_Table has been updated successfully.", vbInformation

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
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
                Set iaData = New IALevelData
                
                ' Initialize data
                With iaData
                    .GCI = gci
                    .Region = SafeString(data(i, colIndices(COL_REGION)))
                    .Manager = SafeString(data(i, colIndices(COL_FUND_MANAGER)))
                    .ECAIndiaAnalyst = SafeString(data(i, colIndices(COL_ECA_INDIA_ANALYST)))
                End With
                
                If hasManualData Then
                    iaData.ManualData = GetManualDataFromCollection(manualData, gci)
                End If
                
                ProcessPortfolioData.Add Item:=iaData, Key:=gci
            End If
            
            ' Update counts and data
            Dim currentData As IALevelData
            Set currentData = ProcessPortfolioData.Item(gci)
            
            UpdateIALevelData currentData, _
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
    Set wbPrev = Workbooks.Open(previousFile, ReadOnly:=True)
    
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
    
    If iaPrevTable Is Nothing Then
        MsgBox "IA_Table not found in previous version.", vbCritical
        wbPrev.Close False
        ImportManualData = False
        Exit Function
    End If
    
    ' Import data
    Dim data As Variant
    data = iaPrevTable.DataBodyRange.Value
    
    Dim manualCols As Variant
    manualCols = Array(14, 15, 16, 17, 18, 19, 20, 21)  ' Manual column indices
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim gci As String
        gci = data(i, GetColumnIndex(iaPrevTable, COL_FUND_MANAGER_GCI))
        
        If Len(gci) > 0 Then
            Dim manualVals(1 To 8) As Variant
            Dim j As Long
            For j = LBound(manualCols) To UBound(manualCols)
                manualVals(j + 1) = data(i, manualCols(j))
            Next j
            If Not CollectionHasKey(manualData, gci) Then
                manualData.Add Item:=manualVals, Key:=gci
            End If
        End If
    Next i
    
    wbPrev.Close False
    ImportManualData = True
End Function

Private Sub WriteIATableData(ByVal iaTable As ListObject, ByVal data As Collection)
    If data.Count = 0 Then Exit Sub
    
    ' Clear existing data if any
    If Not iaTable.DataBodyRange Is Nothing Then
        iaTable.DataBodyRange.Delete
    End If
    
    ' Prepare array for writing
    Dim result() As Variant
    ReDim result(1 To data.Count, 1 To 21)
    
    Dim i As Long: i = 1
    Dim item As IALevelData
    
    For Each item In data
        With item
            result(i, 1) = .GCI
            result(i, 2) = .Region
            result(i, 3) = .Manager
            result(i, 4) = .ECAIndiaAnalyst
            result(i, 5) = GetTriggerStatus(.TriggerCount, .NonTriggerCount)
            result(i, 6) = GetNavSourcesString(.NavSources)
            result(i, 7) = .ClientContacts
            result(i, 8) = .TriggerCount
            result(i, 9) = .NonTriggerCount
            result(i, 10) = .TriggerCount + .NonTriggerCount
            result(i, 11) = .MissingTriggerCount
            result(i, 12) = .MissingNonTriggerCount
            result(i, 13) = .MissingTriggerCount + .MissingNonTriggerCount
            
            ' Manual data columns (14-21)
            If Not IsEmpty(.ManualData) Then
                Dim j As Long
                For j = 1 To 8
                    result(i, j + 13) = .ManualData(j)
                Next j
            End If
        End With
        i = i + 1
    Next item
    
    ' Add required rows one at a time
    Dim rowsNeeded As Long
    rowsNeeded = data.Count - 1  ' Subtract 1 because table starts with one row
    
    Dim k As Long
    For k = 1 To rowsNeeded
        iaTable.ListRows.Add
    Next k
    
    ' Write data
    iaTable.DataBodyRange.Value = result
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
    ' Delete existing table if present
    On Error Resume Next
    If Not ws.ListObjects("IA_Table") Is Nothing Then
        ws.ListObjects("IA_Table").Delete
    End If
    On Error GoTo 0
    
    ' Define headers
    Dim headers As Variant
    headers = Array( _
        COL_FUND_MANAGER_GCI, COL_REGION, COL_FUND_MANAGER, COL_ECA_INDIA_ANALYST, _
        COL_TRIGGER_NON_TRIGGER, COL_NAV_SOURCE, "Client Contact(s)", _
        "Trigger", "Non-Trigger", "Total Funds", _
        "Missing Trigger", "Missing Non-Trigger", "Total Missing", _
        "Days to Report", "1st Client Outreach Date", "2nd Client Outreach Date", _
        "OA Escalation Date", "NOA Escalation Date", "Escalation Name", _
        "Final Status", "Comments" _
    )
    
    ' Write headers
    Dim col As Long
    For col = LBound(headers) To UBound(headers)
        ws.Cells(1, col + 1).Value = headers(col)
    Next col
    
    ' Create table
    Set InitializeIATable = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)), _
        XlListObjectHasHeaders:=xlYes)
    
    InitializeIATable.Name = "IA_Table"
    
    ' Basic formatting
    With InitializeIATable.HeaderRowRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
End Function

Private Function ValidateRequiredColumns(tbl As ListObject) As Boolean
    Dim requiredCols As Variant
    requiredCols = Array( _
        COL_FUND_MANAGER_GCI, COL_REGION, COL_FUND_MANAGER, COL_ECA_INDIA_ANALYST, _
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

Private Sub FormatIATable(tbl As ListObject)
    With tbl
        ' Reset table style first
        .TableStyle = "TableStyleMedium2"
        
        ' Format entire table font
        With .Range
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Bold = False
        End With
        
        ' Format header row
        With .HeaderRowRange
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(217, 225, 242)  ' Light blue background
            .HorizontalAlignment = xlCenter
            ' Ensure text wrapping for headers
            .WrapText = True
            ' Set row height for better readability
            .RowHeight = 30
        End With
        
        ' Format data body
        If Not .DataBodyRange Is Nothing Then
            With .DataBodyRange
                .Font.Bold = False
                .Interior.ColorIndex = xlNone
                .VerticalAlignment = xlCenter
            End With
        End If
        
        ' AutoFit columns with a max width
        .Range.Columns.AutoFit
        Dim col As ListColumn
        For Each col In .ListColumns
            If col.Range.ColumnWidth > 30 Then
                col.Range.ColumnWidth = 30
            ElseIf col.Range.ColumnWidth < 8 Then
                col.Range.ColumnWidth = 8
            End If
        Next col
        
        ' Format date columns
        Dim dateColumns As Variant
        dateColumns = Array("1st Client Outreach Date", "2nd Client Outreach Date", _
                          "OA Escalation Date", "NOA Escalation Date")
        
        Dim colName As Variant
        For Each colName In dateColumns
            On Error Resume Next
            With .ListColumns(CStr(colName)).DataBodyRange
                .NumberFormat = "dd-mmm-yyyy"
                .HorizontalAlignment = xlCenter
            End With
            On Error GoTo 0
        Next colName
        
        ' Format numeric columns
        With .ListColumns("Days to Report").DataBodyRange
            .NumberFormat = "0"
            .HorizontalAlignment = xlRight
        End With
        
        ' Center align specific columns
        Dim centerColumns As Variant
        centerColumns = Array("Trigger", "Non-Trigger", "Total Funds", _
                            "Missing Trigger", "Missing Non-Trigger", "Total Missing")
        
        For Each colName In centerColumns
            On Error Resume Next
            .ListColumns(CStr(colName)).DataBodyRange.HorizontalAlignment = xlCenter
            On Error GoTo 0
        Next colName
        
        ' Apply borders
        With .Range.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)  ' Light gray borders
        End With
        
        ' Set alternating row colors
        .ShowTableStyleRowStripes = True
        .ShowTableStyleColumnStripes = False
        
        ' Ensure frozen panes
        .Range.Worksheet.Activate
        ActiveWindow.FreezePanes = False
        .HeaderRowRange.Rows(1).Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

Private Function GetUserManualDataPreference() As Boolean
    GetUserManualDataPreference = (MsgBox("Load manual data from previous version?", _
        vbYesNo + vbQuestion, "Manual Data Import") = vbYes)
End Function

'--------------------------------------------------------------------------------
' Utility Functions
'--------------------------------------------------------------------------------
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
    result.Add COL_ECA_INDIA_ANALYST, GetColumnIndex(tbl, COL_ECA_INDIA_ANALYST)
    result.Add COL_TRIGGER_NON_TRIGGER, GetColumnIndex(tbl, COL_TRIGGER_NON_TRIGGER)
    result.Add COL_NAV_SOURCE, GetColumnIndex(tbl, COL_NAV_SOURCE)
    result.Add COL_PRIMARY_CONTACT, GetColumnIndex(tbl, COL_PRIMARY_CONTACT)
    result.Add COL_SECONDARY_CONTACT, GetColumnIndex(tbl, COL_SECONDARY_CONTACT)
    result.Add COL_WKS_MISSING, GetColumnIndex(tbl, COL_WKS_MISSING)
    
    Set GetColumnIndices = result
End Function

Private Function SafeString(val As Variant) As String
    If IsError(val) Or IsEmpty(val) Or IsNull(val) Then
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

Private Sub UpdateIALevelData(ByRef data As IALevelData, _
                            ByVal triggerStatus As String, _
                            ByVal navSource As String, _
                            ByVal primaryContact As String, _
                            ByVal secondaryContact As String, _
                            ByVal weeksMissing As String)
    With data
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
            If Not CollectionContains(.NavSources, navSource) Then
                .NavSources.Add navSource
            End If
        End If
        
        ' Update Client Contacts
        UpdateClientContacts data, primaryContact, secondaryContact
    End With
End Sub

Private Function CollectionContains(col As Collection, item As String) As Boolean
    Dim var As Variant
    For Each var In col
        If var = item Then
            CollectionContains = True
            Exit Function
        End If
    Next var
    CollectionContains = False
End Function

Private Sub UpdateClientContacts(ByRef data As IALevelData, _
                               ByVal primary As String, _
                               ByVal secondary As String)
    Dim contacts As String
    contacts = ""
    
    ' Add primary contact if present
    If Len(primary) > 0 Then
        contacts = primary
    End If
    
    ' Add secondary contact if present
    If Len(secondary) > 0 Then
        If Len(contacts) > 0 Then
            contacts = contacts & "; " & secondary
        Else
            contacts = secondary
        End If
    End If
    
    ' Update if we have contacts
    If Len(contacts) > 0 Then
        If Len(data.ClientContacts) > 0 Then
            ' Check if contacts are already present
            If InStr(1, data.ClientContacts, contacts, vbTextCompare) = 0 Then
                data.ClientContacts = data.ClientContacts & "; " & contacts
            End If
        Else
            data.ClientContacts = contacts
        End If
    End If
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
        If Len(result) > 0 Then
            result = result & "; "
        End If
        result = result & CStr(src)
    Next src
    
    GetNavSourcesString = result
End Function
