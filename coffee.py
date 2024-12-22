' === Configuration and Type Definitions ===
Option Explicit

Private Type ProgressData
    StartTime As Double
    TotalItems As Long
    CurrentItem As Long
End Type

Private Type CacheData
    Dict As Scripting.Dictionary
    Initialized As Boolean
End Type

' === Global Variables ===
Private Progress As ProgressData
Private Cache As CacheData

' === Constants ===
Private Const INPUT_FOLDER As String = "C:\ProjectCoffee\Input\"
Private Const TRIGGER_FILE As String = "Trigger.csv"
Private Const ALL_FUNDS_FILE As String = "All_Funds.csv"
Private Const NON_TRIGGER_FILE As String = "Non_Trigger.csv"

' === Main Control Procedure ===
Public Sub RunProjectCoffee()
    On Error GoTo ErrorHandler
    
    ' Performance optimization
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    ' Initialize progress
    InitializeProgress 4 ' 3 main processes + final step
    
    ' Initialize tables
    InitializeTables
    
    ' Verify files
    If Not VerifyInputFiles Then
        MsgBox "Required input files are missing. Please check the input folder:" & vbNewLine & _
               INPUT_FOLDER, vbCritical
        GoTo CleanUp
    End If
    
    ' Execute main processes
    ProcessAllDataWithPaths INPUT_FOLDER & TRIGGER_FILE, _
                          INPUT_FOLDER & ALL_FUNDS_FILE, _
                          INPUT_FOLDER & NON_TRIGGER_FILE
    UpdateProgress
    
    MacroOmega
    UpdateProgress
    
    ' Get previous version file
    Dim previousFile As String
    previousFile = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*", , _
                  "Select the previous version of the tracker file")
    
    If previousFile <> "False" Then
        MacroGammaWithPath previousFile
        UpdateProgress
        
        MsgBox "Project Coffee processing completed successfully!", vbInformation
    Else
        MsgBox "Previous version file not selected. Process partially completed.", vbExclamation
    End If

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    CleanupMemory
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbNewLine & _
           "Error in line: " & Erl, vbCritical
    Resume CleanUp
End Sub

' === Core Utility Functions ===
Private Function VerifyInputFiles() As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Check folder and files
    If Not FSO.FolderExists(INPUT_FOLDER) Then
        VerifyInputFiles = False
        Exit Function
    End If
    
    Dim requiredFiles As Variant
    requiredFiles = Array(TRIGGER_FILE, ALL_FUNDS_FILE, NON_TRIGGER_FILE)
    
    Dim file As Variant
    For Each file In requiredFiles
        If Not FSO.FileExists(INPUT_FOLDER & file) Then
            VerifyInputFiles = False
            Exit Function
        End If
    Next file
    
    VerifyInputFiles = True
End Function

Private Sub InitializeProgress(totalItems As Long)
    With Progress
        .StartTime = Timer
        .TotalItems = totalItems
        .CurrentItem = 0
    End With
End Sub

Private Sub UpdateProgress(Optional incrementBy As Long = 1)
    With Progress
        .CurrentItem = .CurrentItem + incrementBy
        
        ' Calculate progress
        Dim percentComplete As Double
        percentComplete = (.CurrentItem / .TotalItems) * 100
        
        ' Calculate time estimates
        Dim elapsedTime As Double
        elapsedTime = Timer - .StartTime
        
        Dim estimatedTotalTime As Double
        estimatedTotalTime = (elapsedTime / percentComplete) * 100
        
        Dim remainingTime As Double
        remainingTime = estimatedTotalTime - elapsedTime
        
        ' Update status
        Application.StatusBar = Format(percentComplete, "0.0") & "% Complete. " & _
                              "Est. Time Remaining: " & Format(remainingTime / 86400, "hh:mm:ss")
    End With
    DoEvents
End Sub

Private Sub CleanupMemory()
    If Not Cache.Dict Is Nothing Then
        Cache.Dict.RemoveAll
        Set Cache.Dict = Nothing
    End If
    Cache.Initialized = False
    
    ' Force garbage collection
    Dim i As Long
    For i = 1 To 2
        DoEvents
        Dim tmp() As String
        Erase tmp
    Next i
End Sub

' === CSV and Data Processing Functions ===
Private Function FastReadCSV(filePath As String) As Variant
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Configure and open connection
    With conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & GetDirectoryPath(filePath) & ";" & _
                          "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
        .Open
    End With
    
    ' Read data
    Dim rs As Object
    Set rs = conn.Execute("SELECT * FROM [" & GetFileName(filePath) & "]")
    
    ' Convert to array
    FastReadCSV = rs.GetRows
    
    ' Cleanup
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Function

Private Sub ProcessAllDataWithPaths(triggerPath As String, allFundsPath As String, nonTriggerPath As String)
    ' Setup Portfolio sheet
    Dim wsPortfolio As Worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Setup PortfolioTable
    Dim portfolioTable As ListObject
    Set portfolioTable = EnsureTableExists(wsPortfolio, "PortfolioTable")
    
    ' Clear existing data
    If Not portfolioTable.DataBodyRange Is Nothing Then
        portfolioTable.DataBodyRange.Delete
    End If
    
    ' Process files
    ProcessTriggerData FastReadCSV(triggerPath), portfolioTable
    ProcessAllFundsData FastReadCSV(allFundsPath), portfolioTable
    ProcessNonTriggerData FastReadCSV(nonTriggerPath), portfolioTable
End Sub

Private Sub ProcessTriggerData(data As Variant, portfolioTable As ListObject)
    ' Pre-allocate array for processed data
    Dim processedData() As Variant
    ReDim processedData(1 To UBound(data, 2), 1 To portfolioTable.ListColumns.Count)
    
    ' Process each row
    Dim i As Long
    For i = 1 To UBound(data, 2)
        ' Map and transform data
        processedData(i, GetColumnIndex(portfolioTable, "Region")) = _
            Replace(Replace(data(GetSourceColumnIndex("Region"), i), "US", "AMRS"), "ASIA", "APAC")
            
        processedData(i, GetColumnIndex(portfolioTable, "Fund Manager")) = _
            data(GetSourceColumnIndex("Fund Manager"), i)
            
        processedData(i, GetColumnIndex(portfolioTable, "Fund GCI")) = _
            data(GetSourceColumnIndex("Fund GCI"), i)
            
        processedData(i, GetColumnIndex(portfolioTable, "Fund Name")) = _
            data(GetSourceColumnIndex("Fund Name"), i)
            
        processedData(i, GetColumnIndex(portfolioTable, "Wks Missing")) = _
            data(GetSourceColumnIndex("Wks Missing"), i)
            
        processedData(i, GetColumnIndex(portfolioTable, "Credit Officer")) = _
            data(GetSourceColumnIndex("Credit Officer"), i)
            
        ' Set Trigger status
        processedData(i, GetColumnIndex(portfolioTable, "Trigger/Non-Trigger")) = "Trigger"
    Next i
    
    ' Write processed data to table
    If UBound(processedData, 1) > 0 Then
        portfolioTable.ListRows.Add
        portfolioTable.DataBodyRange.Value = processedData
    End If
End Sub

Private Function ProcessAllFundsData(data As Variant, portfolioTable As ListObject)
    ' Create dictionary for Fund GCI to IA GCI mapping
    Dim dictFundGCI As Object
    Set dictFundGCI = CreateObject("Scripting.Dictionary")
    
    ' Process Approved records
    Dim i As Long
    For i = 1 To UBound(data, 2)
        If LCase(data(GetSourceColumnIndex("Review Status"), i)) = "approved" Then
            Dim fundGCI As String
            fundGCI = data(GetSourceColumnIndex("Fund GCI"), i)
            
            If Not dictFundGCI.exists(fundGCI) Then
                dictFundGCI.Add fundGCI, data(GetSourceColumnIndex("IA GCI"), i)
            End If
        End If
    Next i
    
    ' Update Portfolio table
    If Not portfolioTable.DataBodyRange Is Nothing Then
        Dim portfolioData As Variant
        portfolioData = portfolioTable.DataBodyRange.Value
        
        ' Update Fund Manager GCI
        For i = 1 To UBound(portfolioData)
            If dictFundGCI.exists(portfolioData(i, GetColumnIndex(portfolioTable, "Fund GCI"))) Then
                portfolioData(i, GetColumnIndex(portfolioTable, "Fund Manager GCI")) = _
                    dictFundGCI(portfolioData(i, GetColumnIndex(portfolioTable, "Fund GCI")))
            Else
                portfolioData(i, GetColumnIndex(portfolioTable, "Fund Manager GCI")) = "No Match Found"
            End If
        Next i
        
        ' Write updates
        portfolioTable.DataBodyRange.Value = portfolioData
    End If
End Function

Private Sub ProcessNonTriggerData(data As Variant, portfolioTable As ListObject)
    ' Initialize arrays
    Dim numRows As Long
    numRows = UBound(data, 2)
    
    Dim processedData() As Variant
    ReDim processedData(1 To numRows, 1 To portfolioTable.ListColumns.Count)
    
    ' Track valid rows
    Dim validRows As Long
    validRows = 0
    
    ' Process each row
    Dim i As Long
    For i = 1 To numRows
        ' Skip FI-ASIA records
        If LCase(data(GetSourceColumnIndex("Region"), i)) <> "fi-asia" Then
            validRows = validRows + 1
            
            ' Map standard columns
            processedData(validRows, GetColumnIndex(portfolioTable, "Region")) = _
                data(GetSourceColumnIndex("Region"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Family")) = _
                data(GetSourceColumnIndex("Family"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Fund Manager GCI")) = _
                data(GetSourceColumnIndex("Fund Manager GCI"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Fund Manager")) = _
                data(GetSourceColumnIndex("Fund Manager"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Fund GCI")) = _
                data(GetSourceColumnIndex("Fund GCI"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Fund Name")) = _
                data(GetSourceColumnIndex("Fund Name"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Credit Officer")) = _
                data(GetSourceColumnIndex("Credit Officer"), i)
            processedData(validRows, GetColumnIndex(portfolioTable, "Wks Missing")) = _
                data(GetSourceColumnIndex("Weeks Missing"), i)
                
            ' Set Non-Trigger status
            processedData(validRows, GetColumnIndex(portfolioTable, "Trigger/Non-Trigger")) = "Non-Trigger"
        End If
    Next i
    
    ' Write valid data to table
    If validRows > 0 Then
        ' Resize array to remove empty rows
        Dim finalData() As Variant
        ReDim finalData(1 To validRows, 1 To portfolioTable.ListColumns.Count)
        
        ' Copy valid rows to final array
        Dim j As Long
        For i = 1 To validRows
            For j = 1 To portfolioTable.ListColumns.Count
                finalData(i, j) = processedData(i, j)
            Next j
        Next i
        
        ' Add to portfolio table
        portfolioTable.ListRows.Add
        portfolioTable.DataBodyRange.Resize(validRows, portfolioTable.ListColumns.Count).Value = finalData
    End If
End Sub

' === Helper Functions ===
Private Function GetDirectoryPath(filePath As String) As String
    GetDirectoryPath = Left(filePath, InStrRev(filePath, "\"))
End Function


Private Function GetFileName(filePath As String) As String
    GetFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
End Function

Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    On Error GoTo 0
    
    If GetColumnIndex = 0 Then
        MsgBox "Column '" & columnName & "' not found in table '" & tbl.Name & "'", vbExclamation
    End If
End Function

Private Function GetSourceColumnIndex(columnName As String) As Long
    ' Map source column names to indices
    Select Case LCase(columnName)
        Case "region": GetSourceColumnIndex = 1
        Case "family": GetSourceColumnIndex = 2
        Case "fund manager gci": GetSourceColumnIndex = 3
        Case "fund manager": GetSourceColumnIndex = 4
        Case "fund gci": GetSourceColumnIndex = 5
        Case "fund name": GetSourceColumnIndex = 6
        Case "credit officer": GetSourceColumnIndex = 7
        Case "wks missing", "weeks missing": GetSourceColumnIndex = 8
        Case "review status": GetSourceColumnIndex = 9
        Case "ia gci": GetSourceColumnIndex = 10
        Case Else: GetSourceColumnIndex = 0
    End Select
End Function

' === Macro Omega Implementation ===
Private Sub MacroOmega()
    On Error GoTo ErrorHandler
    
    ' Performance optimization
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' Setup worksheets and tables
    Dim wsPortfolio As Worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    Dim wsRepository As Worksheet
    Set wsRepository = ThisWorkbook.Sheets("Repository")
    
    Dim portfolioTable As ListObject
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    
    Dim repoTable As ListObject
    Set repoTable = wsRepository.ListObjects("Repo_DB")
    
    ' Verify tables exist
    If portfolioTable Is Nothing Then
        MsgBox "PortfolioTable not found!", vbCritical
        GoTo CleanUp
    End If
    
    If repoTable Is Nothing Then
        MsgBox "Repo_DB table not found!", vbCritical
        GoTo CleanUp
    End If
    
    ' Verify required columns exist
    If Not VerifyOmegaColumns(portfolioTable, repoTable) Then
        GoTo CleanUp
    End If
    
    ' Load data into arrays
    Dim portfolioData As Variant
    portfolioData = portfolioTable.DataBodyRange.Value
    
    Dim repoData As Variant
    repoData = repoTable.DataBodyRange.Value
    
    ' Create repository dictionary
    Dim repoDict As Object
    Set repoDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize progress
    InitializeProgress UBound(repoData, 1)
    
    ' Build repository dictionary
    Dim i As Long
    For i = 1 To UBound(repoData, 1)
        If Not IsEmpty(repoData(i, GetColumnIndex(repoTable, "Fund GCI"))) Then
            If Not repoDict.exists(repoData(i, GetColumnIndex(repoTable, "Fund GCI"))) Then
                repoDict.Add repoData(i, GetColumnIndex(repoTable, "Fund GCI")), _
                    Array( _
                        repoData(i, GetColumnIndex(repoTable, "NAV Source")), _
                        repoData(i, GetColumnIndex(repoTable, "Primary Client Contact")), _
                        repoData(i, GetColumnIndex(repoTable, "Secondary Client Contact")), _
                        repoData(i, GetColumnIndex(repoTable, "Chaser")) _
                    )
            End If
        End If
        UpdateProgress
    Next i
    
    ' Process portfolio updates
    If Not portfolioTable.DataBodyRange Is Nothing Then
        ' Initialize update array
        Dim updatedData() As Variant
        ReDim updatedData(1 To UBound(portfolioData, 1), 1 To 4)
        
        ' Initialize progress for portfolio processing
        InitializeProgress UBound(portfolioData, 1)
        
        ' Update portfolio data
        For i = 1 To UBound(portfolioData, 1)
            Dim fundGCI As Variant
            fundGCI = portfolioData(i, GetColumnIndex(portfolioTable, "Fund GCI"))
            
            If Not IsEmpty(fundGCI) Then
                If repoDict.exists(fundGCI) Then
                    ' Get repository data
                    Dim repoValues As Variant
                    repoValues = repoDict(fundGCI)
                    
                    ' Update values
                    updatedData(i, 1) = repoValues(0) ' NAV Source
                    updatedData(i, 2) = repoValues(1) ' Primary Client Contact
                    updatedData(i, 3) = repoValues(2) ' Secondary Client Contact
                    updatedData(i, 4) = repoValues(3) ' Chaser
                Else
                    ' Set default values for no match
                    updatedData(i, 1) = "No Match Found"
                    updatedData(i, 2) = "No Match Found"
                    updatedData(i, 3) = "No Match Found"
                    updatedData(i, 4) = "No Match Found"
                End If
            End If
            UpdateProgress
        Next i
        
        ' Write updates to portfolio table
        With portfolioTable
            .ListColumns("NAV Source").DataBodyRange.Value = _
                Application.Index(updatedData, , 1)
            .ListColumns("Primary Client Contact").DataBodyRange.Value = _
                Application.Index(updatedData, , 2)
            .ListColumns("Secondary Client Contact").DataBodyRange.Value = _
                Application.Index(updatedData, , 3)
            .ListColumns("Chaser").DataBodyRange.Value = _
                Application.Index(updatedData, , 4)
        End With
    End If
    
    MsgBox "Macro Omega completed successfully!", vbInformation
    
CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Error in Macro Omega: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' === Omega Helper Functions ===
Private Function VerifyOmegaColumns(portfolioTable As ListObject, repoTable As ListObject) As Boolean
    ' Define required columns for each table
    Dim portfolioColumns As Variant
    portfolioColumns = Array("Fund GCI", "NAV Source", "Primary Client Contact", _
                           "Secondary Client Contact", "Chaser")
                           
    Dim repoColumns As Variant
    repoColumns = Array("Fund GCI", "NAV Source", "Primary Client Contact", _
                       "Secondary Client Contact", "Chaser")
    
    ' Verify portfolio columns
    Dim col As Variant
    For Each col In portfolioColumns
        If GetColumnIndex(portfolioTable, CStr(col)) = 0 Then
            MsgBox "Required column '" & col & "' missing in PortfolioTable!", vbCritical
            VerifyOmegaColumns = False
            Exit Function
        End If
    Next col
    
    ' Verify repository columns
    For Each col In repoColumns
        If GetColumnIndex(repoTable, CStr(col)) = 0 Then
            MsgBox "Required column '" & col & "' missing in Repo_DB!", vbCritical
            VerifyOmegaColumns = False
            Exit Function
        End If
    Next col
    
    VerifyOmegaColumns = True
End Function

Private Sub EnsureOmegaColumns(tbl As ListObject)
    ' Define required columns
    Dim requiredColumns As Variant
    requiredColumns = Array("NAV Source", "Primary Client Contact", _
                          "Secondary Client Contact", "Chaser")
    
    ' Add missing columns
    Dim col As Variant
    For Each col In requiredColumns
        If GetColumnIndex(tbl, CStr(col)) = 0 Then
            tbl.ListColumns.Add.Name = CStr(col)
        End If
    Next col
End Sub

Private Function CreateRepoMapping(repoData As Variant, repoTable As ListObject) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(repoData, 1)
        Dim key As Variant
        key = repoData(i, GetColumnIndex(repoTable, "Fund GCI"))
        
        If Not IsEmpty(key) And Not dict.exists(key) Then
            dict.Add key, Array( _
                repoData(i, GetColumnIndex(repoTable, "NAV Source")), _
                repoData(i, GetColumnIndex(repoTable, "Primary Client Contact")), _
                repoData(i, GetColumnIndex(repoTable, "Secondary Client Contact")), _
                repoData(i, GetColumnIndex(repoTable, "Chaser")) _
            )
        End If
    Next i
    
    Set CreateRepoMapping = dict
End Function


' === Macro Gamma Implementation ===
Private Sub MacroGammaWithPath(previousFilePath As String)
    On Error GoTo ErrorHandler
    
    ' Performance optimization
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' Open previous workbook
    Dim wbPrevious As Workbook
    Set wbPrevious = Workbooks.Open(previousFilePath, ReadOnly:=True)
    
    ' Process each region
    Dim regions As Variant
    regions = Array("EMEA", "AMRS", "APAC")
    
    ' Initialize progress
    InitializeProgress UBound(regions) + 1
    
    Dim region As Variant
    For Each region In regions
        ProcessRegion CStr(region), ThisWorkbook, wbPrevious
        UpdateProgress
    Next region
    
    ' Cleanup
    wbPrevious.Close SaveChanges:=False
    
    MsgBox "Macro Gamma completed successfully!", vbInformation
    
CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Error in Macro Gamma: " & Err.Description, vbCritical
    If Not wbPrevious Is Nothing Then
        wbPrevious.Close SaveChanges:=False
    End If
    Resume CleanUp
End Sub

Private Sub ProcessRegion(region As String, currentWB As Workbook, previousWB As Workbook)
    ' Get worksheets and tables
    Dim wsCurrent As Worksheet
    Set wsCurrent = currentWB.Sheets(region)
    
    Dim wsPrevious As Worksheet
    Set wsPrevious = previousWB.Sheets(region)
    
    Dim tblCurrent As ListObject
    Set tblCurrent = wsCurrent.ListObjects(region)
    
    Dim tblPrevious As ListObject
    Set tblPrevious = wsPrevious.ListObjects(region)
    
    ' Verify required columns exist
    Dim requiredColumns As Variant
    requiredColumns = Array("Last Action", "Last Action Date", "Action Taker", _
                           "CPO/ECA", "Remediation Action Holder", "Comments", _
                           "Consolidated Comments", "Family")
                           
    If Not VerifyColumns(tblCurrent, requiredColumns) Then
        MsgBox "Required columns missing in current " & region & " table. Please check the structure.", vbCritical
        Exit Sub
    End If
    
    If Not VerifyColumns(tblPrevious, requiredColumns) Then
        MsgBox "Required columns missing in previous " & region & " table. Please check the structure.", vbCritical
        Exit Sub
    End If
    
    ' Load data into arrays
    Dim currentData As Variant
    currentData = tblCurrent.DataBodyRange.Value
    
    Dim previousData As Variant
    previousData = tblPrevious.DataBodyRange.Value
    
    ' Create lookup dictionary for previous data
    Dim dictPrevious As Object
    Set dictPrevious = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(previousData, 1)
        If Not IsEmpty(previousData(i, GetColumnIndex(tblPrevious, "Family"))) Then
            If Not dictPrevious.exists(previousData(i, GetColumnIndex(tblPrevious, "Family"))) Then
                dictPrevious.Add previousData(i, GetColumnIndex(tblPrevious, "Family")), i
            End If
        End If
    Next i
    
    ' Update current data
    For i = 1 To UBound(currentData, 1)
        Dim familyKey As Variant
        familyKey = currentData(i, GetColumnIndex(tblCurrent, "Family"))
        
        If Not IsEmpty(familyKey) Then
            If dictPrevious.exists(familyKey) Then
                UpdateColumns currentData, previousData, i, dictPrevious(familyKey), _
                             tblCurrent, tblPrevious
            End If
        End If
    Next i
    
    ' Write updates back to table
    tblCurrent.DataBodyRange.Value = currentData
    
    ' Apply data validation
    ApplyLastActionValidation tblCurrent
End Sub

Private Sub UpdateColumns(ByRef currentData As Variant, ByRef previousData As Variant, _
                        currentRow As Long, previousRow As Long, _
                        currentTbl As ListObject, previousTbl As ListObject)
    ' Define columns to update
    Dim columnsToUpdate As Variant
    columnsToUpdate = Array("Last Action", "Last Action Date", "Action Taker", _
                           "CPO/ECA", "Remediation Action Holder", "Comments", _
                           "Consolidated Comments")
    
    ' Update each column
    Dim col As Variant
    For Each col In columnsToUpdate
        Dim currentCol As Long, previousCol As Long
        currentCol = GetColumnIndex(currentTbl, CStr(col))
        previousCol = GetColumnIndex(previousTbl, CStr(col))
        
        If currentCol > 0 And previousCol > 0 Then
            ' Handle special case for Consolidated Comments
            If CStr(col) = "Consolidated Comments" Then
                Dim previousComment As String
                previousComment = GetNonEmptyValue(previousData(previousRow, previousCol))
                
                If Len(previousComment) > 0 Then
                    currentData(currentRow, currentCol) = previousComment
                End If
            Else
                ' Handle regular columns
                currentData(currentRow, currentCol) = previousData(previousRow, previousCol)
            End If
        End If
    Next col
End Sub

Private Function GetNonEmptyValue(value As Variant) As String
    If IsEmpty(value) Or IsNull(value) Then
        GetNonEmptyValue = ""
    Else
        GetNonEmptyValue = CStr(value)
    End If
End Function

Private Sub ApplyLastActionValidation(tbl As ListObject)
    On Error Resume Next
    
    With tbl.ListColumns("Last Action").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Pending,1st Outreach,2nd Outreach,1st Onshore Review," & _
                      "1st Escalation,2nd Escalation,3rd Escalation,Post-Update,In-Closing"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    On Error GoTo 0
End Sub

Private Function VerifyColumns(tbl As ListObject, requiredColumns As Variant) As Boolean
    VerifyColumns = True
    
    Dim col As Variant
    For Each col In requiredColumns
        If GetColumnIndex(tbl, CStr(col)) = 0 Then
            VerifyColumns = False
            Exit Function
        End If
    Next col
End Function

Private Sub EnsureRequiredColumns(tbl As ListObject)
    ' Define required columns
    Dim requiredColumns As Variant
    requiredColumns = Array("Last Action", "Last Action Date", "Action Taker", _
                           "CPO/ECA", "Remediation Action Holder", "Comments", _
                           "Consolidated Comments", "Family")
    
    ' Add missing columns
    Dim col As Variant
    For Each col In requiredColumns
        If GetColumnIndex(tbl, CStr(col)) = 0 Then
            tbl.ListColumns.Add.Name = CStr(col)
        End If
    Next col
End Sub


' === Table Initialization Functions ===
Private Sub InitializeTables()
    On Error GoTo ErrorHandler
    
    ' Initialize Portfolio table
    InitializePortfolioTable
    
    ' Initialize regional tables
    InitializeRegionalTables
    
    ' Initialize Repository table
    InitializeRepositoryTable
    
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing tables: " & Err.Description, vbCritical
    Err.Clear
End Sub

Private Sub InitializePortfolioTable()
    ' Get Portfolio worksheet
    Dim wsPortfolio As Worksheet
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    
    ' Ensure PortfolioTable exists
    Dim portfolioTable As ListObject
    Set portfolioTable = EnsureTableExists(wsPortfolio, "PortfolioTable")
    
    ' Define required columns
    Dim requiredColumns As Variant
    requiredColumns = Array("Region", "Family", "Fund Manager GCI", "Fund Manager", _
                           "Fund GCI", "Fund Name", "Wks Missing", "Credit Officer", _
                           "Trigger/Non-Trigger", "NAV Source", "Primary Client Contact", _
                           "Secondary Client Contact", "Chaser")
    
    ' Add missing columns
    AddMissingColumns portfolioTable, requiredColumns
End Sub

Private Sub InitializeRegionalTables()
    ' Process each region
    Dim regions As Variant
    regions = Array("EMEA", "AMRS", "APAC")
    
    Dim region As Variant
    For Each region In regions
        ' Get regional worksheet
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(CStr(region))
        
        ' Ensure regional table exists
        Dim tbl As ListObject
        Set tbl = EnsureTableExists(ws, CStr(region))
        
        ' Define required columns for regional tables
        Dim requiredColumns As Variant
        requiredColumns = Array("Family", "Last Action", "Last Action Date", _
                              "Action Taker", "CPO/ECA", "Remediation Action Holder", _
                              "Comments", "Consolidated Comments")
        
        ' Add missing columns
        AddMissingColumns tbl, requiredColumns
    Next region
End Sub

Private Sub InitializeRepositoryTable()
    ' Get Repository worksheet
    Dim wsRepository As Worksheet
    Set wsRepository = ThisWorkbook.Sheets("Repository")
    
    ' Ensure Repo_DB table exists
    Dim repoTable As ListObject
    Set repoTable = EnsureTableExists(wsRepository, "Repo_DB")
    
    ' Define required columns
    Dim requiredColumns As Variant
    requiredColumns = Array("Fund GCI", "NAV Source", "Primary Client Contact", _
                           "Secondary Client Contact", "Chaser")
    
    ' Add missing columns
    AddMissingColumns repoTable, requiredColumns
End Sub

' === Table Management Functions ===
Private Function EnsureTableExists(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set EnsureTableExists = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If EnsureTableExists Is Nothing Then
        ' Create new table if it doesn't exist
        Dim dataRange As Range
        Set dataRange = GetDataRange(ws)
        
        Set EnsureTableExists = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        EnsureTableExists.Name = tableName
        
        ' Apply basic formatting
        With EnsureTableExists
            .TableStyle = "TableStyleMedium2"
            .ShowTotals = False
            .ShowAutoFilter = True
        End With
    End If
End Function

Private Sub AddMissingColumns(tbl As ListObject, requiredColumns As Variant)
    Dim col As Variant
    For Each col In requiredColumns
        On Error Resume Next
        Dim idx As Long
        idx = tbl.ListColumns(CStr(col)).Index
        On Error GoTo 0
        
        If idx = 0 Then
            ' Add new column
            tbl.ListColumns.Add.Name = CStr(col)
            
            ' Apply specific formatting based on column type
            ApplyColumnFormatting tbl.ListColumns(CStr(col))
        End If
    Next col
End Sub

Private Sub ApplyColumnFormatting(col As ListColumn)
    Select Case col.Name
        Case "Last Action Date"
            col.DataBodyRange.NumberFormat = "dd/mm/yyyy"
        Case "Wks Missing"
            col.DataBodyRange.NumberFormat = "0"
        Case "Last Action"
            ApplyLastActionValidation col.Parent
    End Select
End Sub

Private Function GetDataRange(ws As Worksheet) As Range
    ' Get the used range
    Dim usedRange As Range
    Set usedRange = ws.UsedRange
    
    ' If sheet is empty, create a single cell range
    If usedRange.Cells.Count = 1 And IsEmpty(usedRange.Cells(1, 1)) Then
        Set GetDataRange = ws.Range("A1")
    Else
        ' Find the last used row and column
        Dim lastRow As Long, lastCol As Long
        lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        ' Set the range
        Set GetDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
End Function

' === Additional Utility Functions ===
Private Sub FormatAllTables()
    ' Format Portfolio table
    FormatTable ThisWorkbook.Sheets("Portfolio").ListObjects("PortfolioTable")
    
    ' Format regional tables
    Dim regions As Variant
    regions = Array("EMEA", "AMRS", "APAC")
    
    Dim region As Variant
    For Each region In regions
        FormatTable ThisWorkbook.Sheets(CStr(region)).ListObjects(CStr(region))
    Next region
    
    ' Format Repository table
    FormatTable ThisWorkbook.Sheets("Repository").ListObjects("Repo_DB")
End Sub

Private Sub FormatTable(tbl As ListObject)
    With tbl
        ' Apply table style
        .TableStyle = "TableStyleMedium2"
        
        ' Format header row
        With .HeaderRow
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)
        End With
        
        ' Auto-fit columns
        .Range.Columns.AutoFit
        
        ' Apply borders
        With .Range.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(191, 191, 191)
        End With
        
        ' Enable auto-filter
        .ShowAutoFilter = True
        
        ' Disable totals row
        .ShowTotals = False
    End With
End Sub

Private Sub CleanupWorkbook()
    ' Clear any filters
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Next ws
    
    ' Reset zoom level
    For Each ws In ThisWorkbook.Worksheets
        ws.Zoom = 100
    Next ws
    
    ' Clear clipboard
    Application.CutCopyMode = False
End Sub
