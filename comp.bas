Option Explicit

' Declare API function for key state
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Sub AutomatedDataProcessing()
    Dim rawFilePath As String
    Dim wbRaw As Workbook
    Dim wsRaw As Worksheet
    Dim wsApproved As Worksheet
    Dim headerRange As Range
    Dim reviewStatusCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim i As Long, j As Long
    Dim sampleSheets As Collection
    Dim sampleSheet As Worksheet
    Dim approvedDataRows As Long
    Dim randIndices As Collection
    Dim index As Long
    Dim isRunning As Boolean
    Dim startTime As Double
    Dim fDialog As FileDialog
    Dim fileSelected As Boolean
    Dim fileExt As String
    
    ' Initialize
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup
    
    ' 1. File Upload and Import
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select Raw Data File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        fileSelected = .Show
        If fileSelected = False Then
            MsgBox "No file selected. Macro will exit.", vbExclamation
            GoTo Cleanup
        End If
        rawFilePath = .SelectedItems(1)
    End With
    
    ' Determine file extension
    fileExt = LCase(Mid(rawFilePath, InStrRev(rawFilePath, ".") + 1))
    
    ' Open or Import the raw data file
    If fileExt = "csv" Then
        ' Import CSV into a new worksheet
        Set wsRaw = ImportCSV(rawFilePath)
    Else
        ' Open Excel workbook and copy data
        Set wbRaw = Workbooks.Open(Filename:=rawFilePath, ReadOnly:=True)
        Set wsRaw = wbRaw.Sheets(1) ' Assuming data is in the first sheet
        wsRaw.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsRaw = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsRaw.Name = "RawData"
        wbRaw.Close SaveChanges:=False
    End If
    
    ' 2. Initial Data Cleanup
    ' Delete the first row
    wsRaw.Rows(1).Delete
    
    ' Remove all blank rows by checking each row for empty cells
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    
    ' Loop from bottom to top to delete blank rows
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(wsRaw.Rows(i)) = 0 Then
            wsRaw.Rows(i).Delete
        End If
    Next i
    
    ' 3. Filtering Process
    ' Locate "Review Status" column by searching the first row
    lastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column
    reviewStatusCol = 0
    For j = 1 To lastCol
        If Trim(LCase(wsRaw.Cells(1, j).Value)) = Trim(LCase("Review Status")) Then
            reviewStatusCol = j
            Exit For
        End If
    Next j
    
    ' Error handling for missing "Review Status" column
    If reviewStatusCol = 0 Then
        MsgBox """Review Status"" column not found.", vbCritical
        GoTo Cleanup
    End If
    
    ' Apply filter to keep only rows with "Approved" status
    With wsRaw
        .AutoFilterMode = False
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).AutoFilter Field:=reviewStatusCol, Criteria1:="Approved"
    End With
    
    ' Check if there are visible rows after filtering
    On Error Resume Next
    Set rng = wsRaw.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo Cleanup
    If rng Is Nothing Then
        MsgBox "No rows with ""Approved"" status found.", vbInformation
        GoTo Cleanup
    End If
    
    ' Create or overwrite "ApprovedData" sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("ApprovedData").Delete
    Application.DisplayAlerts = True
    On Error GoTo Cleanup
    
    Set wsApproved = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsApproved.Name = "ApprovedData"
    
    ' Copy filtered data to "ApprovedData"
    rng.Copy Destination:=wsApproved.Range("A1")
    
    ' Remove AutoFilter
    wsRaw.AutoFilterMode = False
    
    ' Optional: Delete RawData sheet to clean up
    Application.DisplayAlerts = False
    wsRaw.Delete
    Application.DisplayAlerts = True
    Set wsRaw = Nothing ' **Set wsRaw to Nothing after deletion**
    
    ' Check if there is data to process
    approvedDataRows = wsApproved.Cells(wsApproved.Rows.Count, 1).End(xlUp).Row
    If approvedDataRows < 2 Then
        MsgBox "No data available in ""ApprovedData"" after filtering.", vbInformation
        GoTo Cleanup
    End If
    
    ' 4. Random Sampling Loop
    isRunning = True
    Do While isRunning
        Set sampleSheets = New Collection
        
        ' Suppress alerts before deleting sample sheets
        Application.DisplayAlerts = False
        ' Create 5 sample sheets
        For i = 1 To 5
            On Error Resume Next
            Set sampleSheet = Worksheets("Sample" & i)
            If Not sampleSheet Is Nothing Then
                sampleSheet.Delete
            End If
            On Error GoTo 0
            
            Set sampleSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            sampleSheet.Name = "Sample" & i
            sampleSheets.Add sampleSheet
        Next i
        ' Restore alerts after deletion
        Application.DisplayAlerts = True
        
        ' Copy headers from "ApprovedData"
        Set headerRange = wsApproved.Rows(1)
        For Each sampleSheet In sampleSheets
            headerRange.Copy Destination:=sampleSheet.Rows(1)
        Next sampleSheet
        
        ' Prepare for random selection
        Set randIndices = New Collection
        Randomize
        For i = 2 To approvedDataRows ' Assuming row 1 is header
            randIndices.Add i
        Next i
        
        ' Shuffle the randIndices collection
        Call ShuffleCollection(randIndices)
        
        ' Select first 100 or less if not enough
        Dim totalSamples As Long
        totalSamples = Application.WorksheetFunction.Min(100, randIndices.Count - 1)
        
        ' Distribute samples to each sample sheet
        For Each sampleSheet In sampleSheets
            ' Optional: Clear previous data
            sampleSheet.Rows("2:" & sampleSheet.Rows.Count).ClearContents
            
            If randIndices.Count - 1 >= 100 Then
                For j = 1 To 100
                    index = randIndices(j)
                    wsApproved.Rows(index).Copy Destination:=sampleSheet.Cells(j + 1, 1)
                Next j
            Else
                For j = 1 To (randIndices.Count - 1)
                    index = randIndices(j)
                    wsApproved.Rows(index).Copy Destination:=sampleSheet.Cells(j + 1, 1)
                Next j
            End If
        Next sampleSheet
        
        ' 5. Cleanup and Repeat
        ' Wait for 5 seconds
        startTime = Timer
        Do While Timer < startTime + 5
            DoEvents
            ' Check if Esc key is pressed
            If GetAsyncKeyState(vbKeyEscape) <> 0 Then
                isRunning = False
                Exit Do
            End If
        Loop
        
        ' Suppress alerts before deleting sample sheets
        Application.DisplayAlerts = False
        ' Delete all sample sheets
        For Each sampleSheet In sampleSheets
            sampleSheet.Delete
        Next sampleSheet
        ' Restore alerts after deletion
        Application.DisplayAlerts = True
    Loop
    
    MsgBox "Macro execution stopped by user.", vbInformation

Cleanup:
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    If Not wsRaw Is Nothing Then wsRaw.AutoFilterMode = False
    If Not wsApproved Is Nothing Then wsApproved.AutoFilterMode = False
    On Error GoTo 0
End Sub

' Helper function to shuffle a collection
Sub ShuffleCollection(col As Collection)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim upper As Long
    upper = col.Count
    Randomize
    For i = upper To 2 Step -1
        j = Int((i - 1 + 1) * Rnd + 1) ' Random between 1 and i
        ' Swap items at positions i and j
        temp = col(i)
        col.Remove i
        col.Add temp, Before:=j
    Next i
End Sub

' Function to import CSV into a new worksheet
Function ImportCSV(filePath As String) As Worksheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "RawData"
    
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(1) ' General format
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    Set ImportCSV = ws
End Function
