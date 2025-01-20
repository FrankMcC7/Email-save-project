Option Explicit

Sub ProcessAllData()
    ' Variable declarations
    Dim wsPortfolio As Worksheet
    Dim wbMaster As Workbook
    Dim wbTrigger As Workbook, wbAllFunds As Workbook, wbNonTrigger As Workbook
    Dim triggerFile As String, allFundsFile As String, nonTriggerFile As String
    Dim triggerSheet As Worksheet, allFundsSheet As Worksheet, nonTriggerSheet As Worksheet
    Dim triggerTable As ListObject, portfolioTable As ListObject, allFundsTable As ListObject, nonTriggerTable As ListObject
    Dim dictFundGCI As Object, dictLatestNAVDate As Object
    Dim i As Long, j As Long
    Dim baseFolder As String
    Dim dataArray As Variant
    Dim triggerHeaders As Variant, nonTriggerHeaders As Variant
    Dim headerMapping As Object
    
    ' Additional Performance Optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' Define file paths - MODIFY THIS PATH TO MATCH YOUR ENVIRONMENT
    baseFolder = "C:\Data\NAV Reports\"  
    triggerFile = baseFolder & "Trigger.csv"
    allFundsFile = baseFolder & "All Funds.csv"
    nonTriggerFile = baseFolder & "Non-Trigger.csv"

    On Error GoTo ErrorHandler

    ' Verify file existence
    If Dir(triggerFile) = "" Or Dir(allFundsFile) = "" Or Dir(nonTriggerFile) = "" Then
        MsgBox "One or more required files not found. Please check paths.", vbCritical
        GoTo CleanUp
    End If

    ' Set up the Portfolio sheet
    Set wbMaster = ThisWorkbook
    Set wsPortfolio = wbMaster.Sheets("Portfolio")
    
    ' Create header mapping dictionary
    Set headerMapping = CreateObject("Scripting.Dictionary")
    With headerMapping
        .Add "Region", "Region"
        .Add "Fund Manager", "Fund Manager"
        .Add "Fund GCI", "Fund GCI"
        .Add "Fund Name", "Fund Name"
        .Add "Wks Missing", "Wks Missing"
        .Add "Credit Officer", "Credit Officer"
        .Add "Req NAV Date", "Required NAV Date"
        .Add "Latest NAV Date", "Latest NAV Date"
        .Add "Family", "Family"
    End With

    ' === Step 1: Set up Portfolio Table ===
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    
    If portfolioTable Is Nothing Then
        Set portfolioTable = wsPortfolio.ListObjects.Add(xlSrcRange, wsPortfolio.UsedRange, , xlYes)
        portfolioTable.Name = "PortfolioTable"
    End If

    ' Clear existing data and ensure required columns exist
    If Not portfolioTable.DataBodyRange Is Nothing Then portfolioTable.DataBodyRange.Delete
    
    ' Ensure all required columns exist
    Dim reqColumns As Variant
    reqColumns = Array("Latest NAV Date", "Required NAV Date", "Trigger/Non-Trigger")
    For i = LBound(reqColumns) To UBound(reqColumns)
        On Error Resume Next
        If portfolioTable.ListColumns(reqColumns(i)).Index = 0 Then
            portfolioTable.ListColumns.Add.Name = reqColumns(i)
        End If
        On Error GoTo 0
    Next i

    ' === Step 2: Process Trigger.csv using arrays ===
    Set wbTrigger = Workbooks.Open(triggerFile, UpdateLinks:=False, ReadOnly:=True)
    Set triggerSheet = wbTrigger.Sheets(1)
    
    ' Load entire trigger data into array for faster processing
    dataArray = triggerSheet.UsedRange.Value
    
    ' Process trigger data in memory
    Dim triggerData() As Variant
    ReDim triggerData(1 To UBound(dataArray, 1) - 1, 1 To UBound(dataArray, 2))
    
    ' Copy and process data in memory
    For i = 2 To UBound(dataArray, 1)
        For j = 1 To UBound(dataArray, 2)
            triggerData(i - 1, j) = dataArray(i, j)
            ' Process regions inline
            If j = 1 Then ' Assuming Region is first column
                Select Case triggerData(i - 1, j)
                    Case "US": triggerData(i - 1, j) = "AMRS"
                    Case "ASIA": triggerData(i - 1, j) = "APAC"
                End Select
            End If
        Next j
    Next i
    
    wbTrigger.Close SaveChanges:=False
    
    ' === Step 3: Process All Funds.csv using arrays ===
    Set wbAllFunds = Workbooks.Open(allFundsFile, UpdateLinks:=False, ReadOnly:=True)
    Set allFundsSheet = wbAllFunds.Sheets(1)
    
    ' Always delete the first row
    allFundsSheet.Rows(1).Delete
    
    ' Clean up the All Funds sheet
    With allFundsSheet
        
        ' Find the actual header row (the one with "Fund GCI", "IA GCI", etc.)
        Dim headerRow As Long
        headerRow = 1
        Do While headerRow <= .UsedRange.Rows.Count
            If Not IsEmpty(.Cells(headerRow, 1)) Then
                Dim headerFound As Boolean
                headerFound = False
                For i = 1 To .UsedRange.Columns.Count
                    If Trim(CStr(.Cells(headerRow, i).Value)) Like "*Fund GCI*" Or _
                       Trim(CStr(.Cells(headerRow, i).Value)) Like "*IA GCI*" Then
                        headerFound = True
                        Exit For
                    End If
                Next i
                If headerFound Then Exit Do
            End If
            headerRow = headerRow + 1
        Loop
        
        ' Delete any rows above the header row if needed
        If headerRow > 1 Then
            .Rows("1:" & headerRow - 1).Delete
        End If
        
        ' Clean up header row - trim spaces and remove any special characters
        For i = 1 To .UsedRange.Columns.Count
            Dim headerValue As String
            headerValue = Trim(.Cells(1, i).Value)
            headerValue = Replace(headerValue, vbLf, "")
            headerValue = Replace(headerValue, vbCr, "")
            headerValue = Replace(headerValue, vbNewLine, "")
            .Cells(1, i).Value = headerValue
        Next i
    End With
    
    ' Now load the cleaned data into memory
    Set dictFundGCI = CreateObject("Scripting.Dictionary")
    Set dictLatestNAVDate = CreateObject("Scripting.Dictionary")
    
    dataArray = allFundsSheet.UsedRange.Value
    
    ' Create quick lookup dictionaries with validation
    Dim fundGCICol As Long, iaGCICol As Long, latestNAVCol As Long, statusCol As Long
    fundGCICol = 0: iaGCICol = 0: latestNAVCol = 0: statusCol = 0
    
    ' Debug print the cleaned headers
    Debug.Print "All Funds Headers after cleaning:"
    For i = 1 To UBound(dataArray, 2)
        Debug.Print i & ": " & dataArray(1, i)
        Select Case Trim(CStr(dataArray(1, i)))
            Case "Fund GCI": fundGCICol = i
            Case "IA GCI": iaGCICol = i
            Case "Latest NAV Date": latestNAVCol = i
            Case "Review Status": statusCol = i
        End Select
    Next i
    
    ' Validate required columns were found
    If fundGCICol = 0 Or iaGCICol = 0 Or latestNAVCol = 0 Or statusCol = 0 Then
        MsgBox "Required columns not found in All Funds.csv:" & vbCrLf & _
               IIf(fundGCICol = 0, "- Fund GCI" & vbCrLf, "") & _
               IIf(iaGCICol = 0, "- IA GCI" & vbCrLf, "") & _
               IIf(latestNAVCol = 0, "- Latest NAV Date" & vbCrLf, "") & _
               IIf(statusCol = 0, "- Review Status", ""), vbCritical
        GoTo CleanUp
    End If
    
    Debug.Print "Column Indexes found:"
    Debug.Print "Fund GCI: " & fundGCICol
    Debug.Print "IA GCI: " & iaGCICol
    Debug.Print "Latest NAV Date: " & latestNAVCol
    Debug.Print "Review Status: " & statusCol
    
    ' Build dictionaries in memory with additional validation
    For i = 2 To UBound(dataArray, 1)
        If i <= UBound(dataArray, 1) And statusCol <= UBound(dataArray, 2) Then
            If Trim(CStr(dataArray(i, statusCol))) = "Approved" Then
                Dim fundGCI As Variant
                fundGCI = dataArray(i, fundGCICol)
                
                If Not IsEmpty(fundGCI) And Not IsNull(fundGCI) Then
                    If Not dictFundGCI.exists(fundGCI) Then
                        dictFundGCI.Add fundGCI, dataArray(i, iaGCICol)
                        dictLatestNAVDate.Add fundGCI, dataArray(i, latestNAVCol)
                    End If
                End If
            End If
        End If
    Next i
    
    wbAllFunds.Close SaveChanges:=False
    
    ' === Step 4: Process Non-Trigger.csv using arrays ===
    Set wbNonTrigger = Workbooks.Open(nonTriggerFile, UpdateLinks:=False, ReadOnly:=True)
    Set nonTriggerSheet = wbNonTrigger.Sheets(1)
    
    dataArray = nonTriggerSheet.UsedRange.Value
    
    ' Filter and process non-trigger data in memory
    Dim nonTriggerRows As Long
    nonTriggerRows = UBound(dataArray, 1) - 1
    
    Dim nonTriggerData() As Variant
    ReDim nonTriggerData(1 To nonTriggerRows, 1 To UBound(dataArray, 2))
    
    Dim validRow As Long: validRow = 0
    For i = 2 To UBound(dataArray, 1)
        If dataArray(i, 1) <> "FI-ASIA" Then ' Assuming Region is first column
            validRow = validRow + 1
            For j = 1 To UBound(dataArray, 2)
                nonTriggerData(validRow, j) = dataArray(i, j)
            Next j
        End If
    Next i
    
    wbNonTrigger.Close SaveChanges:=False
    
    ' === Step 5: Write all data to Portfolio table ===
    ' Combine trigger and non-trigger data
    Dim finalData() As Variant
    ReDim finalData(1 To UBound(triggerData, 1) + validRow, 1 To portfolioTable.ListColumns.Count)
    
    ' Copy trigger data
    For i = 1 To UBound(triggerData, 1)
        For j = 1 To UBound(triggerData, 2)
            finalData(i, j) = triggerData(i, j)
        Next j
        ' Set Trigger/Non-Trigger flag
        finalData(i, portfolioTable.ListColumns("Trigger/Non-Trigger").Index) = "Trigger"
    Next i
    
    ' Copy non-trigger data
    Dim currentRow As Long
    currentRow = UBound(triggerData, 1) + 1
    For i = 1 To validRow
        For j = 1 To UBound(nonTriggerData, 2)
            finalData(currentRow, j) = nonTriggerData(i, j)
        Next j
        finalData(currentRow, portfolioTable.ListColumns("Trigger/Non-Trigger").Index) = "Non-Trigger"
        currentRow = currentRow + 1
    Next i
    
    ' Write final data to sheet in one operation
    portfolioTable.ListObject.Range.Resize(UBound(finalData, 1) + 1, UBound(finalData, 2)).Value = finalData
    
    MsgBox "Data processed successfully!", vbInformation

CleanUp:
    ' Reset Application settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' Clean up objects
    Set wsPortfolio = Nothing
    Set wbMaster = Nothing
    Set wbTrigger = Nothing
    Set wbAllFunds = Nothing
    Set wbNonTrigger = Nothing
    Set headerMapping = Nothing
    Set dictFundGCI = Nothing
    Set dictLatestNAVDate = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
