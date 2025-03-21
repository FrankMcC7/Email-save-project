Sub UpdateOmegaFromRiskFile()
    Dim mainWs As Worksheet, updateWs As Worksheet
    Dim mainTable As ListObject, updateTable As ListObject
    Dim updateWb As Workbook
    Dim fd As FileDialog
    Dim filePath As String
    
    ' Set reference to the main worksheet "Risk Data" and table "Omega" in the current workbook.
    Set mainWs = ThisWorkbook.Sheets("Risk Data")
    Set mainTable = mainWs.ListObjects("Omega")
    
    ' Prompt the user to select the update file.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the update file (must contain sheet 'Risk Data' with a table)"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Exiting macro."
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    
    ' Open the update workbook.
    Set updateWb = Workbooks.Open(filePath)
    
    ' Check if the update workbook has a sheet named "Risk Data".
    On Error Resume Next
    Set updateWs = updateWb.Sheets("Risk Data")
    On Error GoTo 0
    If updateWs Is Nothing Then
        MsgBox "Sheet 'Risk Data' not found in the update file."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Use the first table available in the "Risk Data" sheet.
    If updateWs.ListObjects.Count = 0 Then
        MsgBox "No table found in sheet 'Risk Data' of the update file."
        updateWb.Close SaveChanges:=False
        Exit Sub
    Else
        Set updateTable = updateWs.ListObjects(1)
    End If
    
    ' Get column indexes for required columns in the update table.
    Dim colFundGCI_Upd As Long, colLeverage_Upd As Long, colLeverageTier_Upd As Long
    Dim colTransparencyTier_Upd As Long, colLiquidityTier_Upd As Long, colFundType_Upd As Long, colComments_Upd As Long
    On Error Resume Next
    colFundGCI_Upd = updateTable.ListColumns("Fund GCI").Index
    colLeverage_Upd = updateTable.ListColumns("Leverage").Index
    colLeverageTier_Upd = updateTable.ListColumns("Leverage Tier").Index
    colTransparencyTier_Upd = updateTable.ListColumns("Transparency Tier").Index
    colLiquidityTier_Upd = updateTable.ListColumns("Liquidity Tier").Index
    colFundType_Upd = updateTable.ListColumns("Fund Type").Index
    colComments_Upd = updateTable.ListColumns("Comments").Index
    On Error GoTo 0
    If colFundGCI_Upd = 0 Or colLeverage_Upd = 0 Or colLeverageTier_Upd = 0 Or colTransparencyTier_Upd = 0 _
        Or colLiquidityTier_Upd = 0 Or colFundType_Upd = 0 Or colComments_Upd = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'Leverage', 'Leverage Tier', 'Transparency Tier', " & _
               "'Liquidity Tier', 'Fund Type', 'Comments') not found in the update table."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Get column indexes for required columns in the main table "Omega".
    Dim colFundGCI_Omega As Long, colLeverage_Omega As Long, colLeverageTier_Omega As Long
    Dim colTransparencyTier_Omega As Long, colLiquidityTier_Omega As Long, colFundType_Omega As Long, colComments_Omega As Long
    On Error Resume Next
    colFundGCI_Omega = mainTable.ListColumns("Fund GCI").Index
    colLeverage_Omega = mainTable.ListColumns("Leverage").Index
    colLeverageTier_Omega = mainTable.ListColumns("Leverage Tier").Index
    colTransparencyTier_Omega = mainTable.ListColumns("Transparency Tier").Index
    colLiquidityTier_Omega = mainTable.ListColumns("Liquidity Tier").Index
    colFundType_Omega = mainTable.ListColumns("Fund Type").Index
    colComments_Omega = mainTable.ListColumns("Comments").Index
    On Error GoTo 0
    If colFundGCI_Omega = 0 Or colLeverage_Omega = 0 Or colLeverageTier_Omega = 0 Or colTransparencyTier_Omega = 0 _
        Or colLiquidityTier_Omega = 0 Or colFundType_Omega = 0 Or colComments_Omega = 0 Then
        MsgBox "One or more required columns ('Fund GCI', 'Leverage', 'Leverage Tier', 'Transparency Tier', " & _
               "'Liquidity Tier', 'Fund Type', 'Comments') not found in the main table 'Omega'."
        updateWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim updateRow As ListRow
    Dim fundGCIValue As Variant
    Dim leverageValue As Variant, leverageTierValue As Variant, transparencyTierValue As Variant
    Dim liquidityTierValue As Variant, fundTypeValue As Variant, commentsValue As Variant
    Dim foundCell As Range
    Dim rowIndex As Long
    
    ' Loop through each row in the update table.
    For Each updateRow In updateTable.ListRows
        fundGCIValue = updateRow.Range.Cells(1, colFundGCI_Upd).Value
        
        ' Find matching row in main table based on Fund GCI.
        Set foundCell = mainTable.DataBodyRange.Columns(colFundGCI_Omega).Find(What:=fundGCIValue, LookIn:=xlValues, LookAt:=xlWhole)
        If Not foundCell Is Nothing Then
            ' Calculate the relative row index within table Omega.
            rowIndex = foundCell.Row - mainTable.DataBodyRange.Rows(1).Row + 1
            
            ' Update only if the Leverage column in the main table is blank.
            If Trim(mainTable.DataBodyRange.Cells(rowIndex, colLeverage_Omega).Value & "") = "" Then
                leverageValue = updateRow.Range.Cells(1, colLeverage_Upd).Value
                leverageTierValue = updateRow.Range.Cells(1, colLeverageTier_Upd).Value
                transparencyTierValue = updateRow.Range.Cells(1, colTransparencyTier_Upd).Value
                liquidityTierValue = updateRow.Range.Cells(1, colLiquidityTier_Upd).Value
                fundTypeValue = updateRow.Range.Cells(1, colFundType_Upd).Value
                commentsValue = updateRow.Range.Cells(1, colComments_Upd).Value
                
                mainTable.DataBodyRange.Cells(rowIndex, colLeverage_Omega).Value = leverageValue
                mainTable.DataBodyRange.Cells(rowIndex, colLeverageTier_Omega).Value = leverageTierValue
                mainTable.DataBodyRange.Cells(rowIndex, colTransparencyTier_Omega).Value = transparencyTierValue
                mainTable.DataBodyRange.Cells(rowIndex, colLiquidityTier_Omega).Value = liquidityTierValue
                mainTable.DataBodyRange.Cells(rowIndex, colFundType_Omega).Value = fundTypeValue
                mainTable.DataBodyRange.Cells(rowIndex, colComments_Omega).Value = commentsValue
            End If
        End If
    Next updateRow
    
    ' Close the update workbook without saving changes.
    updateWb.Close SaveChanges:=False
    MsgBox "Update of table Omega completed successfully!"
End Sub
