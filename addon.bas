Private Sub FormatIATable(tbl As ListObject)
    With tbl
        ' Reset table style first
        .TableStyle = "TableStyleMedium2"
        
        ' Format entire table font
        With .Range
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Bold = False
            .WrapText = False  ' Ensure no text wrapping
        End With
        
        ' Format header row with better contrast
        With .HeaderRowRange
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)  ' White text
            .Interior.Color = RGB(68, 84, 106)  ' Dark blue background
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 30
        End With
        
        ' Format data body
        If Not .DataBodyRange Is Nothing Then
            With .DataBodyRange
                .Font.Bold = False
                .Interior.ColorIndex = xlNone
                .VerticalAlignment = xlCenter
                .WrapText = False  ' Ensure no text wrapping in data rows
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
    End With
End Sub
