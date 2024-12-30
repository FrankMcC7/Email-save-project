Dim rowsNeeded As Long
    rowsNeeded = data.Count - 1  ' Subtract 1 because table starts with one row
    
    Dim k As Long
    For k = 1 To rowsNeeded
        iaTable.ListRows.Add
    Next k
