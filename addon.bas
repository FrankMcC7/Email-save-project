Private Function GetManualDataFromCollection(ByVal manualData As Collection, ByVal gci As String) As Variant
    On Error Resume Next
    GetManualDataFromCollection = manualData.Item(gci)
    On Error GoTo 0
End Function
