Sub MacroOmega()
    Dim wsPortfolio As Worksheet
    Dim wsRepository As Worksheet
    Dim portfolioTable As ListObject
    Dim repoTable As ListObject
    Dim dictRepo As Object
    Dim i As Long
    Dim numRowsPortfolio As Long
    Dim portfolioFundGCI As Variant
    Dim portfolioNAVSource As Variant
    Dim portfolioPrimaryContact As Variant
    Dim portfolioSecondaryContact As Variant
    Dim portfolioChaser As Variant
    Dim repoFundGCI As Variant
    Dim repoNAVSource As Variant
    Dim repoPrimaryContact As Variant
    Dim repoSecondaryContact As Variant
    Dim repoChaser As Variant
    Dim numRowsRepo As Long
    Dim key As String

    On Error GoTo ErrorHandler

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Set up the Portfolio sheet and table
    Set wsPortfolio = ThisWorkbook.Sheets("Portfolio")
    On Error Resume Next
    Set portfolioTable = wsPortfolio.ListObjects("PortfolioTable")
    On Error GoTo 0
    If portfolioTable Is Nothing Then
        MsgBox "PortfolioTable does not exist on the Portfolio sheet.", vbCritical
        GoTo CleanUp
    End If

    ' Set up the Repository sheet and table
    Set wsRepository = ThisWorkbook.Sheets("Repository")
    On Error Resume Next
    Set repoTable = wsRepository.ListObjects("Repo_DB")
    On Error GoTo 0
    If repoTable Is Nothing Then
        MsgBox "Repo_DB table does not exist on the Repository sheet.", vbCritical
        GoTo CleanUp
    End If

    ' Read Repo_DB data into arrays
    numRowsRepo = repoTable.DataBodyRange.Rows.Count
    If numRowsRepo = 0 Then
        MsgBox "Repo_DB table has no data.", vbInformation
        GoTo CleanUp
    End If
    repoFundGCI = repoTable.ListColumns("Fund GCI").DataBodyRange.Value
    repoNAVSource = repoTable.ListColumns("NAV Source").DataBodyRange.Value
    repoPrimaryContact = repoTable.ListColumns("Primary Client Contact").DataBodyRange.Value
    repoSecondaryContact = repoTable.ListColumns("Secondary Client Contact").DataBodyRange.Value
    repoChaser = repoTable.ListColumns("Chaser").DataBodyRange.Value

    ' Create a dictionary for Repo_DB using Fund GCI as key
    Set dictRepo = CreateObject("Scripting.Dictionary")
    For i = 1 To numRowsRepo
        key = CStr(repoFundGCI(i, 1))
        If Not dictRepo.Exists(key) Then
            dictRepo.Add key, Array(repoNAVSource(i, 1), repoPrimaryContact(i, 1), repoSecondaryContact(i, 1), repoChaser(i, 1))
        End If
    Next i

    ' Read Portfolio data into arrays
    numRowsPortfolio = portfolioTable.DataBodyRange.Rows.Count
    If numRowsPortfolio = 0 Then
        MsgBox "No data in PortfolioTable to update.", vbInformation
        GoTo CleanUp
    End If

    portfolioFundGCI = portfolioTable.ListColumns("Fund GCI").DataBodyRange.Value
    ReDim portfolioNAVSource(1 To numRowsPortfolio, 1 To 1)
    ReDim portfolioPrimaryContact(1 To numRowsPortfolio, 1 To 1)
    ReDim portfolioSecondaryContact(1 To numRowsPortfolio, 1 To 1)
    ReDim portfolioChaser(1 To numRowsPortfolio, 1 To 1)

    ' Populate the new columns in PortfolioTable
    For i = 1 To numRowsPortfolio
        key = CStr(portfolioFundGCI(i, 1))
        If dictRepo.Exists(key) Then
            portfolioNAVSource(i, 1) = dictRepo(key)(0)
            portfolioPrimaryContact(i, 1) = dictRepo(key)(1)
            portfolioSecondaryContact(i, 1) = dictRepo(key)(2)
            portfolioChaser(i, 1) = dictRepo(key)(3)
        Else
            portfolioNAVSource(i, 1) = "No Match Found"
            portfolioPrimaryContact(i, 1) = "No Match Found"
            portfolioSecondaryContact(i, 1) = "No Match Found"
            portfolioChaser(i, 1) = "No Match Found"
        End If
    Next i

    ' Write updated data back to PortfolioTable
    portfolioTable.ListColumns("NAV Source").DataBodyRange.Value = portfolioNAVSource
    portfolioTable.ListColumns("Primary Client Contact").DataBodyRange.Value = portfolioPrimaryContact
    portfolioTable.ListColumns("Secondary Client Contact").DataBodyRange.Value = portfolioSecondaryContact
    portfolioTable.ListColumns("Chaser").DataBodyRange.Value = portfolioChaser

    ' === Completion Message ===
    MsgBox "Macro Omega has successfully updated the PortfolioTable.", vbInformation

CleanUp:
    ' Reset Application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in Macro Omega: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
