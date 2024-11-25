import os
import datetime
import pythoncom
import win32com.client as win32
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

EMAIL_LOG_FILE = "backup_email_log.xlsx"  # File to log email details

def create_search_macro(workbook):
    """
    Add a VBA macro to the Excel workbook to create a Search button.
    """
    vba_code = """
Sub SearchEmails()
    Dim searchValue As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim firstAddress As String
    Dim foundRange As Range
    Dim resultsSheet As Worksheet
    Dim rowNum As Integer

    ' Prompt for search value
    searchValue = InputBox("Enter the search term (Date, Email, or Subject):", "Search Emails")
    If searchValue = "" Then Exit Sub

    ' Set current worksheet
    Set ws = ThisWorkbook.Worksheets("Email Logs")

    ' Clear any previous highlights
    ws.Rows.Interior.ColorIndex = xlNone

    ' Search for the value in the Email Logs sheet
    Set foundRange = ws.UsedRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not foundRange Is Nothing Then
        ' Highlight matching cells
        firstAddress = foundRange.Address
        Do
            foundRange.EntireRow.Interior.Color = RGB(255, 255, 0) ' Highlight row in yellow
            Set foundRange = ws.UsedRange.FindNext(foundRange)
        Loop While Not foundRange Is Nothing And foundRange.Address <> firstAddress
    Else
        MsgBox "No matching records found for: " & searchValue, vbExclamation
    End If
End Sub
"""

    # Save VBA code in a file
    vba_file_path = os.path.join(os.getcwd(), "search_macro.vba")
    with open(vba_file_path, "w") as vba_file:
        vba_file.write(vba_code)

    print(f"VBA code saved to {vba_file_path}. Add it to the Excel file using the Excel Developer Tools.")
