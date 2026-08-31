Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    ApplyWorkbookProtection
    RefreshAllPeople
    ThisWorkbook.Worksheets(SHEET_INPUT).Activate
    ThisWorkbook.Worksheets(SHEET_INPUT).Range("B9").Select
SafeExit:
    Application.EnableEvents = True
End Sub
