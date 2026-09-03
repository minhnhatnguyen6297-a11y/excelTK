Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit
    Application.EnableEvents = False
    ApplyWorkbookProtection
    NormalizeProfileTextInputs ThisWorkbook.Worksheets(SHEET_INPUT).Range("B41:C1048576")
    NormalizeAllTaiSanInputs
    RefreshAllPeople
    RefreshWordSelectionDisplay
    ThisWorkbook.Worksheets(SHEET_INPUT).Activate
    ThisWorkbook.Worksheets(SHEET_INPUT).Range("B9").Select
SafeExit:
    Application.EnableEvents = True
End Sub
