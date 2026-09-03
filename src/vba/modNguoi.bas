Attribute VB_Name = "modNguoi"
Option Explicit

Public Sub RefreshAllPeople()
    Dim lo As ListObject
    Dim rows As Range
    Dim rowIndex As Long
    Dim maxStt As Long
    Dim currentStt As Long
    Dim personId As String
    Dim levelValue As Long
    Dim previousLevel As Long

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set rows = lo.DataBodyRange

    NormalizeAllDates False

    For rowIndex = 1 To rows.Rows.Count
        If PersonHasData(rowIndex) Then
            currentStt = SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0)
            If currentStt > maxStt Then maxStt = currentStt
        End If
    Next rowIndex

    previousLevel = 1
    For rowIndex = 1 To rows.Rows.Count
        If PersonHasData(rowIndex) Then
            currentStt = SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0)
            If currentStt <= 0 Then
                maxStt = maxStt + 1
                PersonCell(rowIndex, COL_STT).Value2 = maxStt
            End If
            SyncPersonTechStt rowIndex
        End If

        If rowIndex <= 2 Then
            PersonCell(rowIndex, COL_LEVEL).Value2 = 0
            PersonTechCell(rowIndex, COL_PARENT).ClearContents
            PersonTechCell(rowIndex, COL_OWNER).Value2 = PersonHasData(rowIndex)
        ElseIf PersonHasData(rowIndex) Then
            levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
            If levelValue < 1 Or levelValue > 4 Then
                If previousLevel >= 1 And previousLevel <= GetHangTKToiDa() Then
                    levelValue = previousLevel
                Else
                    levelValue = 1
                End If
                PersonCell(rowIndex, COL_LEVEL).Value2 = levelValue
            End If
            previousLevel = levelValue
            PersonTechCell(rowIndex, COL_OWNER).Value2 = False
        Else
            If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) > 0 Then
                PersonTechCell(rowIndex, COL_ID).ClearContents
                PersonTechCell(rowIndex, COL_PARENT).ClearContents
                PersonTechCell(rowIndex, COL_OWNER).Value2 = False
                PersonTechCell(rowIndex, COL_REFUSAL_GROUP).ClearContents
                PersonTechCell(rowIndex, COL_STT).ClearContents
            End If
            PersonCell(rowIndex, COL_STT).ClearContents
            PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
            If rowIndex > 2 And SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0) < 1 Then
                PersonCell(rowIndex, COL_LEVEL).Value2 = previousLevel
            End If
        End If

        If PersonHasData(rowIndex) Then
            personId = Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_ID).Value2))
            If Len(personId) = 0 Then PersonTechCell(rowIndex, COL_ID).Value2 = NextPersonId()

            currentStt = SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0)
            If currentStt <= 0 Then
                maxStt = maxStt + 1
                PersonCell(rowIndex, COL_STT).Value2 = maxStt
            End If
            SyncPersonTechStt rowIndex
        End If
    Next rowIndex

    NormalizeAllDates False
    RefreshPersonExtensionColumns
    RebuildParentLinks
    UpdateAllStatuses
End Sub

Public Sub RefreshPersonExtensionColumns()
    Dim ws As Worksheet
    Dim columnNumber As Long
    Dim rowNumber As Long
    Dim headerText As String
    Dim cell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    For columnNumber = PERSON_EXTENSION_FIRST_COLUMN To PERSON_EXTENSION_LAST_COLUMN
        headerText = Trim$(ValueToExportText(ws.Cells(PERSON_EXTENSION_HEADER_ROW, columnNumber).Value2))
        ws.Columns(columnNumber).Hidden = (Len(headerText) = 0)
        For rowNumber = PERSON_EXTENSION_FIRST_ROW To PERSON_EXTENSION_LAST_ROW
            Set cell = ws.Cells(rowNumber, columnNumber)
            NormalizePersonExtensionCell cell
        Next rowNumber
    Next columnNumber
End Sub

Public Sub HandlePersonExtensionChange(ByVal changedRange As Range)
    Dim ws As Worksheet
    Dim affected As Range
    Dim oneCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    Set affected = Intersect(changedRange, PersonExtensionRange())
    If affected Is Nothing Then Exit Sub

    For Each oneCell In affected.Cells
        If oneCell.Row = PERSON_EXTENSION_HEADER_ROW Then
            ws.Columns(oneCell.Column).Hidden = (Len(Trim$(ValueToExportText(oneCell.Value2))) = 0)
        ElseIf oneCell.Row >= PERSON_EXTENSION_FIRST_ROW Then
            NormalizePersonExtensionCell oneCell
        End If
    Next oneCell
End Sub

Public Sub DongBoTruong()
    SyncPersonExtensionFields
End Sub

Public Sub SyncPersonExtensionFields()
    Dim ws As Worksheet
    Dim columnNumber As Long
    Dim rowNumber As Long
    Dim sourceRow As Long
    Dim headerText As String
    Dim sourceCell As Range
    Dim targetCell As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    For columnNumber = PERSON_EXTENSION_FIRST_COLUMN To PERSON_EXTENSION_LAST_COLUMN
        headerText = Trim$(ValueToExportText(ws.Cells(PERSON_EXTENSION_HEADER_ROW, columnNumber).Value2))
        If Len(headerText) > 0 Then
            sourceRow = 0
            For rowNumber = PERSON_EXTENSION_FIRST_ROW To PERSON_EXTENSION_LAST_ROW
                Set sourceCell = ws.Cells(rowNumber, columnNumber)
                NormalizePersonExtensionCell sourceCell
                If sourceCell.HasFormula Then
                    sourceRow = rowNumber
                    Exit For
                End If
            Next rowNumber

            If sourceRow > 0 Then
                Set sourceCell = ws.Cells(sourceRow, columnNumber)
                For rowNumber = sourceRow + 1 To PERSON_EXTENSION_LAST_ROW
                    Set targetCell = ws.Cells(rowNumber, columnNumber)
                    NormalizePersonExtensionCell targetCell
                    If targetCell.HasFormula Then
                        targetCell.NumberFormat = "General"
                    ElseIf Len(Trim$(ValueToExportText(targetCell.Value2))) = 0 Then
                        targetCell.FormulaR1C1 = sourceCell.FormulaR1C1
                        targetCell.NumberFormat = "General"
                    Else
                        targetCell.NumberFormat = "@"
                    End If
                Next rowNumber
            End If
        End If
    Next columnNumber

    RefreshPersonExtensionColumns
End Sub

Private Sub NormalizePersonExtensionCell(ByVal targetCell As Range)
    Dim enteredText As String
    Dim conversionFailed As Boolean

    If targetCell.HasFormula Then
        targetCell.NumberFormat = "General"
        Exit Sub
    End If

    enteredText = ValueToExportText(targetCell.Value2)
    If Left$(enteredText, 1) = "=" Then
        On Error Resume Next
        targetCell.NumberFormat = "General"
        targetCell.Formula = enteredText
        conversionFailed = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0
        If Not conversionFailed And targetCell.HasFormula Then Exit Sub
    End If
    targetCell.NumberFormat = "@"
End Sub

Public Sub RebuildParentLinks()
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim parentRow As Long
    Dim levelValue As Long

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If Not PersonHasData(rowIndex) Then
            If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) > 0 Then
                PersonTechCell(rowIndex, COL_PARENT).ClearContents
            End If
        Else
            levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
            If rowIndex <= 2 Or levelValue <= 1 Then
                PersonTechCell(rowIndex, COL_PARENT).ClearContents
            Else
                parentRow = FindPreviousParentRow(rowIndex, levelValue - 1)
                If parentRow > 0 Then
                    PersonTechCell(rowIndex, COL_PARENT).Value2 = PersonTechCell(parentRow, COL_ID).Value2
                Else
                    PersonTechCell(rowIndex, COL_PARENT).ClearContents
                End If
            End If
        End If
    Next rowIndex
End Sub

Public Function FindPreviousParentRow(ByVal childRow As Long, ByVal parentLevel As Long) As Long
    Dim rowIndex As Long
    For rowIndex = childRow - 1 To 1 Step -1
        If PersonHasData(rowIndex) Then
            If SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0) = parentLevel Then
                FindPreviousParentRow = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
    FindPreviousParentRow = 0
End Function

Public Sub UpdateAllStatuses()
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        UpdatePersonStatus rowIndex
    Next rowIndex
End Sub

Public Sub UpdatePersonStatus(ByVal rowIndex As Long)
    If Not PersonHasData(rowIndex) Then
        If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) > 0 Then
            PersonTechCell(rowIndex, COL_STATUS).ClearContents
        End If
    ElseIf HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        PersonTechCell(rowIndex, COL_STATUS).Value2 = "Đã chết"
    ElseIf SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2) Then
        PersonTechCell(rowIndex, COL_STATUS).Value2 = "Nhận đất"
    Else
        PersonTechCell(rowIndex, COL_STATUS).Value2 = "Từ chối"
    End If
End Sub

Public Sub HandlePeopleChange(ByVal changedRange As Range)
    Dim lo As ListObject
    Dim affected As Range
    Dim oneCell As Range
    Dim affectedRows As Object
    Dim rowKey As Variant
    Dim rowIndex As Long
    Dim deathWasChanged As Boolean
    Dim receiveWasChanged As Boolean

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set affected = Intersect(changedRange, lo.DataBodyRange)
    If affected Is Nothing Then Exit Sub
    Set affectedRows = CreateObject("Scripting.Dictionary")

    NormalizeChangedDateCells changedRange, True

    For Each oneCell In affected.Cells
        rowIndex = oneCell.Row - lo.DataBodyRange.Row + 1
        If rowIndex >= 1 And rowIndex <= lo.DataBodyRange.Rows.Count Then
            If Not affectedRows.Exists(CStr(rowIndex)) Then affectedRows.Add CStr(rowIndex), rowIndex
        End If
    Next oneCell

    For Each rowKey In affectedRows.Keys
        rowIndex = CLng(affectedRows(rowKey))
        InitializePersonRow rowIndex
        NormalizeChangedDateCells lo.DataBodyRange.Rows.Item(rowIndex), False

        deathWasChanged = Not Intersect(affected, PersonCell(rowIndex, COL_DEATH)) Is Nothing
        receiveWasChanged = Not Intersect(affected, PersonCell(rowIndex, COL_RECEIVE)) Is Nothing
        If PersonHasData(rowIndex) Then
            If Not Intersect(affected, PersonCell(rowIndex, COL_ISSUED)) Is Nothing Then
                UpdateGiayToForRow rowIndex
            End If
            If deathWasChanged Or receiveWasChanged Then
                ResolveDeathReceiveConflict rowIndex, deathWasChanged
            End If
        End If
        UpdatePersonStatus rowIndex
    Next rowKey

End Sub

Private Sub InitializePersonRow(ByVal rowIndex As Long)
    Dim levelValue As Long
    Dim previousLevel As Long
    Dim maxStt As Long
    Dim scanRow As Long

    If PersonHasData(rowIndex) And SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) <= 0 Then
        For scanRow = 1 To PeopleTable().DataBodyRange.Rows.Count
            maxStt = Application.Max(maxStt, SafeLong(PersonCell(scanRow, COL_STT).Value2, 0))
        Next scanRow
        PersonCell(rowIndex, COL_STT).Value2 = maxStt + 1
    End If
    If PersonHasData(rowIndex) Then SyncPersonTechStt rowIndex

    If rowIndex <= 2 Then
        PersonCell(rowIndex, COL_LEVEL).Value2 = 0
        PersonTechCell(rowIndex, COL_PARENT).ClearContents
        PersonTechCell(rowIndex, COL_OWNER).Value2 = PersonHasData(rowIndex)
    ElseIf PersonHasData(rowIndex) Then
        levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
        If levelValue < 1 Or levelValue > 4 Then
            previousLevel = PreviousBranchLevel(rowIndex)
            If previousLevel >= 1 And previousLevel <= GetHangTKToiDa() Then
                levelValue = previousLevel
            Else
                levelValue = 1
            End If
            PersonCell(rowIndex, COL_LEVEL).Value2 = levelValue
        End If
        PersonTechCell(rowIndex, COL_OWNER).Value2 = False
    Else
        If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) > 0 Then
            PersonTechCell(rowIndex, COL_ID).ClearContents
            PersonTechCell(rowIndex, COL_PARENT).ClearContents
            PersonTechCell(rowIndex, COL_OWNER).Value2 = False
            PersonTechCell(rowIndex, COL_REFUSAL_GROUP).ClearContents
            PersonTechCell(rowIndex, COL_STT).ClearContents
        End If
        PersonCell(rowIndex, COL_STT).ClearContents
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
        Exit Sub
    End If

    If Len(Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_ID).Value2))) = 0 Then
        PersonTechCell(rowIndex, COL_ID).Value2 = NextPersonId()
    End If

    If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) <= 0 Then
        For scanRow = 1 To PeopleTable().DataBodyRange.Rows.Count
            maxStt = Application.Max(maxStt, SafeLong(PersonCell(scanRow, COL_STT).Value2, 0))
        Next scanRow
        PersonCell(rowIndex, COL_STT).Value2 = maxStt + 1
    End If
    SyncPersonTechStt rowIndex
End Sub

Private Function PreviousBranchLevel(ByVal rowIndex As Long) As Long
    Dim scanRow As Long

    For scanRow = rowIndex - 1 To 3 Step -1
        If PersonHasData(scanRow) Then
            PreviousBranchLevel = SafeLong(PersonCell(scanRow, COL_LEVEL).Value2, 1)
            Exit Function
        End If
    Next scanRow
    PreviousBranchLevel = 1
End Function

Private Sub ResolveDeathReceiveConflict(ByVal rowIndex As Long, ByVal deathWasChanged As Boolean)
    Dim hasDeath As Boolean
    Dim receivesLand As Boolean
    Dim answer As VbMsgBoxResult

    hasDeath = HasEngineDate(rowIndex, COL_DEATH_CALC)
    receivesLand = SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2)
    If Not hasDeath Or Not receivesLand Then Exit Sub

    If deathWasChanged Then
        answer = MsgBox("Người đã chết không thể đồng thời nhận đất." & vbCrLf & _
                        "Chọn Có để giữ ngày chết và bỏ Nhận đất.", _
                        vbYesNo + vbExclamation, "Dữ liệu mâu thuẫn")
        If answer = vbYes Then
            PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
        Else
            PersonCell(rowIndex, COL_DEATH).ClearContents
        End If
    Else
        MsgBox "Người đã chết không thể nhận đất. Ô Nhận đất sẽ được bỏ chọn.", _
               vbExclamation, "Dữ liệu mâu thuẫn"
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
    End If
End Sub
