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
        If rowIndex <= 2 Then
            PersonCell(rowIndex, COL_LEVEL).Value2 = 0
            PersonCell(rowIndex, COL_PARENT).ClearContents
            PersonCell(rowIndex, COL_OWNER).Value2 = PersonHasData(rowIndex)
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
            PersonCell(rowIndex, COL_OWNER).Value2 = False
        Else
            PersonCell(rowIndex, COL_STT).ClearContents
            PersonCell(rowIndex, COL_ID).ClearContents
            PersonCell(rowIndex, COL_PARENT).ClearContents
            PersonCell(rowIndex, COL_OWNER).Value2 = False
            PersonCell(rowIndex, COL_REFUSAL_GROUP).ClearContents
            PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
            If rowIndex > 2 And SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0) < 1 Then
                PersonCell(rowIndex, COL_LEVEL).Value2 = previousLevel
            End If
        End If

        If PersonHasData(rowIndex) Then
            personId = Trim$(CStr(PersonCell(rowIndex, COL_ID).Value2))
            If Len(personId) = 0 Then PersonCell(rowIndex, COL_ID).Value2 = NextPersonId()

            currentStt = SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0)
            If currentStt <= 0 Then
                maxStt = maxStt + 1
                PersonCell(rowIndex, COL_STT).Value2 = maxStt
            End If
        End If
    Next rowIndex

    RebuildParentLinks
    UpdateAllStatuses
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
            PersonCell(rowIndex, COL_PARENT).ClearContents
        Else
            levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
            If rowIndex <= 2 Or levelValue <= 1 Then
                PersonCell(rowIndex, COL_PARENT).ClearContents
            Else
                parentRow = FindPreviousParentRow(rowIndex, levelValue - 1)
                If parentRow > 0 Then
                    PersonCell(rowIndex, COL_PARENT).Value2 = PersonCell(parentRow, COL_ID).Value2
                Else
                    PersonCell(rowIndex, COL_PARENT).ClearContents
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
        PersonCell(rowIndex, COL_STATUS).ClearContents
    ElseIf HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        PersonCell(rowIndex, COL_STATUS).Value2 = "Đã chết"
    ElseIf SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2) Then
        PersonCell(rowIndex, COL_STATUS).Value2 = "Nhận đất"
    Else
        PersonCell(rowIndex, COL_STATUS).Value2 = "Từ chối"
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
    Dim nameWasChanged As Boolean

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

    nameWasChanged = Not Intersect(affected, lo.ListColumns(COL_NAME).DataBodyRange) Is Nothing

    For Each rowKey In affectedRows.Keys
        rowIndex = CLng(affectedRows(rowKey))
        InitializePersonRow rowIndex

        deathWasChanged = Not Intersect(affected, PersonCell(rowIndex, COL_DEATH)) Is Nothing
        receiveWasChanged = Not Intersect(affected, PersonCell(rowIndex, COL_RECEIVE)) Is Nothing
        If deathWasChanged Or receiveWasChanged Then
            ResolveDeathReceiveConflict rowIndex, deathWasChanged
        End If

        UpdatePersonStatus rowIndex
    Next rowKey

    If nameWasChanged Then RebuildParentLinks
End Sub

Private Sub InitializePersonRow(ByVal rowIndex As Long)
    Dim levelValue As Long
    Dim previousLevel As Long
    Dim maxStt As Long
    Dim scanRow As Long

    If rowIndex <= 2 Then
        PersonCell(rowIndex, COL_LEVEL).Value2 = 0
        PersonCell(rowIndex, COL_PARENT).ClearContents
        PersonCell(rowIndex, COL_OWNER).Value2 = PersonHasData(rowIndex)
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
        PersonCell(rowIndex, COL_OWNER).Value2 = False
    Else
        PersonCell(rowIndex, COL_STT).ClearContents
        PersonCell(rowIndex, COL_ID).ClearContents
        PersonCell(rowIndex, COL_PARENT).ClearContents
        PersonCell(rowIndex, COL_OWNER).Value2 = False
        PersonCell(rowIndex, COL_REFUSAL_GROUP).ClearContents
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
        Exit Sub
    End If

    If Len(Trim$(CStr(PersonCell(rowIndex, COL_ID).Value2))) = 0 Then
        PersonCell(rowIndex, COL_ID).Value2 = NextPersonId()
    End If

    If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) <= 0 Then
        For scanRow = 1 To PeopleTable().DataBodyRange.Rows.Count
            maxStt = Application.Max(maxStt, SafeLong(PersonCell(scanRow, COL_STT).Value2, 0))
        Next scanRow
        PersonCell(rowIndex, COL_STT).Value2 = maxStt + 1
    End If
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
