Attribute VB_Name = "modHangTKNhanh"
Option Explicit

Public Sub HandleLevelClick(ByVal targetCell As Range)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim currentLevel As Long
    Dim nextLevel As Long
    Dim maxLevel As Long
    Dim lastDescendant As Long
    Dim checkRow As Long
    Dim shiftedLevel As Long
    Dim delta As Long
    Dim answer As VbMsgBoxResult

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(targetCell, lo.ListColumns(COL_LEVEL).DataBodyRange) Is Nothing Then Exit Sub

    rowIndex = targetCell.Row - lo.DataBodyRange.Row + 1
    rowCount = lo.DataBodyRange.Rows.Count

    If rowIndex <= 2 Then
        MsgBox "Hai vị trí đầu là chủ đất và luôn ở Hàng TK 0.", vbInformation, "Hàng TK chủ đất"
        GoTo MoveSelection
    End If

    If Not PersonHasData(rowIndex) Then
        MsgBox "Hãy nhập họ tên trước khi đổi Hàng TK.", vbInformation, "Chưa có người"
        GoTo MoveSelection
    End If

    maxLevel = GetHangTKToiDa()
    If maxLevel < 1 Then
        MsgBox "Hàng TK tối đa đang là 0 nên chỉ có thể nhập chủ đất.", _
               vbExclamation, "Không thể đổi Hàng TK"
        GoTo MoveSelection
    End If

    currentLevel = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 1)
    If currentLevel < 1 Or currentLevel > maxLevel Then currentLevel = 1

    If currentLevel >= maxLevel Then
        nextLevel = 1
    Else
        nextLevel = currentLevel + 1
    End If
    delta = nextLevel - currentLevel

    lastDescendant = rowIndex
    For checkRow = rowIndex + 1 To rowCount
        If Not PersonHasData(checkRow) Then Exit For
        If SafeLong(PersonCell(checkRow, COL_LEVEL).Value2, 0) <= currentLevel Then Exit For
        lastDescendant = checkRow
    Next checkRow

    If lastDescendant > rowIndex Then
        answer = MsgBox("Dòng này đang có " & CStr(lastDescendant - rowIndex) & _
                        " người ở các Hàng TK sau." & vbCrLf & _
                        "Đổi Hàng TK sẽ dịch chuyển cả khối. Tiếp tục?", _
                        vbYesNo + vbQuestion, "Đổi Hàng TK cả nhánh")
        If answer <> vbYes Then GoTo MoveSelection
    End If

    For checkRow = rowIndex + 1 To lastDescendant
        shiftedLevel = SafeLong(PersonCell(checkRow, COL_LEVEL).Value2, 0) + delta
        If shiftedLevel < 1 Or shiftedLevel > maxLevel Then
            MsgBox "Không thể đổi Hàng TK vì một người con sẽ vượt phạm vi H1–H" & _
                   CStr(maxLevel) & ".", vbExclamation, "Đã hủy thao tác"
            GoTo MoveSelection
        End If
    Next checkRow

    PersonCell(rowIndex, COL_LEVEL).Value2 = nextLevel
    For checkRow = rowIndex + 1 To lastDescendant
        PersonCell(checkRow, COL_LEVEL).Value2 = _
            SafeLong(PersonCell(checkRow, COL_LEVEL).Value2, 0) + delta
    Next checkRow

    RebuildParentLinks

MoveSelection:
    GoToPersonCell rowIndex, COL_NAME
End Sub

Public Sub HandleReceiveClick(ByVal targetCell As Range)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim hasDeath As Boolean

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(targetCell, lo.ListColumns(COL_RECEIVE).DataBodyRange) Is Nothing Then Exit Sub

    rowIndex = targetCell.Row - lo.DataBodyRange.Row + 1
    If Not PersonHasData(rowIndex) Then
        MsgBox "Hãy nhập họ tên trước khi chọn Nhận đất.", vbInformation, "Chưa có người"
        GoTo MoveSelection
    End If

    hasDeath = HasEngineDate(rowIndex, COL_DEATH_CALC)
    If hasDeath Then
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
        MsgBox "Người đã chết không thể nhận đất trong bản MVP.", _
               vbExclamation, "Không thể chọn"
    ElseIf SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2) Then
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 0
    Else
        PersonCell(rowIndex, COL_RECEIVE).Value2 = 1
    End If

    UpdatePersonStatus rowIndex

MoveSelection:
    GoToPersonCell rowIndex, COL_NAME
End Sub
