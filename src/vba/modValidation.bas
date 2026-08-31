Attribute VB_Name = "modValidation"
Option Explicit

Private mIssueRow As Long
Private mErrorCount As Long
Private mWarningCount As Long

Public Sub KiemTraDuLieu()
    Dim isValid As Boolean
    isValid = ValidateWorkbook(True)
    If isValid Then
        MsgBox "Dữ liệu không có lỗi chặn xuất.", vbInformation, "Kiểm tra hoàn tất"
    Else
        MsgBox "Có lỗi cần sửa trước khi xuất. Xem sheet KiemTra.", _
               vbExclamation, "Chưa thể xuất"
    End If
End Sub

Public Function ValidateWorkbook(Optional ByVal showResults As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim wsCheck As Worksheet
    Dim idToRow As Object
    Dim idToLevel As Object
    Dim idToParent As Object
    Dim identitySeen As Object
    Dim rowIndex As Long
    Dim levelValue As Long
    Dim maxLevel As Long
    Dim personId As String
    Dim parentId As String
    Dim identityKey As String
    Dim key As Variant
    Dim currentId As String
    Dim visited As Object
    Dim capacity As Long
    Dim personCount As Long

    Set lo = PeopleTable()
    Set wsCheck = ThisWorkbook.Worksheets(SHEET_CHECK)
    Set idToRow = CreateObject("Scripting.Dictionary")
    Set idToLevel = CreateObject("Scripting.Dictionary")
    Set idToParent = CreateObject("Scripting.Dictionary")
    Set identitySeen = CreateObject("Scripting.Dictionary")

    NormalizeAllDates False
    PrepareCheckSheet wsCheck
    maxLevel = GetHangTKToiDa()
    If maxLevel < 0 Or maxLevel > 4 Then
        AddIssue "Lỗi", "CFG_SO_TANG", 0, vbNullString, _
                 "Hàng TK tối đa phải nằm trong khoảng 0 đến 4.", _
                 ThisWorkbook.Names("HangTKToiDa").RefersToRange
    End If

    If lo.DataBodyRange Is Nothing Then
        AddIssue "Lỗi", "DATA_EMPTY", 0, vbNullString, _
                 "Bảng người chưa có dòng dữ liệu.", Nothing
        GoTo FinishValidation
    End If

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(rowIndex) Then
            personCount = personCount + 1
            personId = Trim$(CStr(PersonCell(rowIndex, COL_ID).Value2))
            levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
            parentId = Trim$(CStr(PersonCell(rowIndex, COL_PARENT).Value2))

            If Len(personId) = 0 Then
                AddPersonIssue "Lỗi", "ID_MISSING", rowIndex, _
                               "Người này chưa có NguoiID."
            ElseIf idToRow.Exists(personId) Then
                AddPersonIssue "Lỗi", "ID_DUPLICATE", rowIndex, _
                               "NguoiID " & personId & " đang bị trùng."
            Else
                idToRow.Add personId, rowIndex
                idToLevel.Add personId, levelValue
                idToParent.Add personId, parentId
            End If

            If levelValue < 0 Or levelValue > maxLevel Then
                AddPersonIssue "Lỗi", "LEVEL_RANGE", rowIndex, _
                               "Hàng TK phải nằm trong H0–H" & CStr(maxLevel) & ".", COL_LEVEL
            End If

            If rowIndex <= 2 Then
                If levelValue <> 0 Or Not SafeBool(PersonCell(rowIndex, COL_OWNER).Value2) Then
                    AddPersonIssue "Lỗi", "OWNER_LEVEL", rowIndex, _
                                   "Hai vị trí đầu có dữ liệu phải là chủ đất ở Hàng TK 0.", COL_LEVEL
                End If
            ElseIf levelValue = 0 Then
                AddPersonIssue "Lỗi", "OWNER_POSITION", rowIndex, _
                               "Chỉ hai vị trí đầu được dùng Hàng TK 0 trong MVP.", COL_LEVEL
            End If

            If levelValue <= 1 Then
                If Len(parentId) > 0 Then
                    AddPersonIssue "Lỗi", "PARENT_NOT_EMPTY", rowIndex, _
                                   "Người ở Hàng TK 0/1 phải để trống ParentNguoiID.", COL_PARENT
                End If
            End If

            If HasEngineDate(rowIndex, COL_DEATH_CALC) And _
               SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2) Then
                AddPersonIssue "Lỗi", "DEAD_RECEIVE", rowIndex, _
                               "Người đã chết không thể đồng thời nhận đất.", COL_RECEIVE
            End If

            ValidatePersonDates rowIndex

            identityKey = NormalizeKey(PersonCell(rowIndex, COL_NAME).Value2) & "|" & _
                          EngineDateDisplay(rowIndex, COL_BIRTH_CALC)
            If Len(Replace$(identityKey, "|", vbNullString)) > 0 Then
                If identitySeen.Exists(identityKey) Then
                    AddPersonIssue "Cảnh báo", "PERSON_SIMILAR", rowIndex, _
                                   "Có người khác trùng họ tên và ngày sinh.", COL_NAME
                Else
                    identitySeen.Add identityKey, rowIndex
                End If
            End If
        ElseIf RowHasNonNameData(rowIndex) Then
            AddPersonIssue "Lỗi", "NAME_MISSING", rowIndex, _
                           "Dòng có dữ liệu nhưng chưa nhập họ tên.", COL_NAME
        End If
    Next rowIndex

    For Each key In idToRow.Keys
        rowIndex = CLng(idToRow(key))
        levelValue = CLng(idToLevel(key))
        parentId = CStr(idToParent(key))

        If levelValue > 1 Then
            If Len(parentId) = 0 Then
                AddPersonIssue "Lỗi", "PARENT_MISSING", rowIndex, _
                               "Người từ Hàng TK 2 trở lên chưa có người cha nhánh ở hàng ngay trên.", COL_PARENT
            ElseIf Not idToRow.Exists(parentId) Then
                AddPersonIssue "Lỗi", "PARENT_UNKNOWN", rowIndex, _
                               "ParentNguoiID không tồn tại trong bảng người.", COL_PARENT
            ElseIf CLng(idToLevel(parentId)) <> levelValue - 1 Then
                AddPersonIssue "Lỗi", "PARENT_LEVEL", rowIndex, _
                               "Người cha nhánh phải ở Hàng TK ngay phía trên.", COL_PARENT
            End If
        End If
    Next key

    For Each key In idToRow.Keys
        Set visited = CreateObject("Scripting.Dictionary")
        currentId = CStr(key)
        Do While Len(currentId) > 0 And idToParent.Exists(currentId)
            If visited.Exists(currentId) Then
                rowIndex = CLng(idToRow(key))
                AddPersonIssue "Lỗi", "PARENT_CYCLE", rowIndex, _
                               "Chuỗi ParentNguoiID đang tạo vòng tròn.", COL_PARENT
                Exit Do
            End If
            visited.Add currentId, True
            currentId = CStr(idToParent(currentId))
        Loop
    Next key

    For Each key In idToRow.Keys
        rowIndex = CLng(idToRow(key))
        If HasEngineDate(rowIndex, COL_DEATH_CALC) Then
            If Not HasDirectChild(CStr(key), idToParent) Then
                AddPersonIssue "Cảnh báo", "DEAD_WITHOUT_CHILD", rowIndex, _
                               "Nhánh có người đã chết nhưng chưa có người ở Hàng TK kế tiếp.", COL_DEATH
            End If
        End If
    Next key

    capacity = SafeLong(ThisWorkbook.Names("SucChuaNguoi").RefersToRange.Value2, 100)
    If capacity > 0 And personCount > capacity Then
        AddIssue "Lỗi", "TEMPLATE_CAPACITY", 0, vbNullString, _
                 "Có " & CStr(personCount) & " người nhưng template chỉ cho phép " & _
                 CStr(capacity) & " người.", ThisWorkbook.Worksheets(SHEET_INPUT).Range("B4")
    End If

    If Len(PersonName(2)) = 0 Then
        AddIssue "Cảnh báo", "SECOND_OWNER_EMPTY", 2, vbNullString, _
                 "Vị trí chủ đất thứ hai đang trống; hiểu là hồ sơ có một chủ đất.", _
                 PersonCell(2, COL_NAME)
    End If

    If HasRefusalWithoutGroup() Then
        AddIssue "Cảnh báo", "REFUSAL_DEFAULT_GROUP", 0, vbNullString, _
                 "Người từ chối chưa được chia nhóm; MVP sẽ dùng TC_DEFAULT.", Nothing
    End If

FinishValidation:
    wsCheck.Range("B2").Value2 = mErrorCount
    wsCheck.Range("B3").Value2 = mWarningCount
    ValidateWorkbook = (mErrorCount = 0)

    If showResults Then
        wsCheck.Visible = xlSheetVisible
        wsCheck.Activate
        wsCheck.Range("A1").Select
    End If
End Function

Private Sub PrepareCheckSheet(ByVal wsCheck As Worksheet)
    wsCheck.Range("A5:F1000").ClearContents
    On Error Resume Next
    wsCheck.Hyperlinks.Delete
    On Error GoTo 0

    wsCheck.Range("A4:F4").Value = Array("Mức", "Mã lỗi", "STT", "Họ tên", "Nội dung", "Ô cần sửa")
    mIssueRow = 5
    mErrorCount = 0
    mWarningCount = 0
End Sub

Private Sub AddPersonIssue(ByVal severity As String, ByVal issueCode As String, _
                           ByVal rowIndex As Long, ByVal message As String, _
                           Optional ByVal headerName As String = COL_NAME)
    AddIssue severity, issueCode, _
             SafeLong(PersonCell(rowIndex, COL_STT).Value2, rowIndex), _
             PersonName(rowIndex), message, PersonCell(rowIndex, headerName)
End Sub

Private Sub AddIssue(ByVal severity As String, ByVal issueCode As String, _
                     ByVal displayOrder As Long, ByVal displayName As String, _
                     ByVal message As String, ByVal targetCell As Range)
    Dim wsCheck As Worksheet
    Set wsCheck = ThisWorkbook.Worksheets(SHEET_CHECK)

    wsCheck.Cells(mIssueRow, 1).Value2 = severity
    wsCheck.Cells(mIssueRow, 2).Value2 = issueCode
    If displayOrder > 0 Then wsCheck.Cells(mIssueRow, 3).Value2 = displayOrder
    wsCheck.Cells(mIssueRow, 4).Value2 = displayName
    wsCheck.Cells(mIssueRow, 5).Value2 = message

    If Not targetCell Is Nothing Then
        wsCheck.Cells(mIssueRow, 6).Value2 = targetCell.Worksheet.Name & "!" & targetCell.Address(False, False)
        wsCheck.Hyperlinks.Add Anchor:=wsCheck.Cells(mIssueRow, 6), Address:="", _
                               SubAddress:="'" & targetCell.Worksheet.Name & "'!" & targetCell.Address, _
                               TextToDisplay:=targetCell.Worksheet.Name & "!" & targetCell.Address(False, False)
    End If

    If severity = "Lỗi" Then
        mErrorCount = mErrorCount + 1
    Else
        mWarningCount = mWarningCount + 1
    End If
    mIssueRow = mIssueRow + 1
End Sub

Private Function RowHasNonNameData(ByVal rowIndex As Long) As Boolean
    RowHasNonNameData = _
        Len(Trim$(CStr(PersonCell(rowIndex, COL_BIRTH).Value2))) > 0 Or _
        Len(Trim$(CStr(PersonCell(rowIndex, COL_DEATH).Value2))) > 0 Or _
        Len(Trim$(CStr(PersonCell(rowIndex, COL_DOCNO).Value2))) > 0 Or _
        Len(Trim$(CStr(PersonCell(rowIndex, COL_ISSUED).Value2))) > 0 Or _
        Len(Trim$(CStr(PersonCell(rowIndex, COL_ADDRESS).Value2))) > 0 Or _
        SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2)
End Function

Private Sub ValidatePersonDates(ByVal rowIndex As Long)
    If Len(Trim$(CStr(PersonCell(rowIndex, COL_BIRTH).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_BIRTH_CALC) Then
        AddPersonIssue "Lỗi", "DATE_BIRTH_INVALID", rowIndex, _
                       "Ngày sinh chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_BIRTH
    End If

    If Len(Trim$(CStr(PersonCell(rowIndex, COL_DEATH).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        AddPersonIssue "Lỗi", "DATE_DEATH_INVALID", rowIndex, _
                       "Ngày chết chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_DEATH
    End If

    If Len(Trim$(CStr(PersonCell(rowIndex, COL_ISSUED).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_ISSUED_CALC) Then
        AddPersonIssue "Lỗi", "DATE_ISSUED_INVALID", rowIndex, _
                       "Ngày cấp chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_ISSUED
    End If

    If HasEngineDate(rowIndex, COL_BIRTH_CALC) And _
       HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        If CDate(PersonCell(rowIndex, COL_DEATH_CALC).Value) < _
           CDate(PersonCell(rowIndex, COL_BIRTH_CALC).Value) Then
            AddPersonIssue "Lỗi", "DATE_DEATH_BEFORE_BIRTH", rowIndex, _
                           "Ngày chết không được trước ngày sinh.", COL_DEATH
        End If
    End If
End Sub

Private Function HasDirectChild(ByVal parentId As String, ByVal idToParent As Object) As Boolean
    Dim key As Variant
    For Each key In idToParent.Keys
        If CStr(idToParent(key)) = parentId Then
            HasDirectChild = True
            Exit Function
        End If
    Next key
End Function

Private Function HasRefusalWithoutGroup() As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(rowIndex) Then
            If Not HasEngineDate(rowIndex, COL_DEATH_CALC) And _
               Not SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2) And _
               Len(Trim$(CStr(PersonCell(rowIndex, COL_REFUSAL_GROUP).Value2))) = 0 Then
                HasRefusalWithoutGroup = True
                Exit Function
            End If
        End If
    Next rowIndex
End Function
