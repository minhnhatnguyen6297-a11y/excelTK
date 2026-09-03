Attribute VB_Name = "modValidation"
Option Explicit

Private mIssueRow As Long
Private mErrorCount As Long
Private mWarningCount As Long

Public Sub KiemTraDuLieu()
    Dim isValid As Boolean
    isValid = ValidateWorkbook(True)
    If isValid Then
        MsgBox "Không có lỗi nội dung dữ liệu.", vbInformation, "Kiểm tra hoàn tất"
    Else
        MsgBox "Có vấn đề dữ liệu để xem trong sheet KiemTra. Bạn vẫn có thể xuất văn bản.", _
               vbExclamation, "Kết quả chẩn đoán"
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
    Dim personCount As Long
    Dim assetIndex As Long
    Dim hasAssetData As Boolean

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
            personId = Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_ID).Value2))
            levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
            parentId = Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_PARENT).Value2))

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
                If levelValue <> 0 Or Not SafeBool(PersonTechCell(rowIndex, COL_OWNER).Value2) Then
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

    ValidateTemplateCapacity personCount

    If Len(PersonName(2)) = 0 Then
        AddIssue "Cảnh báo", "SECOND_OWNER_EMPTY", 2, vbNullString, _
                 "Vị trí chủ đất thứ hai đang trống; hiểu là hồ sơ có một chủ đất.", _
                 PersonCell(2, COL_NAME)
    End If

    If HasRefusalWithoutGroup() Then
        AddIssue "Cảnh báo", "REFUSAL_DEFAULT_GROUP", 0, vbNullString, _
                 "Người từ chối chưa được chia nhóm; MVP sẽ dùng TC_DEFAULT.", Nothing
    End If

    ValidateDuplicateLabels

    For assetIndex = 1 To ASSET_CARD_COUNT
        ValidateAssetCard assetIndex, hasAssetData
    Next assetIndex
    If Not hasAssetData Then
        AddIssue "Lỗi", "ASSET_NONE", 0, vbNullString, _
                 "Chưa có phiếu tài sản nào có dữ liệu.", AssetValueCell(1, ASSET_FIELD_LOAI_SO_OFFSET)
    End If
    If Not TaiSanHasData(1) And (TaiSanHasData(2) Or TaiSanHasData(3)) Then
        AddIssue "Cảnh báo", "ASSET_1_EMPTY", 1, vbNullString, _
                 "Phiếu TÀI SẢN 1 đang trống; hậu tố Word sẽ dồn lên.", AssetValueCell(1, ASSET_FIELD_LOAI_SO_OFFSET)
    End If
    If TemplateRequiresPlaceholder(13) Then
        If Len(Trim$(ValueToExportText(ThisWorkbook.Worksheets(SHEET_INPUT).Range("C41").Value2))) = 0 Then
            AddIssue "Lỗi", "NIEM_YET_MISSING", 0, vbNullString, _
                     "Mẫu đang chọn có placeholder Niêm Yết nhưng NiemYet đang trống.", ThisWorkbook.Worksheets(SHEET_INPUT).Range("C41")
        End If
    End If
    If TemplateRequiresPlaceholder(14) Then
        If Len(Trim$(ValueToExportText(ThisWorkbook.Worksheets(SHEET_INPUT).Range("C43").Value2))) = 0 Then
            AddIssue "Cảnh báo", "AUTHORITY_MISSING", 0, vbNullString, _
                     "Mẫu đang chọn có placeholder Người ủy quyền nhưng ô đang trống.", ThisWorkbook.Worksheets(SHEET_INPUT).Range("C43")
        End If
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

Private Sub ValidateDuplicateLabels()
    Dim wsInput As Worksheet
    Dim seen As Object
    Dim columnNumber As Long
    Dim profileRow As Long
    Dim labelText As String
    Dim normalizedLabel As String

    Set wsInput = ThisWorkbook.Worksheets(SHEET_INPUT)

    ' Người mở rộng: every column becomes the same token root for each
    ' person slot, so duplicate normalized headers are ambiguous.
    Set seen = CreateObject("Scripting.Dictionary")
    For columnNumber = PERSON_EXTENSION_FIRST_COLUMN To PERSON_EXTENSION_LAST_COLUMN
        labelText = Trim$(ValueToExportText(wsInput.Cells(PERSON_EXTENSION_HEADER_ROW, columnNumber).Value2))
        normalizedLabel = NormalizeKey(labelText)
        If Len(normalizedLabel) > 0 Then
            If seen.Exists(normalizedLabel) Then
                AddIssue "Cảnh báo", "LABEL_DUPLICATE", 0, vbNullString, _
                         "Tiêu đề Người mở rộng trùng sau chuẩn hóa: " & normalizedLabel & ".", _
                         wsInput.Cells(PERSON_EXTENSION_HEADER_ROW, columnNumber)
            Else
                seen.Add normalizedLabel, columnNumber
            End If
        End If
    Next columnNumber

    ' Hồ sơ has no slot suffix, therefore duplicate labels collide directly.
    Set seen = CreateObject("Scripting.Dictionary")
    For profileRow = PROFILE_FIRST_DATA_ROW To ProfileLastDataRow()
        labelText = Trim$(ValueToExportText(wsInput.Cells(profileRow, 2).Value2))
        normalizedLabel = NormalizeKey(labelText)
        If Len(normalizedLabel) > 0 Then
            If seen.Exists(normalizedLabel) Then
                AddIssue "Cảnh báo", "LABEL_DUPLICATE", 0, vbNullString, _
                         "Nhãn Hồ sơ trùng sau chuẩn hóa: " & normalizedLabel & ".", _
                         wsInput.Cells(profileRow, 2)
            Else
                seen.Add normalizedLabel, profileRow
            End If
        End If
    Next profileRow
End Sub

Private Sub ValidateTemplateCapacity(ByVal personCount As Long)
    Dim catalog As Worksheet
    Dim templateName As String
    Dim rowIndex As Long
    Dim peopleCapacity As Long
    Dim assetCapacity As Long
    Dim assetCount As Long
    Dim found As Boolean
    templateName = Dir$(GetConfigText("DuongDanMauWord"))
    Set catalog = ThisWorkbook.Worksheets("DanhMuc")
    For rowIndex = 4 To catalog.Cells(catalog.Rows.Count, 10).End(xlUp).Row
        If CStr(catalog.Cells(rowIndex, 10).Value2) = templateName Then
            peopleCapacity = SafeLong(catalog.Cells(rowIndex, 11).Value2, 0)
            assetCapacity = SafeLong(catalog.Cells(rowIndex, 12).Value2, 0)
            found = True
            Exit For
        End If
    Next rowIndex
    If Not found Or peopleCapacity <= 0 Or assetCapacity <= 0 Then
        AddIssue "Lỗi", "TEMPLATE_CAPACITY_UNKNOWN", 0, vbNullString, _
                 "Không xác định được sức chứa của mẫu đang chọn.", ThisWorkbook.Worksheets("CauHinh").Range("B8")
        Exit Sub
    End If
    assetCount = 0
    For rowIndex = 1 To ASSET_CARD_COUNT
        If TaiSanHasData(rowIndex) Then assetCount = assetCount + 1
    Next rowIndex
    If personCount > peopleCapacity Then
        AddIssue "Lỗi", "TEMPLATE_PEOPLE_CAPACITY", 0, vbNullString, _
                 "Mẫu chỉ chứa được " & CStr(peopleCapacity) & " người; hiện có " & CStr(personCount) & ".", _
                 ThisWorkbook.Worksheets(SHEET_INPUT).Range("B4")
    End If
    If assetCount > assetCapacity Then
        AddIssue "Lỗi", "TEMPLATE_ASSET_CAPACITY", 0, vbNullString, _
                 "Mẫu chỉ chứa được " & CStr(assetCapacity) & " phiếu tài sản; hiện có " & CStr(assetCount) & ".", _
                 AssetValueCell(1, ASSET_FIELD_LOAI_SO_OFFSET)
    End If
End Sub

Private Sub ValidateAssetCard(ByVal assetIndex As Long, ByRef hasAnyAsset As Boolean)
    Dim anyField As Boolean
    Dim area As Double
    Dim areaPresent As Boolean
    Dim componentTotal As Double
    Dim offset As Long
    Dim componentOffset As Variant
    Dim valueText As String
    Dim rawValue As Variant
    anyField = AssetFieldHasAnyValue(assetIndex)
    If Not anyField Then Exit Sub
    hasAnyAsset = True
    If Not TaiSanHasData(assetIndex) Then
        AddIssue "Cảnh báo", "ASSET_PARTIAL", assetIndex, vbNullString, _
                 "Phiếu TÀI SẢN " & CStr(assetIndex) & " có vài trường lẻ nhưng chưa đủ dữ liệu nhận diện; nghi nhập dở.", _
                 AssetValueCell(assetIndex, ASSET_FIELD_LOAI_SO_OFFSET)
        Exit Sub
    End If
    If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, 1).Value2))) = 0 Then AddAssetIssue assetIndex, "ASSET_LOAI_SO", "Thiếu Loại sổ.", 1
    If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, 4).Value2))) = 0 Then AddAssetIssue assetIndex, "ASSET_SO_THUA", "Thiếu Số thửa.", 4
    If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, 6).Value2))) = 0 Then AddAssetIssue assetIndex, "ASSET_DIA_CHI", "Thiếu Địa chỉ đất.", 6
    If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, 7).Value2))) = 0 Then AddAssetIssue assetIndex, "ASSET_DIEN_TICH", "Thiếu Diện tích.", 7
    If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, 9).Value2))) = 0 Then AddAssetIssue assetIndex, "ASSET_LOAI_DAT", "Thiếu Loại đất.", 9
    For Each componentOffset In Array(7, 11, 12, 13, 14)
        offset = CLng(componentOffset)
        rawValue = AssetValueCell(assetIndex, offset).Value2
        valueText = Trim$(ValueToExportText(rawValue))
        If Len(valueText) > 0 Then
            If IsError(rawValue) Then
                AddAssetIssue assetIndex, "ASSET_NOT_NUMBER", "Giá trị phải là số.", offset
            ElseIf VarType(rawValue) = vbString And Not IsNumeric(rawValue) Then
                AddAssetIssue assetIndex, "ASSET_NOT_NUMBER", "Giá trị phải là số.", offset
            ElseIf Not IsNumeric(rawValue) Then
                AddAssetIssue assetIndex, "ASSET_NOT_NUMBER", "Giá trị phải là số.", offset
            ElseIf CDbl(rawValue) < 0 Then
                AddAssetIssue assetIndex, "ASSET_NEGATIVE", "Giá trị không được âm.", offset
            End If
            If Not IsError(rawValue) And IsNumeric(rawValue) And CDbl(rawValue) >= 0 Then
                If offset = 7 Then
                    area = CDbl(rawValue)
                    areaPresent = True
                Else
                    componentTotal = componentTotal + CDbl(rawValue)
                End If
            End If
        End If
    Next componentOffset
    If areaPresent And componentTotal > area Then
        AddAssetIssue assetIndex, "ASSET_AREA_OVER", "Tổng ONT + CLN + NTS + LUC lớn hơn Diện tích.", 7
    ElseIf areaPresent And componentTotal < area Then
        AddAssetIssue assetIndex, "ASSET_AREA_UNALLOCATED", "Tổng diện tích thành phần nhỏ hơn Diện tích; còn phần chưa phân loại.", 7, "Cảnh báo"
    End If
End Sub

Private Function AssetFieldHasAnyValue(ByVal assetIndex As Long) As Boolean
    Dim offset As Long
    For offset = 1 To ASSET_FIELD_COUNT
        If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, offset).Value2))) > 0 Then AssetFieldHasAnyValue = True: Exit Function
    Next offset
End Function

Private Sub AddAssetIssue(ByVal assetIndex As Long, ByVal code As String, ByVal message As String, ByVal offset As Long, Optional ByVal severity As String = vbNullString)
    If Len(severity) = 0 Then severity = "Lỗi"
    AddIssue severity, code, assetIndex, vbNullString, "Phiếu TÀI SẢN " & CStr(assetIndex) & ": " & message, AssetValueCell(assetIndex, offset)
End Sub

Private Function TemplateRequiresPlaceholder(ByVal flagColumn As Long) As Boolean
    Dim catalog As Worksheet, templateName As String, rowIndex As Long
    templateName = Dir$(GetConfigText("DuongDanMauWord"))
    Set catalog = ThisWorkbook.Worksheets("DanhMuc")
    For rowIndex = 4 To catalog.Cells(catalog.Rows.Count, 10).End(xlUp).Row
        If CStr(catalog.Cells(rowIndex, 10).Value2) = templateName Then
            TemplateRequiresPlaceholder = CBool(catalog.Cells(rowIndex, flagColumn).Value2)
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub PrepareCheckSheet(ByVal wsCheck As Worksheet)
    wsCheck.Range("A5:F1000").ClearContents
    On Error Resume Next
    wsCheck.Hyperlinks.Delete
    On Error GoTo 0

    wsCheck.Range("A4:F4").Value2 = Array("Mức", "Mã lỗi", "STT", "Họ tên", "Nội dung", "Ô cần sửa")
    mIssueRow = 5
    mErrorCount = 0
    mWarningCount = 0
End Sub

Private Sub AddPersonIssue(ByVal severity As String, ByVal issueCode As String, _
                           ByVal rowIndex As Long, ByVal message As String, _
                           Optional ByVal headerName As String = COL_NAME)
    Dim targetCell As Range

    On Error Resume Next
    Set targetCell = PersonCell(rowIndex, headerName)
    On Error GoTo 0
    If targetCell Is Nothing Then Set targetCell = PersonTechCell(rowIndex, headerName)

    AddIssue severity, issueCode, _
             SafeLong(PersonCell(rowIndex, COL_STT).Value2, rowIndex), _
             PersonName(rowIndex), message, targetCell
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
        Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_BIRTH).Value2))) > 0 Or _
        Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_DEATH).Value2))) > 0 Or _
        Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_DOCNO).Value2))) > 0 Or _
        Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_ISSUED).Value2))) > 0 Or _
        Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_ADDRESS).Value2))) > 0 Or _
        SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2)
End Function

Private Sub ValidatePersonDates(ByVal rowIndex As Long)
    If Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_BIRTH).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_BIRTH_CALC) Then
        AddPersonIssue "Lỗi", "DATE_BIRTH_INVALID", rowIndex, _
                       "Ngày sinh chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_BIRTH
    End If

    If Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_DEATH).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        AddPersonIssue "Lỗi", "DATE_DEATH_INVALID", rowIndex, _
                       "Ngày chết chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_DEATH
    End If

    If Len(Trim$(ValueToExportText(PersonCell(rowIndex, COL_ISSUED).Value2))) > 0 And _
       Not HasEngineDate(rowIndex, COL_ISSUED_CALC) Then
        AddPersonIssue "Lỗi", "DATE_ISSUED_INVALID", rowIndex, _
                       "Ngày cấp chưa hợp lệ. Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", COL_ISSUED
    End If

    If HasEngineDate(rowIndex, COL_BIRTH_CALC) And _
       HasEngineDate(rowIndex, COL_DEATH_CALC) Then
        If CDate(PersonTechCell(rowIndex, COL_DEATH_CALC).Value2) < _
           CDate(PersonTechCell(rowIndex, COL_BIRTH_CALC).Value2) Then
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
               Len(Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_REFUSAL_GROUP).Value2))) = 0 Then
                HasRefusalWithoutGroup = True
                Exit Function
            End If
        End If
    Next rowIndex
End Function
