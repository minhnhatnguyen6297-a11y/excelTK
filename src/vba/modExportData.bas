Attribute VB_Name = "modExportData"
Option Explicit

Public Sub BuildExportData()
    Dim wsExport As Worksheet
    Dim nextRow As Long

    Set wsExport = ThisWorkbook.Worksheets(SHEET_EXPORT)
    ' Vung AW:BK chua tblNguoiKyThuat. Chi xoa vung xuat trung gian A:V
    ' de khong lam mat du lieu ky thuat Nguoi.
    wsExport.Range("A:V").ClearContents

    wsExport.Range("A1").Value2 = "DỮ LIỆU XUẤT — KHÔNG NHẬP TAY"
    nextRow = 3
    nextRow = WriteGroupSection(wsExport, nextRow, "ChuDat")
    nextRow = WriteGroupSection(wsExport, nextRow, "NguoiDaChet")
    nextRow = WriteGroupSection(wsExport, nextRow, "NguoiNhanDat")
    nextRow = WriteGroupSection(wsExport, nextRow, "CacNhomTuChoi")
    nextRow = WriteGroupSection(wsExport, nextRow, "CayNhanh")
    RefreshTaiSanTable
    nextRow = WriteAssetSection(wsExport, nextRow)
    nextRow = WriteHoSoSection(wsExport, nextRow)

    wsExport.Columns("A:V").AutoFit
End Sub

Private Function WriteAssetSection(ByVal ws As Worksheet, ByVal startRow As Long) As Long
    Dim lo As ListObject, r As Long, c As Long
    Set lo = ThisWorkbook.Worksheets(SHEET_EXPORT).ListObjects(TABLE_ASSET_EXPORT)
    ws.Cells(startRow, 1).Value2 = "TaiSan"
    For c = 1 To lo.ListColumns.Count: ws.Cells(startRow + 1, c).Value2 = lo.ListColumns(c).Name: Next c
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.Count
            For c = 1 To lo.ListColumns.Count: ws.Cells(startRow + 1 + r, c).Value2 = lo.DataBodyRange.Cells(r, c).Value2: Next c
        Next r
        WriteAssetSection = startRow + 3 + lo.DataBodyRange.Rows.Count
    Else
        WriteAssetSection = startRow + 3
    End If
End Function

Private Function WriteHoSoSection(ByVal ws As Worksheet, ByVal startRow As Long) As Long
    Dim wsInput As Worksheet
    Dim profileRow As Long
    Dim outputColumn As Long
    Dim labelText As String
    Dim valueText As Variant

    Set wsInput = ThisWorkbook.Worksheets(SHEET_INPUT)
    ws.Cells(startRow, 1).Value2 = "HoSo"
    outputColumn = 1
    For profileRow = PROFILE_FIRST_DATA_ROW To ProfileLastDataRow()
        labelText = Trim$(ValueToExportText(wsInput.Cells(profileRow, 2).Value2))
        If Len(labelText) > 0 Then
            ws.Cells(startRow + 1, outputColumn).Value2 = ProfileFieldName(labelText)
            ws.Cells(startRow + 2, outputColumn).NumberFormat = "@"
            valueText = ValueToExportText(wsInput.Cells(profileRow, 3).Value2)
            ws.Cells(startRow + 2, outputColumn).Value2 = valueText
            outputColumn = outputColumn + 1
        End If
    Next profileRow
    WriteHoSoSection = startRow + 4
End Function

Private Function ProfileFieldName(ByVal labelText As String) As String
    Select Case NormalizeKey(labelText)
        Case NormalizeKey(UnicodeText("004E 0069 00EA 006D 0020 0079 1EBF 0074")): ProfileFieldName = "NiemYet"
        Case NormalizeKey(UnicodeText("0053 1ED1 0020 0063 00F4 006E 0067 0020 0063 0068 1EE9 006E 0067")): ProfileFieldName = "SoCongChung"
        Case NormalizeKey(UnicodeText("004E 0067 01B0 1EDD 0069 0020 1EE7 0079 0020 0071 0075 0079 1EC1 006E")): ProfileFieldName = "NguoiUyQuyen"
        Case NormalizeKey(UnicodeText("004E 0067 01B0 1EDD 0069 0020 1EE7 0079 0020 0071 0075 0079 1EC1 006E 0020 0032")): ProfileFieldName = "NguoiUyQuyen2"
        Case Else: ProfileFieldName = labelText
    End Select
End Function

Private Function SlotPersonRow(ByVal slot As Long) As Long
    Dim lo As ListObject, r As Long, i As Long, j As Long, ownerCount As Long, otherCount As Long
    Dim owners() As Long, others() As Long, swapRow As Long
    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(r) Then
            If SafeBool(PersonTechCell(r, COL_OWNER).Value2) Then ownerCount = ownerCount + 1 Else otherCount = otherCount + 1
        End If
    Next r
    If ownerCount > 0 Then ReDim owners(1 To ownerCount)
    If otherCount > 0 Then ReDim others(1 To otherCount)
    ownerCount = 0: otherCount = 0
    For r = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(r) Then
            If SafeBool(PersonTechCell(r, COL_OWNER).Value2) Then ownerCount = ownerCount + 1: owners(ownerCount) = r Else otherCount = otherCount + 1: others(otherCount) = r
        End If
    Next r
    For i = 1 To ownerCount - 1: For j = i + 1 To ownerCount: If SafeLong(PersonCell(owners(j), COL_STT).Value2, 0) < SafeLong(PersonCell(owners(i), COL_STT).Value2, 0) Then swapRow = owners(i): owners(i) = owners(j): owners(j) = swapRow
    Next j, i
    For i = 1 To otherCount - 1: For j = i + 1 To otherCount: If SafeLong(PersonCell(others(j), COL_STT).Value2, 0) < SafeLong(PersonCell(others(i), COL_STT).Value2, 0) Then swapRow = others(i): others(i) = others(j): others(j) = swapRow
    Next j, i
    If slot <= ownerCount And slot <= 2 Then SlotPersonRow = owners(slot)
    If slot > 2 And slot - 2 <= otherCount Then SlotPersonRow = others(slot - 2)
End Function

Public Function StaticTokenValue(ByVal tokenName As String, ByRef found As Boolean) As String
    Dim baseName As String
    Dim slot As Long
    Dim personMatches As Long
    Dim assetMatches As Long
    Dim profileMatches As Long
    Dim totalMatches As Long
    Dim resolvedValue As String

    found = False
    tokenName = LCase$(Trim$(tokenName))
    If Len(tokenName) = 0 Then Exit Function

    SplitTokenSlot tokenName, baseName, slot
    If slot > 0 Then
        personMatches = PersonStaticMatches(baseName, slot, resolvedValue)
        totalMatches = personMatches
        If personMatches = 0 Then resolvedValue = vbNullString

        assetMatches = AssetStaticMatches(baseName, slot, resolvedValue)
        totalMatches = totalMatches + assetMatches
    End If

    profileMatches = ProfileStaticMatches(tokenName, resolvedValue)
    totalMatches = totalMatches + profileMatches

    If totalMatches = 0 Then Exit Function
    found = True
    If totalMatches > 1 Then
        StaticTokenValue = "#TRUNG_TEN:" & tokenName & "#"
    Else
        StaticTokenValue = resolvedValue
    End If
End Function

Private Sub SplitTokenSlot(ByVal tokenName As String, ByRef baseName As String, ByRef slot As Long)
    Dim digitText As String
    Dim charIndex As Long
    Dim currentChar As String

    baseName = tokenName
    slot = 0
    For charIndex = Len(tokenName) To 1 Step -1
        currentChar = Mid$(tokenName, charIndex, 1)
        If currentChar >= "0" And currentChar <= "9" Then
            digitText = currentChar & digitText
        Else
            Exit For
        End If
    Next charIndex

    If Len(digitText) > 0 Then
        slot = SafeLong(digitText, 0)
        baseName = Left$(tokenName, Len(tokenName) - Len(digitText))
    End If
End Sub

Private Function PersonStaticMatches(ByVal baseName As String, ByVal slot As Long, _
                                     ByRef valueText As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim columnNumber As Long
    Dim extensionColumn As Long
    Dim rowIndex As Long
    Dim standardHeader As String
    Dim headerText As String

    If slot < 1 Then Exit Function
    Set lo = PeopleTable()
    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)

    standardHeader = PersonHeaderForToken(baseName)
    If Len(standardHeader) > 0 Then PersonStaticMatches = 1

    For columnNumber = PERSON_EXTENSION_FIRST_COLUMN To PERSON_EXTENSION_LAST_COLUMN
        headerText = NormalizeKey(ws.Cells(PERSON_EXTENSION_HEADER_ROW, columnNumber).Value2)
        If Len(headerText) > 0 And headerText = baseName Then
            PersonStaticMatches = PersonStaticMatches + 1
            extensionColumn = columnNumber
        End If
    Next columnNumber

    If PersonStaticMatches <> 1 Then Exit Function
    rowIndex = SlotPersonRow(slot)
    If rowIndex = 0 Then Exit Function
    If Len(standardHeader) > 0 Then
        valueText = ValueToExportText(PersonCell(rowIndex, standardHeader).Value2)
    ElseIf extensionColumn > 0 Then
        valueText = ValueToExportText(ws.Cells(lo.DataBodyRange.Row + rowIndex - 1, extensionColumn).Value2)
    End If
End Function

Private Function PersonHeaderForToken(ByVal baseName As String) As String
    Select Case baseName
        Case "ten": PersonHeaderForToken = COL_NAME
        Case "namsinh": PersonHeaderForToken = COL_BIRTH
        Case "namchet": PersonHeaderForToken = COL_DEATH
        Case "cccd": PersonHeaderForToken = COL_DOCNO
        Case "ngaycap": PersonHeaderForToken = COL_ISSUED
        Case "diachi": PersonHeaderForToken = COL_ADDRESS
        Case "loaicc": PersonHeaderForToken = COL_LOAI_CC
        Case "noicapcc": PersonHeaderForToken = COL_NOI_CAP_CC
        Case "thuongtru": PersonHeaderForToken = COL_NHAN_DIA_CHI
    End Select
End Function

Private Function AssetStaticMatches(ByVal baseName As String, ByVal slot As Long, _
                                    ByRef valueText As String) As Long
    Dim ws As Worksheet
    Dim fieldOffset As Long
    Dim matchedOffset As Long
    Dim labelText As String

    If slot < 1 Or slot > ASSET_CARD_COUNT Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    For fieldOffset = 1 To ASSET_FIELD_COUNT
        labelText = NormalizeKey(ws.Cells(ASSET_FIRST_CARD_ROW + fieldOffset, _
                                         ASSET_FIRST_CARD_COLUMN).Value2)
        If Len(labelText) > 0 And labelText = baseName Then
            AssetStaticMatches = AssetStaticMatches + 1
            matchedOffset = fieldOffset
        End If
    Next fieldOffset

    If AssetStaticMatches = 1 Then
        valueText = ValueToExportText(AssetValueCell(slot, matchedOffset).Value2)
    End If
End Function

Private Function ProfileStaticMatches(ByVal tokenName As String, ByRef valueText As String) As Long
    Dim ws As Worksheet
    Dim profileRow As Long
    Dim labelText As String
    Dim matchedRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    For profileRow = PROFILE_FIRST_DATA_ROW To ProfileLastDataRow()
        labelText = NormalizeKey(ws.Cells(profileRow, 2).Value2)
        If Len(labelText) > 0 And labelText = tokenName Then
            ProfileStaticMatches = ProfileStaticMatches + 1
            matchedRow = profileRow
        End If
    Next profileRow

    If ProfileStaticMatches = 1 Then
        valueText = ValueToExportText(ws.Cells(matchedRow, 3).Value2)
    End If
End Function

Private Function WriteGroupSection(ByVal wsExport As Worksheet, ByVal startRow As Long, _
                                   ByVal groupName As String) As Long
    Dim headers As Variant
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim groupId As String
    Dim columnIndex As Long

    headers = Array("Nhom", "NhomID", "STTNhap", "NguoiID", "HoTen", _
                    "NgaySinh", "NgaySinhGoc", "NgaySinhTinh", _
                    "NgayChet", "NgayChetGoc", "NgayChetTinh", "SoGiayTo", _
                    "NgayCap", "NgayCapGoc", "NgayCapTinh", "DiaChi", "HangTK", _
                    "ParentNguoiID", "LaChuDat", "NhanDat", "TrangThai", "TuoiLucChet")

    wsExport.Cells(startRow, 1).Value2 = groupName
    For columnIndex = LBound(headers) To UBound(headers)
        wsExport.Cells(startRow + 1, columnIndex + 1).Value2 = headers(columnIndex)
    Next columnIndex
    outputRow = startRow + 2

    Set lo = PeopleTable()
    If Not lo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To lo.DataBodyRange.Rows.Count
            If ShouldIncludePerson(rowIndex, groupName) Then
                groupId = vbNullString
                If groupName = "CacNhomTuChoi" Then
                    groupId = Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_REFUSAL_GROUP).Value2))
                    If Len(groupId) = 0 Then groupId = "TC_DEFAULT"
                End If
                WriteExportPerson wsExport, outputRow, groupName, groupId, rowIndex
                outputRow = outputRow + 1
            End If
        Next rowIndex
    End If

    If outputRow = startRow + 2 Then
        wsExport.Cells(outputRow, 1).Value2 = groupName
        wsExport.Cells(outputRow, 5).Value2 = "(không có dữ liệu)"
        outputRow = outputRow + 1
    End If

    WriteGroupSection = outputRow + 2
End Function

Private Function ShouldIncludePerson(ByVal rowIndex As Long, ByVal groupName As String) As Boolean
    Dim hasDeath As Boolean
    Dim receivesLand As Boolean

    If Not PersonHasData(rowIndex) Then Exit Function
    hasDeath = HasEngineDate(rowIndex, COL_DEATH_CALC)
    receivesLand = SafeBool(PersonCell(rowIndex, COL_RECEIVE).Value2)

    Select Case groupName
        Case "ChuDat"
            ShouldIncludePerson = SafeBool(PersonTechCell(rowIndex, COL_OWNER).Value2)
        Case "NguoiDaChet"
            ShouldIncludePerson = hasDeath
        Case "NguoiNhanDat"
            ShouldIncludePerson = (Not hasDeath And receivesLand)
        Case "CacNhomTuChoi"
            ShouldIncludePerson = (Not hasDeath And Not receivesLand)
        Case "CayNhanh"
            ShouldIncludePerson = True
    End Select
End Function

Private Sub WriteExportPerson(ByVal wsExport As Worksheet, ByVal outputRow As Long, _
                              ByVal groupName As String, ByVal groupId As String, _
                              ByVal personRow As Long)
    wsExport.Cells(outputRow, 1).Value2 = groupName
    wsExport.Cells(outputRow, 2).Value2 = groupId
    wsExport.Cells(outputRow, 3).Value2 = PersonCell(personRow, COL_STT).Value2
    wsExport.Cells(outputRow, 4).Value2 = PersonTechCell(personRow, COL_ID).Value2
    wsExport.Cells(outputRow, 5).Value2 = PersonCell(personRow, COL_NAME).Value2
    wsExport.Cells(outputRow, 6).NumberFormat = "@"
    wsExport.Cells(outputRow, 6).Value2 = PersonCell(personRow, COL_BIRTH).Value2
    wsExport.Cells(outputRow, 7).NumberFormat = "@"
    wsExport.Cells(outputRow, 7).Value2 = PersonTechCell(personRow, COL_BIRTH_RAW).Value2
    If HasEngineDate(personRow, COL_BIRTH_CALC) Then
        wsExport.Cells(outputRow, 8).Value2 = PersonTechCell(personRow, COL_BIRTH_CALC).Value2
        wsExport.Cells(outputRow, 8).NumberFormat = "dd/mm/yyyy"
    End If
    wsExport.Cells(outputRow, 9).NumberFormat = "@"
    wsExport.Cells(outputRow, 9).Value2 = PersonCell(personRow, COL_DEATH).Value2
    wsExport.Cells(outputRow, 10).NumberFormat = "@"
    wsExport.Cells(outputRow, 10).Value2 = PersonTechCell(personRow, COL_DEATH_RAW).Value2
    If HasEngineDate(personRow, COL_DEATH_CALC) Then
        wsExport.Cells(outputRow, 11).Value2 = PersonTechCell(personRow, COL_DEATH_CALC).Value2
        wsExport.Cells(outputRow, 11).NumberFormat = "dd/mm/yyyy"
    End If
    wsExport.Cells(outputRow, 12).Value2 = PersonCell(personRow, COL_DOCNO).Value2
    wsExport.Cells(outputRow, 13).NumberFormat = "@"
    wsExport.Cells(outputRow, 13).Value2 = PersonCell(personRow, COL_ISSUED).Value2
    wsExport.Cells(outputRow, 14).NumberFormat = "@"
    wsExport.Cells(outputRow, 14).Value2 = PersonTechCell(personRow, COL_ISSUED_RAW).Value2
    If HasEngineDate(personRow, COL_ISSUED_CALC) Then
        wsExport.Cells(outputRow, 15).Value2 = PersonTechCell(personRow, COL_ISSUED_CALC).Value2
        wsExport.Cells(outputRow, 15).NumberFormat = "dd/mm/yyyy"
    End If
    wsExport.Cells(outputRow, 16).Value2 = PersonCell(personRow, COL_ADDRESS).Value2
    wsExport.Cells(outputRow, 17).Value2 = PersonCell(personRow, COL_LEVEL).Value2
    wsExport.Cells(outputRow, 18).Value2 = PersonTechCell(personRow, COL_PARENT).Value2
    wsExport.Cells(outputRow, 19).Value2 = PersonTechCell(personRow, COL_OWNER).Value2
    wsExport.Cells(outputRow, 20).Value2 = SafeBool(PersonCell(personRow, COL_RECEIVE).Value2)
    wsExport.Cells(outputRow, 21).Value2 = PersonTechCell(personRow, COL_STATUS).Value2
    wsExport.Cells(outputRow, 22).Value2 = AgeAtDeath(personRow)
End Sub

Public Function GroupNameList(ByVal groupName As String) As String
    Dim names As Collection
    Dim lo As ListObject
    Dim rowIndex As Long

    Set names = New Collection
    Set lo = PeopleTable()

    If Not lo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To lo.DataBodyRange.Rows.Count
            If ShouldIncludePerson(rowIndex, groupName) Then names.Add PersonName(rowIndex)
        Next rowIndex
    End If

    GroupNameList = JoinNatural(names)
End Function

Public Function BranchSummary() As String
    Dim parts As Collection
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim levelValue As Long
    Dim parentName As String
    Dim itemText As String

    Set parts = New Collection
    Set lo = PeopleTable()

    If Not lo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To lo.DataBodyRange.Rows.Count
            If PersonHasData(rowIndex) Then
                levelValue = SafeLong(PersonCell(rowIndex, COL_LEVEL).Value2, 0)
                itemText = "Hàng TK " & CStr(levelValue) & ": " & PersonName(rowIndex)
                parentName = ParentPersonName(Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_PARENT).Value2)))
                If Len(parentName) > 0 Then itemText = itemText & " (thuộc nhánh " & parentName & ")"
                parts.Add itemText
            End If
        Next rowIndex
    End If

    BranchSummary = JoinCollectionWithSeparator(parts, "; ")
End Function

Public Function FirstPersonDate(ByVal displayHeader As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(rowIndex) Then
            FirstPersonDate = Trim$(ValueToExportText(PersonCell(rowIndex, displayHeader).Value2))
            Exit Function
        End If
    Next rowIndex
End Function

Public Function FirstPersonAgeAtDeath() As String
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim ageValue As Variant

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If PersonHasData(rowIndex) Then
            ageValue = AgeAtDeath(rowIndex)
            If Not IsEmpty(ageValue) Then FirstPersonAgeAtDeath = CStr(ageValue)
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ParentPersonName(ByVal parentId As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long

    If Len(parentId) = 0 Then Exit Function
    Set lo = PeopleTable()

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        If Trim$(ValueToExportText(PersonTechCell(rowIndex, COL_ID).Value2)) = parentId Then
            ParentPersonName = PersonName(rowIndex)
            Exit Function
        End If
    Next rowIndex
End Function

Private Function JoinNatural(ByVal items As Collection) As String
    Dim itemCount As Long
    itemCount = items.Count

    If itemCount = 0 Then
        JoinNatural = vbNullString
    ElseIf itemCount = 1 Then
        JoinNatural = CStr(items(1))
    ElseIf itemCount = 2 Then
        JoinNatural = CStr(items(1)) & " và " & CStr(items(2))
    Else
        JoinNatural = JoinCollectionRange(items, 1, itemCount - 1, ", ") & _
                      " và " & CStr(items(itemCount))
    End If
End Function

Private Function JoinCollectionRange(ByVal items As Collection, ByVal firstIndex As Long, _
                                     ByVal lastIndex As Long, ByVal separator As String) As String
    Dim index As Long
    Dim result As String

    For index = firstIndex To lastIndex
        If Len(result) > 0 Then result = result & separator
        result = result & CStr(items(index))
    Next index
    JoinCollectionRange = result
End Function

Private Function JoinCollectionWithSeparator(ByVal items As Collection, ByVal separator As String) As String
    If items.Count = 0 Then
        JoinCollectionWithSeparator = vbNullString
    Else
        JoinCollectionWithSeparator = JoinCollectionRange(items, 1, items.Count, separator)
    End If
End Function
