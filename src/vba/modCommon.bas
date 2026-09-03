Attribute VB_Name = "modCommon"
Option Explicit

Public Const SHEET_INPUT As String = "NhapLieu"
Public Const SHEET_CONFIG As String = "CauHinh"
Public Const SHEET_CHECK As String = "KiemTra"
Public Const SHEET_EXPORT As String = "XuatAn"
Public Const TABLE_PEOPLE As String = "tblNguoi"
Public Const TABLE_PEOPLE_TECH As String = "tblNguoiKyThuat"
Public Const TABLE_ASSET_EXPORT As String = "tblTaiSan"

Public Const ASSET_FIRST_CARD_ROW As Long = 8
Public Const ASSET_FIRST_CARD_COLUMN As Long = 27
Public Const ASSET_CARD_WIDTH As Long = 2
Public Const ASSET_CARD_COUNT As Long = 3
Public Const ASSET_FIELD_COUNT As Long = 18
Public Const ASSET_TECH_FIRST_ROW As Long = 41
Public Const ASSET_TECH_ROW_HEIGHT As Long = 20
' XuatAn!BM:BP; kept off NhapLieu while MIN-10 tblTaiSan deployment is paused.
Public Const ASSET_HIDDEN_ID_COLUMN As Long = 65
Public Const ASSET_HIDDEN_ISSUED_RAW_COLUMN As Long = 66
Public Const ASSET_HIDDEN_ISSUED_CALC_COLUMN As Long = 67
Public Const ASSET_HIDDEN_HAS_DATA_COLUMN As Long = 68
Public Const ASSET_FIELD_LOAI_SO_OFFSET As Long = 1
Public Const ASSET_FIELD_SERIAL_OFFSET As Long = 2
Public Const ASSET_FIELD_SO_THUA_OFFSET As Long = 4
Public Const ASSET_FIELD_DIA_CHI_DAT_OFFSET As Long = 6
Public Const ASSET_FIELD_NGAY_CAP_SO_OFFSET As Long = 16

Public Const PERSON_EXTENSION_FIRST_COLUMN As Long = 10
Public Const PERSON_EXTENSION_LAST_COLUMN As Long = 26
Public Const PERSON_EXTENSION_HEADER_ROW As Long = 8
Public Const PERSON_EXTENSION_FIRST_ROW As Long = 9
Public Const PERSON_EXTENSION_LAST_ROW As Long = 38

Public Const PROFILE_HEADER_ROW As Long = 40
Public Const PROFILE_FIRST_DATA_ROW As Long = 41
Public Const PROFILE_LAYOUT_CAP_ROW As Long = 70

Public Const COL_STT As String = "STTNhap"
Public Const COL_NAME As String = "HoTen"
Public Const COL_BIRTH As String = "NgaySinh"
Public Const COL_DEATH As String = "NgayChet"
Public Const COL_DOCNO As String = "SoGiayTo"
Public Const COL_ISSUED As String = "NgayCap"
Public Const COL_ADDRESS As String = "DiaChi"
Public Const COL_LEVEL As String = "HangTK"
Public Const COL_RECEIVE As String = "NhanDat"
Public Const COL_STATUS As String = "TrangThai"
Public Const COL_ID As String = "NguoiID"
Public Const COL_PARENT As String = "ParentNguoiID"
Public Const COL_OWNER As String = "LaChuDat"
Public Const COL_REFUSAL_GROUP As String = "NhomTuChoiID"
Public Const COL_BIRTH_RAW As String = "NgaySinhGoc"
Public Const COL_BIRTH_CALC As String = "NgaySinhTinh"
Public Const COL_DEATH_RAW As String = "NgayChetGoc"
Public Const COL_DEATH_CALC As String = "NgayChetTinh"
Public Const COL_ISSUED_RAW As String = "NgayCapGoc"
Public Const COL_ISSUED_CALC As String = "NgayCapTinh"
Public Const COL_LOAI_CC As String = "LoaiCC"
Public Const COL_NOI_CAP_CC As String = "NoiCapCC"
Public Const COL_NHAN_DIA_CHI As String = "NhanDiaChi"

Public Const PROTECTION_PASSWORD As String = "HoSoTK_MVP_2026"

Public Function PeopleTable() As ListObject
    Set PeopleTable = ThisWorkbook.Worksheets(SHEET_INPUT).ListObjects(TABLE_PEOPLE)
End Function

Public Function PeopleTechTable() As ListObject
    Set PeopleTechTable = ThisWorkbook.Worksheets(SHEET_EXPORT).ListObjects(TABLE_PEOPLE_TECH)
End Function

Public Function ColumnIndex(ByVal headerName As String) As Long
    ColumnIndex = PeopleTable.ListColumns(headerName).Index
End Function

Public Function PersonCell(ByVal rowIndex As Long, ByVal headerName As String) As Range
    Set PersonCell = PeopleTable.DataBodyRange.Cells(rowIndex, ColumnIndex(headerName))
End Function

Public Function PersonTechRowIndex(ByVal peopleRowIndex As Long) As Long
    Dim techTable As ListObject
    Dim inputStt As Long
    Dim techRowIndex As Long

    Set techTable = PeopleTechTable()
    inputStt = SafeLong(PersonCell(peopleRowIndex, COL_STT).Value2, 0)

    If inputStt > 0 Then
        For techRowIndex = 1 To techTable.DataBodyRange.Rows.Count
            If SafeLong(techTable.ListColumns(COL_STT).DataBodyRange.Cells(techRowIndex, 1).Value2, 0) = inputStt Then
                PersonTechRowIndex = techRowIndex
                Exit Function
            End If
        Next techRowIndex
    End If

    For techRowIndex = 1 To techTable.DataBodyRange.Rows.Count
        If Len(Trim$(ValueToExportText(techTable.ListColumns(COL_ID).DataBodyRange.Cells(techRowIndex, 1).Value2))) = 0 Then
            If inputStt > 0 Then
                techTable.ListColumns(COL_STT).DataBodyRange.Cells(techRowIndex, 1).Value2 = inputStt
            End If
            PersonTechRowIndex = techRowIndex
            Exit Function
    End If
    Next techRowIndex

    Err.Raise vbObjectError + 513, "PersonTechRowIndex", "Không tìm thấy dòng kỹ thuật Người."
End Function

Public Function PersonTechCell(ByVal rowIndex As Long, ByVal headerName As String) As Range
    Set PersonTechCell = PeopleTechTable.DataBodyRange.Cells( _
        PersonTechRowIndex(rowIndex), PeopleTechTable.ListColumns(headerName).Index)
End Function

Public Sub SyncPersonTechStt(ByVal rowIndex As Long)
    If SafeLong(PersonCell(rowIndex, COL_STT).Value2, 0) > 0 Then
        PersonTechCell(rowIndex, COL_STT).Value2 = PersonCell(rowIndex, COL_STT).Value2
    End If
End Sub

Public Function PersonName(ByVal rowIndex As Long) As String
    PersonName = Trim$(ValueToExportText(PersonCell(rowIndex, COL_NAME).Value2))
End Function

Public Function PersonHasData(ByVal rowIndex As Long) As Boolean
    PersonHasData = (Len(PersonName(rowIndex)) > 0)
End Function

Public Function PersonExtensionRange() As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    Set PersonExtensionRange = ws.Range(ws.Cells(PERSON_EXTENSION_HEADER_ROW, PERSON_EXTENSION_FIRST_COLUMN), _
                                        ws.Cells(PERSON_EXTENSION_LAST_ROW, PERSON_EXTENSION_LAST_COLUMN))
End Function

Public Function ProfileLastDataRow() As Long
    Dim ws As Worksheet
    Dim lastLabelRow As Long
    Dim lastValueRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    lastLabelRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastValueRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    ProfileLastDataRow = lastLabelRow
    If lastValueRow > ProfileLastDataRow Then ProfileLastDataRow = lastValueRow
    If ProfileLastDataRow < PROFILE_FIRST_DATA_ROW Then ProfileLastDataRow = PROFILE_FIRST_DATA_ROW
End Function

Public Function SafeLong(ByVal value As Variant, Optional ByVal fallback As Long = 0) As Long
    If IsError(value) Or IsEmpty(value) Then
        SafeLong = fallback
    ElseIf Len(Trim$(ValueToExportText(value))) = 0 Then
        SafeLong = fallback
    ElseIf IsNumeric(value) Then
        SafeLong = CLng(value)
    Else
        SafeLong = fallback
    End If
End Function

Public Function SafeBool(ByVal value As Variant) As Boolean
    Dim textValue As String
    If IsError(value) Or IsEmpty(value) Then
        SafeBool = False
    ElseIf VarType(value) = vbBoolean Then
        SafeBool = CBool(value)
    ElseIf IsNumeric(value) Then
        SafeBool = (CDbl(value) <> 0)
    Else
        textValue = LCase$(Trim$(ValueToExportText(value)))
        SafeBool = (textValue = "true" Or textValue = UnicodeText("0063 00F3"))
    End If
End Function

Public Function UnicodeText(ByVal hexCodes As String) As String
    Dim parts() As String
    Dim partIndex As Long
    Dim result As String

    If Len(hexCodes) = 0 Then Exit Function
    parts = Split(hexCodes, " ")

    For partIndex = LBound(parts) To UBound(parts)
        result = result & ChrW$(CLng("&H" & parts(partIndex)))
    Next partIndex

    UnicodeText = result
End Function

Public Function GetHangTKToiDa() As Long
    GetHangTKToiDa = SafeLong(ThisWorkbook.Names("HangTKToiDa").RefersToRange.Value2, 2)
End Function

Public Function GetConfigText(ByVal rangeName As String) As String
    On Error GoTo MissingName
    GetConfigText = Trim$(ValueToExportText(ThisWorkbook.Names(rangeName).RefersToRange.Value2))
    Exit Function
MissingName:
    GetConfigText = vbNullString
End Function

Public Function ValueToExportText(ByVal value As Variant) As String
    Dim errorText As String

    If IsError(value) Then
        On Error Resume Next
        errorText = CStr(value)
        On Error GoTo 0
        Select Case errorText
            Case "Error 2000": ValueToExportText = "#NULL!"
            Case "Error 2007": ValueToExportText = "#DIV/0!"
            Case "Error 2015": ValueToExportText = "#VALUE!"
            Case "Error 2023": ValueToExportText = "#REF!"
            Case "Error 2029": ValueToExportText = "#NAME?"
            Case "Error 2036": ValueToExportText = "#NUM!"
            Case "Error 2042": ValueToExportText = "#N/A"
            Case Else: ValueToExportText = errorText
        End Select
    ElseIf IsEmpty(value) Then
        ValueToExportText = vbNullString
    Else
        ValueToExportText = CStr(value)
    End If
End Function

Public Sub NormalizeInputTextCell(ByVal inputCell As Range)
    ' MIN-24 formula behavior is still pending. Keep the input contract
    ' (Text format) without calculating, rewriting, or promoting formulas.
    inputCell.NumberFormat = "@"
End Sub

Public Sub NormalizeProfileTextInputs(ByVal changedRange As Range)
    Dim affected As Range

    Set affected = Intersect(changedRange, ThisWorkbook.Worksheets(SHEET_INPUT).Range("B41:C1048576"))
    If affected Is Nothing Then Exit Sub
    ' One range assignment keeps the open/change path cheap even though the
    ' profile area is intentionally open-ended.
    affected.NumberFormat = "@"
End Sub

Public Sub SetConfigText(ByVal rangeName As String, ByVal valueText As String)
    On Error GoTo MissingName
    With ThisWorkbook.Names(rangeName).RefersToRange
        .NumberFormat = "@"
        .Value2 = valueText
    End With
    Exit Sub
MissingName:
    Err.Raise vbObjectError + 514, "SetConfigText", "Không tìm thấy cấu hình " & rangeName & "."
End Sub

Public Sub SetConfigNumber(ByVal rangeName As String, ByVal valueNumber As Long)
    On Error GoTo MissingName
    ThisWorkbook.Names(rangeName).RefersToRange.Value2 = valueNumber
    Exit Sub
MissingName:
    Err.Raise vbObjectError + 515, "SetConfigNumber", "Không tìm thấy cấu hình " & rangeName & "."
End Sub

Public Function ShortPathText(ByVal fullPath As String, Optional ByVal maxLength As Long = 48) As String
    Dim prefix As String
    Dim tail As String
    If Len(fullPath) <= maxLength Or maxLength < 8 Then
        ShortPathText = fullPath
        Exit Function
    End If
    prefix = Left$(fullPath, 3)
    tail = Mid$(fullPath, Len(fullPath) - maxLength + 7)
    ShortPathText = prefix & "..." & tail
End Function

Public Function NextPersonId() As String
    Dim counterCell As Range
    Dim nextNumber As Long

    Set counterCell = ThisWorkbook.Names("NguoiIDTiepTheo").RefersToRange
    nextNumber = SafeLong(counterCell.Value2, 1)
    If nextNumber < 1 Then nextNumber = 1

    NextPersonId = "P" & Format$(nextNumber, "0000")
    counterCell.Value2 = nextNumber + 1
End Function

Public Function AssetCardStartRow(ByVal assetIndex As Long) As Long
    If assetIndex < 1 Or assetIndex > ASSET_CARD_COUNT Then Exit Function
    AssetCardStartRow = ASSET_FIRST_CARD_ROW
End Function

Public Function AssetCardStartColumn(ByVal assetIndex As Long) As Long
    If assetIndex < 1 Or assetIndex > ASSET_CARD_COUNT Then Exit Function
    AssetCardStartColumn = ASSET_FIRST_CARD_COLUMN + ((assetIndex - 1) * ASSET_CARD_WIDTH)
End Function

Public Function AssetValueCell(ByVal assetIndex As Long, ByVal fieldOffset As Long) As Range
    Set AssetValueCell = ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item( _
        AssetCardStartRow(assetIndex) + fieldOffset, AssetCardStartColumn(assetIndex) + 1)
End Function

Public Function AssetHiddenCell(ByVal assetIndex As Long, ByVal columnNumber As Long) As Range
    Set AssetHiddenCell = ThisWorkbook.Worksheets(SHEET_EXPORT).Cells.Item( _
        ASSET_TECH_FIRST_ROW + ((assetIndex - 1) * ASSET_TECH_ROW_HEIGHT), columnNumber)
End Function

Public Function NextAssetId() As String
    Dim counterCell As Range
    Dim nextNumber As Long

    Set counterCell = ThisWorkbook.Names("TaiSanIDTiepTheo").RefersToRange
    nextNumber = SafeLong(counterCell.Value2, 1)
    If nextNumber < 1 Then nextNumber = 1

    NextAssetId = "TS" & Format$(nextNumber, "000")
    counterCell.Value2 = nextNumber + 1
End Function

Public Function NormalizeKey(ByVal value As Variant) As String
    Dim textValue As String
    Dim sourceChars As Variant
    Dim targetChars As Variant
    Dim charIndex As Long
    Dim currentChar As String
    Dim result As String

    textValue = LCase$(Trim$(ValueToExportText(value)))
    sourceChars = Array( _
        UnicodeText("00E0"), UnicodeText("00E1"), UnicodeText("1EA3"), UnicodeText("00E3"), UnicodeText("1EA1"), _
        UnicodeText("0103"), UnicodeText("1EB1"), UnicodeText("1EAF"), UnicodeText("1EB3"), UnicodeText("1EB5"), UnicodeText("1EB7"), _
        UnicodeText("00E2"), UnicodeText("1EA7"), UnicodeText("1EA5"), UnicodeText("1EA9"), UnicodeText("1EAB"), UnicodeText("1EAD"), _
        UnicodeText("0111"), _
        UnicodeText("00E8"), UnicodeText("00E9"), UnicodeText("1EBB"), UnicodeText("1EBD"), UnicodeText("1EB9"), _
        UnicodeText("00EA"), UnicodeText("1EC1"), UnicodeText("1EBF"), UnicodeText("1EC3"), UnicodeText("1EC5"), UnicodeText("1EC7"), _
        UnicodeText("00EC"), UnicodeText("00ED"), UnicodeText("1EC9"), UnicodeText("0129"), UnicodeText("1ECB"), _
        UnicodeText("00F2"), UnicodeText("00F3"), UnicodeText("1ECF"), UnicodeText("00F5"), UnicodeText("1ECD"), _
        UnicodeText("00F4"), UnicodeText("1ED3"), UnicodeText("1ED1"), UnicodeText("1ED5"), UnicodeText("1ED7"), UnicodeText("1ED9"), _
        UnicodeText("01A1"), UnicodeText("1EDD"), UnicodeText("1EDB"), UnicodeText("1EDF"), UnicodeText("1EE1"), UnicodeText("1EE3"), _
        UnicodeText("00F9"), UnicodeText("00FA"), UnicodeText("1EE7"), UnicodeText("0169"), UnicodeText("1EE5"), _
        UnicodeText("01B0"), UnicodeText("1EEB"), UnicodeText("1EE9"), UnicodeText("1EED"), UnicodeText("1EEF"), UnicodeText("1EF1"), _
        UnicodeText("1EF3"), UnicodeText("00FD"), UnicodeText("1EF7"), UnicodeText("1EF9"), UnicodeText("1EF5"))
    targetChars = Array( _
        "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", _
        "a", "a", "a", "a", "a", "a", _
        "d", _
        "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", _
        "i", "i", "i", "i", "i", _
        "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", _
        "o", "o", "o", "o", "o", "o", _
        "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", _
        "y", "y", "y", "y", "y")
    For charIndex = LBound(sourceChars) To UBound(sourceChars)
        textValue = Replace$(textValue, sourceChars(charIndex), targetChars(charIndex))
    Next charIndex

    For charIndex = 1 To Len(textValue)
        currentChar = Mid$(textValue, charIndex, 1)
        If (currentChar >= "a" And currentChar <= "z") Or _
           (currentChar >= "0" And currentChar <= "9") Then
            result = result & currentChar
        End If
    Next charIndex
    NormalizeKey = result
End Function

Public Function DateDisplay(ByVal value As Variant) As String
    If IsDate(value) Then
        DateDisplay = Format$(CDate(value), "dd/mm/yyyy")
    Else
        DateDisplay = Trim$(ValueToExportText(value))
    End If
End Function

Public Function HasEngineDate(ByVal rowIndex As Long, ByVal calcHeader As String) As Boolean
    Dim value As Variant

    value = PersonTechCell(rowIndex, calcHeader).Value2
    If IsError(value) Or IsEmpty(value) Then Exit Function
    If IsDate(value) Then
        HasEngineDate = True
    ElseIf IsNumeric(value) Then
        On Error Resume Next
        HasEngineDate = (CDbl(value) >= -657434 And CDbl(value) <= 2958465)
        Err.Clear
        On Error GoTo 0
    End If
End Function

Public Function EngineDateDisplay(ByVal rowIndex As Long, ByVal calcHeader As String) As String
    If HasEngineDate(rowIndex, calcHeader) Then
        EngineDateDisplay = Format$(CDate(PersonTechCell(rowIndex, calcHeader).Value2), "dd/mm/yyyy")
    Else
        EngineDateDisplay = vbNullString
    End If
End Function

Public Function FullYearsBetween(ByVal startDate As Date, ByVal endDate As Date) As Long
    Dim yearsValue As Long

    yearsValue = Year(endDate) - Year(startDate)
    If Format$(endDate, "mmdd") < Format$(startDate, "mmdd") Then yearsValue = yearsValue - 1
    FullYearsBetween = yearsValue
End Function

Public Function AgeAtDeath(ByVal rowIndex As Long) As Variant
    If Not HasEngineDate(rowIndex, COL_BIRTH_CALC) Then Exit Function
    If Not HasEngineDate(rowIndex, COL_DEATH_CALC) Then Exit Function
    If CDate(PersonTechCell(rowIndex, COL_DEATH_CALC).Value2) < _
       CDate(PersonTechCell(rowIndex, COL_BIRTH_CALC).Value2) Then Exit Function

    AgeAtDeath = FullYearsBetween(CDate(PersonTechCell(rowIndex, COL_BIRTH_CALC).Value2), _
                                  CDate(PersonTechCell(rowIndex, COL_DEATH_CALC).Value2))
End Function

Public Function YearDisplay(ByVal value As Variant) As String
    If IsDate(value) Then
        YearDisplay = CStr(Year(CDate(value)))
    Else
        YearDisplay = Trim$(ValueToExportText(value))
    End If
End Function

Public Function SafeFileName(ByVal value As String) As String
    Dim badChars As Variant
    Dim item As Variant
    Dim result As String

    result = Trim$(value)
    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each item In badChars
        result = Replace$(result, CStr(item), "-")
    Next item
    If Len(result) = 0 Then result = "Ho_so_thua_ke"
    SafeFileName = result
End Function

Public Sub GoToPersonCell(ByVal rowIndex As Long, ByVal headerName As String)
    ThisWorkbook.Worksheets(SHEET_INPUT).Activate
    PersonCell(rowIndex, headerName).Select
End Sub

Public Sub ApplyWorkbookProtection()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ProtectSheetStandard ws
    Next ws
End Sub

Public Sub ProtectSheetStandard(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=PROTECTION_PASSWORD
    On Error GoTo 0
    ws.Protect Password:=PROTECTION_PASSWORD, DrawingObjects:=True, _
               Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
               AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub
