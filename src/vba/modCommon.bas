Attribute VB_Name = "modCommon"
Option Explicit

Public Const SHEET_INPUT As String = "NhapLieu"
Public Const SHEET_CONFIG As String = "CauHinh"
Public Const SHEET_CHECK As String = "KiemTra"
Public Const SHEET_EXPORT As String = "XuatAn"
Public Const TABLE_PEOPLE As String = "tblNguoi"
Public Const TABLE_ASSET_EXPORT As String = "tblTaiSan"

Public Const ASSET_SECTION_ROW As Long = 40
Public Const ASSET_FIRST_CARD_ROW As Long = 41
Public Const ASSET_CARD_HEIGHT As Long = 20
Public Const ASSET_CARD_COUNT As Long = 3
Public Const ASSET_FIELD_COUNT As Long = 18
Public Const ASSET_LABEL_COLUMN As Long = 2
Public Const ASSET_VALUE_COLUMN As Long = 3
Public Const ASSET_HIDDEN_ID_COLUMN As Long = 10
Public Const ASSET_HIDDEN_ISSUED_RAW_COLUMN As Long = 11
Public Const ASSET_HIDDEN_ISSUED_CALC_COLUMN As Long = 12
Public Const ASSET_HIDDEN_HAS_DATA_COLUMN As Long = 13
Public Const ASSET_FIELD_LOAI_SO_OFFSET As Long = 1
Public Const ASSET_FIELD_SERIAL_OFFSET As Long = 2
Public Const ASSET_FIELD_SO_THUA_OFFSET As Long = 4
Public Const ASSET_FIELD_DIA_CHI_DAT_OFFSET As Long = 6
Public Const ASSET_FIELD_NGAY_CAP_SO_OFFSET As Long = 16

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

Public Const PROTECTION_PASSWORD As String = "HoSoTK_MVP_2026"

Public Function PeopleTable() As ListObject
    Set PeopleTable = ThisWorkbook.Worksheets(SHEET_INPUT).ListObjects(TABLE_PEOPLE)
End Function

Public Function ColumnIndex(ByVal headerName As String) As Long
    ColumnIndex = PeopleTable.ListColumns(headerName).Index
End Function

Public Function PersonCell(ByVal rowIndex As Long, ByVal headerName As String) As Range
    Set PersonCell = PeopleTable.DataBodyRange.Cells(rowIndex, ColumnIndex(headerName))
End Function

Public Function PersonName(ByVal rowIndex As Long) As String
    PersonName = Trim$(CStr(PersonCell(rowIndex, COL_NAME).Value2))
End Function

Public Function PersonHasData(ByVal rowIndex As Long) As Boolean
    PersonHasData = (Len(PersonName(rowIndex)) > 0)
End Function

Public Function SafeLong(ByVal value As Variant, Optional ByVal fallback As Long = 0) As Long
    If IsError(value) Or IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
        SafeLong = fallback
    ElseIf IsNumeric(value) Then
        SafeLong = CLng(value)
    Else
        SafeLong = fallback
    End If
End Function

Public Function SafeBool(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        SafeBool = CBool(value)
    ElseIf IsNumeric(value) Then
        SafeBool = (CDbl(value) <> 0)
    Else
        SafeBool = (LCase$(Trim$(CStr(value))) = "true" Or Trim$(CStr(value)) = "có")
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
    GetConfigText = Trim$(CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value2))
    Exit Function
MissingName:
    GetConfigText = vbNullString
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
    AssetCardStartRow = ASSET_FIRST_CARD_ROW + ((assetIndex - 1) * ASSET_CARD_HEIGHT)
End Function

Public Function AssetValueCell(ByVal assetIndex As Long, ByVal fieldOffset As Long) As Range
    Set AssetValueCell = ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item( _
        AssetCardStartRow(assetIndex) + fieldOffset, ASSET_VALUE_COLUMN)
End Function

Public Function AssetHiddenCell(ByVal assetIndex As Long, ByVal columnNumber As Long) As Range
    Set AssetHiddenCell = ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item( _
        AssetCardStartRow(assetIndex), columnNumber)
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
    textValue = LCase$(Trim$(CStr(value)))
    Do While InStr(textValue, "  ") > 0
        textValue = Replace$(textValue, "  ", " ")
    Loop
    NormalizeKey = textValue
End Function

Public Function DateDisplay(ByVal value As Variant) As String
    If IsDate(value) Then
        DateDisplay = Format$(CDate(value), "dd/mm/yyyy")
    Else
        DateDisplay = Trim$(CStr(value))
    End If
End Function

Public Function HasEngineDate(ByVal rowIndex As Long, ByVal calcHeader As String) As Boolean
    HasEngineDate = IsDate(PersonCell(rowIndex, calcHeader).Value)
End Function

Public Function EngineDateDisplay(ByVal rowIndex As Long, ByVal calcHeader As String) As String
    If HasEngineDate(rowIndex, calcHeader) Then
        EngineDateDisplay = Format$(CDate(PersonCell(rowIndex, calcHeader).Value), "dd/mm/yyyy")
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
    If CDate(PersonCell(rowIndex, COL_DEATH_CALC).Value) < _
       CDate(PersonCell(rowIndex, COL_BIRTH_CALC).Value) Then Exit Function

    AgeAtDeath = FullYearsBetween(CDate(PersonCell(rowIndex, COL_BIRTH_CALC).Value), _
                                  CDate(PersonCell(rowIndex, COL_DEATH_CALC).Value))
End Function

Public Function YearDisplay(ByVal value As Variant) As String
    If IsDate(value) Then
        YearDisplay = CStr(Year(CDate(value)))
    Else
        YearDisplay = Trim$(CStr(value))
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
        On Error Resume Next
        ws.Unprotect Password:=PROTECTION_PASSWORD
        On Error GoTo 0

        ws.Protect Password:=PROTECTION_PASSWORD, DrawingObjects:=True, _
                   Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, _
                   AllowFiltering:=True
        ws.EnableSelection = xlNoRestrictions
    Next ws
End Sub
