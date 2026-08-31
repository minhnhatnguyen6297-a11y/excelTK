Attribute VB_Name = "modGiayTo"
Option Explicit

Public Sub UpdateGiayToForRow(ByVal rowIndex As Long)
    Dim issuedDate As Variant
    Dim loaiCC As String
    Dim noiCap As String
    Dim nhanDiaChi As String

    issuedDate = PersonCell(rowIndex, COL_ISSUED_CALC).Value
    If Not IsDate(issuedDate) Then
        PersonCell(rowIndex, COL_LOAI_CC).ClearContents
        PersonCell(rowIndex, COL_NOI_CAP_CC).ClearContents
        PersonCell(rowIndex, COL_NHAN_DIA_CHI).ClearContents
        Exit Sub
    End If

    If CDate(issuedDate) < DateSerial(2024, 7, 1) Then
        loaiCC = CStr(ThisWorkbook.Names("DanhMuc_LoaiCCTruocMoc").RefersToRange.Value2)
    Else
        loaiCC = CStr(ThisWorkbook.Names("DanhMuc_LoaiCCTuMoc").RefersToRange.Value2)
    End If
    PersonCell(rowIndex, COL_LOAI_CC).Value2 = loaiCC
    noiCap = LookupCatalogValue("DanhMuc_LoaiGiayTo", "DanhMuc_NoiCapCC", "DanhMuc_NhanDiaChi", 0, loaiCC, 2)
    nhanDiaChi = LookupCatalogValue("DanhMuc_LoaiGiayTo", "DanhMuc_NoiCapCC", "DanhMuc_NhanDiaChi", 0, loaiCC, 3)
    PersonCell(rowIndex, COL_NOI_CAP_CC).Value2 = noiCap
    PersonCell(rowIndex, COL_NHAN_DIA_CHI).Value2 = nhanDiaChi
End Sub

Private Function LookupCatalogValue(ByVal keyName As String, ByVal noicapName As String, _
                                    ByVal nhanName As String, ByVal fixedIndex As Long, _
                                    Optional ByVal keyValue As String = vbNullString, _
                                    Optional ByVal resultKind As Long = 1) As String
    Dim keys As Range
    Dim results As Range
    Dim index As Long
    Dim wanted As String

    Set keys = ThisWorkbook.Names(keyName).RefersToRange
    If resultKind = 2 Then
        Set results = ThisWorkbook.Names(noicapName).RefersToRange
    ElseIf resultKind = 3 Then
        Set results = ThisWorkbook.Names(nhanName).RefersToRange
    Else
        Set results = keys
    End If
    If fixedIndex > 0 Then
        LookupCatalogValue = CStr(results.Cells(fixedIndex, 1).Value2)
        Exit Function
    End If
    wanted = Trim$(keyValue)
    For index = 1 To keys.Rows.Count
        If Trim$(CStr(keys.Cells(index, 1).Value2)) = wanted Then
            LookupCatalogValue = CStr(results.Cells(index, 1).Value2)
            Exit Function
        End If
    Next index
End Function
