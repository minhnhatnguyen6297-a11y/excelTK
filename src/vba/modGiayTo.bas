Attribute VB_Name = "modGiayTo"
Option Explicit

Public Sub UpdateGiayToForRow(ByVal rowIndex As Long)
    Dim issuedDate As Variant
    Dim loaiCC As String
    Dim noiCap As String
    Dim nhanDiaChi As String

    issuedDate = PersonTechCell(rowIndex, COL_ISSUED_CALC).Value2
    If Not IsDate(issuedDate) Then
        PersonTechCell(rowIndex, COL_LOAI_CC).ClearContents
        PersonTechCell(rowIndex, COL_NOI_CAP_CC).ClearContents
        PersonTechCell(rowIndex, COL_NHAN_DIA_CHI).ClearContents
        Exit Sub
    End If

    If CDate(issuedDate) < DateSerial(2024, 7, 1) Then
        loaiCC = CStr(ThisWorkbook.Names("DanhMuc_LoaiCCTruocMoc").RefersToRange.Value2)
    Else
        loaiCC = CStr(ThisWorkbook.Names("DanhMuc_LoaiCCTuMoc").RefersToRange.Value2)
    End If
    PersonTechCell(rowIndex, COL_LOAI_CC).Value2 = loaiCC
    noiCap = LookupCatalogValue("DanhMuc_LoaiGiayTo", "DanhMuc_NoiCapCC", loaiCC)
    nhanDiaChi = LookupCatalogValue("DanhMuc_LoaiGiayTo", "DanhMuc_NhanDiaChi", loaiCC)
    PersonTechCell(rowIndex, COL_NOI_CAP_CC).Value2 = noiCap
    PersonTechCell(rowIndex, COL_NHAN_DIA_CHI).Value2 = nhanDiaChi
End Sub

Private Function LookupCatalogValue(ByVal keyName As String, ByVal resultName As String, ByVal keyValue As String) As String
    Dim keys As Range
    Dim results As Range
    Dim index As Long
    Dim wanted As String

    Set keys = ThisWorkbook.Names(keyName).RefersToRange
    Set results = ThisWorkbook.Names(resultName).RefersToRange
    wanted = Trim$(keyValue)
    For index = 1 To keys.Rows.Count
        If Trim$(CStr(keys.Cells(index, 1).Value2)) = wanted Then
            LookupCatalogValue = CStr(results.Cells(index, 1).Value2)
            Exit Function
        End If
    Next index
End Function
