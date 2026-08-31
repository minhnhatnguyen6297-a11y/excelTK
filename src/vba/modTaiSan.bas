Attribute VB_Name = "modTaiSan"
Option Explicit

Public Sub HandleTaiSanChange(ByVal changedRange As Range)
    Dim assetIndex As Long
    Dim affected As Range
    Dim cardRange As Range

    For assetIndex = 1 To ASSET_CARD_COUNT
        Set cardRange = AssetInputRange(assetIndex)
        Set affected = Intersect(changedRange, cardRange)
        If Not affected Is Nothing Then
            If Not Intersect(affected, AssetValueCell(assetIndex, ASSET_FIELD_NGAY_CAP_SO_OFFSET)) Is Nothing Then
                NormalizeStandaloneDateCell AssetValueCell(assetIndex, ASSET_FIELD_NGAY_CAP_SO_OFFSET), _
                                            AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_RAW_COLUMN), _
                                            AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_CALC_COLUMN)
            End If
            RefreshTaiSanCard assetIndex
        End If
    Next assetIndex
End Sub

Public Function TaiSanHasData(ByVal assetIndex As Long) As Boolean
    TaiSanHasData = (Len(Trim$(CStr(AssetValueCell(assetIndex, ASSET_FIELD_LOAI_SO_OFFSET).Value2))) > 0 Or _
                     Len(Trim$(CStr(AssetValueCell(assetIndex, ASSET_FIELD_SERIAL_OFFSET).Value2))) > 0 Or _
                     Len(Trim$(CStr(AssetValueCell(assetIndex, ASSET_FIELD_SO_THUA_OFFSET).Value2))) > 0 Or _
                     Len(Trim$(CStr(AssetValueCell(assetIndex, ASSET_FIELD_DIA_CHI_DAT_OFFSET).Value2))) > 0)
End Function

Public Sub RefreshTaiSanCard(ByVal assetIndex As Long)
    If TaiSanHasData(assetIndex) Then
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_HAS_DATA_COLUMN).Value2 = True
        If Len(Trim$(CStr(AssetHiddenCell(assetIndex, ASSET_HIDDEN_ID_COLUMN).Value2))) = 0 Then
            AssetHiddenCell(assetIndex, ASSET_HIDDEN_ID_COLUMN).Value2 = NextAssetId()
        End If
    Else
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_ID_COLUMN).ClearContents
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_RAW_COLUMN).ClearContents
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_CALC_COLUMN).ClearContents
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_HAS_DATA_COLUMN).Value2 = False
    End If
End Sub

Public Sub RefreshTaiSanTable()
    Dim wsExport As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim assetIndex As Long
    Dim outputRow As ListRow
    Dim columnIndex As Long
    Dim fieldOffset As Long
    Dim wasProtected As Boolean
    Dim reusedBlankRow As Boolean

    Set wsExport = ThisWorkbook.Worksheets(SHEET_EXPORT)
    wasProtected = wsExport.ProtectContents
    If wasProtected Then wsExport.Unprotect PROTECTION_PASSWORD
    On Error Resume Next
    Set lo = wsExport.ListObjects(TABLE_ASSET_EXPORT)
    On Error GoTo 0
    headers = Array("STTTaiSan", "TaiSanID", "LoaiSo", "Serial", "SoVaoSo", "SoThua", _
                    "SoToBanDo", "DiaChiDat", "DienTich", "HinhThucSuDung", "LoaiDat", _
                    "ThoiHanSuDung", "ONT", "CLN", "NTS", "LUC", "NguonGoc", _
                    "NgayCapSoGoc", "NgayCapSoTinh", "CoQuanCapSo", "GhiChu")

    If lo Is Nothing Then
        For columnIndex = LBound(headers) To UBound(headers)
            wsExport.Cells(1, 27 + columnIndex).Value2 = headers(columnIndex)
        Next columnIndex
        Set lo = wsExport.ListObjects.Add(xlSrcRange, wsExport.Range("AA1:AU2"), , xlYes)
        lo.Name = TABLE_ASSET_EXPORT
    Else
        For columnIndex = LBound(headers) To UBound(headers)
            lo.HeaderRowRange.Cells(1, columnIndex + 1).Value2 = headers(columnIndex)
        Next columnIndex
        lo.Resize wsExport.Range("AA1:AU2")
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    End If

    For assetIndex = 1 To ASSET_CARD_COUNT
        If TaiSanHasData(assetIndex) Then
            If lo.ListRows.Count = 1 And Not reusedBlankRow Then
                Set outputRow = lo.ListRows(1)
                reusedBlankRow = True
            Else
                Set outputRow = lo.ListRows.Add
            End If
            outputRow.Range.Cells(1, 1).Value2 = outputRow.Index
            outputRow.Range.Cells(1, 2).Value2 = AssetHiddenCell(assetIndex, ASSET_HIDDEN_ID_COLUMN).Value2
            For fieldOffset = 1 To ASSET_FIELD_COUNT
                columnIndex = fieldOffset + 2
                If fieldOffset = ASSET_FIELD_NGAY_CAP_SO_OFFSET Then
                    outputRow.Range.Cells(1, 18).NumberFormat = "@"
                    outputRow.Range.Cells(1, 18).Value2 = AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_RAW_COLUMN).Value2
                    outputRow.Range.Cells(1, 19).Value2 = AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_CALC_COLUMN).Value2
                ElseIf fieldOffset < ASSET_FIELD_NGAY_CAP_SO_OFFSET Then
                    outputRow.Range.Cells(1, columnIndex).Value2 = AssetValueCell(assetIndex, fieldOffset).Value2
                ElseIf fieldOffset > ASSET_FIELD_NGAY_CAP_SO_OFFSET Then
                    outputRow.Range.Cells(1, columnIndex + 1).Value2 = AssetValueCell(assetIndex, fieldOffset).Value2
                End If
            Next fieldOffset
        End If
    Next assetIndex
    If wasProtected Then ProtectSheetStandard wsExport
End Sub

Private Function AssetInputRange(ByVal assetIndex As Long) As Range
    Dim firstRow As Long

    firstRow = AssetCardStartRow(assetIndex) + 1
    Set AssetInputRange = ThisWorkbook.Worksheets(SHEET_INPUT).Range( _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow, ASSET_VALUE_COLUMN), _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow + ASSET_FIELD_COUNT - 1, 7))
End Function
