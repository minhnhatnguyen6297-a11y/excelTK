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

Private Function AssetInputRange(ByVal assetIndex As Long) As Range
    Dim firstRow As Long

    firstRow = AssetCardStartRow(assetIndex) + 1
    Set AssetInputRange = ThisWorkbook.Worksheets(SHEET_INPUT).Range( _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow, ASSET_VALUE_COLUMN), _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow + ASSET_FIELD_COUNT - 1, 7))
End Function
