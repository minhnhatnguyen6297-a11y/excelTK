Attribute VB_Name = "modTaiSan"
Option Explicit

Public Sub HandleTaiSanChange(ByVal changedRange As Range)
    Dim assetIndex As Long
    Dim affected As Range
    Dim cardRange As Range
    Dim fieldOffset As Long

    For assetIndex = 1 To ASSET_CARD_COUNT
        Set cardRange = AssetInputRange(assetIndex)
        Set affected = Intersect(changedRange, cardRange)
        If Not affected Is Nothing Then
            For fieldOffset = 1 To ASSET_FIELD_COUNT
                If Not Intersect(affected, AssetValueCell(assetIndex, fieldOffset)) Is Nothing Then
                    NormalizeAssetInputCell AssetValueCell(assetIndex, fieldOffset)
                End If
            Next fieldOffset
            If Not Intersect(affected, AssetValueCell(assetIndex, ASSET_FIELD_NGAY_CAP_SO_OFFSET)) Is Nothing Then
                NormalizeStandaloneDateCell AssetValueCell(assetIndex, ASSET_FIELD_NGAY_CAP_SO_OFFSET), _
                                            AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_RAW_COLUMN), _
                                            AssetHiddenCell(assetIndex, ASSET_HIDDEN_ISSUED_CALC_COLUMN)
            End If
            RefreshTaiSanCard assetIndex
        End If
    Next assetIndex
End Sub

Public Sub NormalizeAllTaiSanInputs()
    Dim assetIndex As Long
    Dim fieldOffset As Long

    For assetIndex = 1 To ASSET_CARD_COUNT
        For fieldOffset = 1 To ASSET_FIELD_COUNT
            NormalizeAssetInputCell AssetValueCell(assetIndex, fieldOffset)
        Next fieldOffset
        RefreshTaiSanCard assetIndex
    Next assetIndex
End Sub

Private Sub NormalizeAssetInputCell(ByVal inputCell As Range)
    NormalizeInputTextCell inputCell
End Sub

Public Function TaiSanHasData(ByVal assetIndex As Long) As Boolean
    Dim fieldOffset As Long

    ' A card contains data when any visible value is present.  Do not use a
    ' short list of identifying fields here: a user may start with area,
    ' land type, a land-use amount, or notes and the card must still receive
    ' an ID and be included in diagnostics/export.
    For fieldOffset = 1 To ASSET_FIELD_COUNT
        If Len(Trim$(ValueToExportText(AssetValueCell(assetIndex, fieldOffset).Value2))) > 0 Then
            TaiSanHasData = True
            Exit Function
        End If
    Next fieldOffset
End Function

Public Sub RefreshTaiSanCard(ByVal assetIndex As Long)
    If TaiSanHasData(assetIndex) Then
        AssetHiddenCell(assetIndex, ASSET_HIDDEN_HAS_DATA_COLUMN).Value2 = True
        If Len(Trim$(ValueToExportText(AssetHiddenCell(assetIndex, ASSET_HIDDEN_ID_COLUMN).Value2))) = 0 Then
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
    ' MIN-10 is paused by the user. Keep this entry point for compatibility,
    ' but do not create, rename, resize, or refresh an asset ListObject.
End Sub

Public Function AssetInputRange(ByVal assetIndex As Long) As Range
    Dim firstRow As Long
    Dim firstColumn As Long

    firstRow = AssetCardStartRow(assetIndex) + 1
    firstColumn = AssetCardStartColumn(assetIndex)
    Set AssetInputRange = ThisWorkbook.Worksheets(SHEET_INPUT).Range( _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow, firstColumn), _
        ThisWorkbook.Worksheets(SHEET_INPUT).Cells.Item(firstRow + ASSET_FIELD_COUNT - 1, firstColumn + 1))
End Function
