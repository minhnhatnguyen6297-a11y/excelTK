Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim lo As ListObject

    On Error GoTo SafeExit
    If Target.CountLarge <> 1 Then Exit Sub
    Set lo = Me.ListObjects(TABLE_PEOPLE)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Application.EnableEvents = False
    If Not Intersect(Target, lo.ListColumns(COL_LEVEL).DataBodyRange) Is Nothing Then
        HandleLevelClick Target
    ElseIf Not Intersect(Target, lo.ListColumns(COL_RECEIVE).DataBodyRange) Is Nothing Then
        HandleReceiveClick Target
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject
    Dim assetRange As Range

    On Error GoTo SafeExit
    Set assetRange = Me.Range(Me.Cells(ASSET_FIRST_CARD_ROW + 1, ASSET_VALUE_COLUMN), _
                                Me.Cells(ASSET_FIRST_CARD_ROW + (ASSET_CARD_HEIGHT * ASSET_CARD_COUNT) - 2, 7))
    If Not Intersect(Target, assetRange) Is Nothing Then
        Application.EnableEvents = False
        HandleTaiSanChange Target
        GoTo SafeExit
    End If

    Set lo = Me.ListObjects(TABLE_PEOPLE)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    HandlePeopleChange Target

SafeExit:
    Application.EnableEvents = True
End Sub
