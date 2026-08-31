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

    On Error GoTo SafeExit
    Set lo = Me.ListObjects(TABLE_PEOPLE)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    HandlePeopleChange Target

SafeExit:
    Application.EnableEvents = True
End Sub

