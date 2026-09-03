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
    Dim assetIndex As Long
    Dim extensionAffected As Boolean
    Dim assetAffected As Boolean
    Dim peopleAffected As Boolean
    Dim profileAffected As Boolean

    On Error GoTo SafeExit
    If Not Intersect(Target, PersonExtensionRange()) Is Nothing Then
        extensionAffected = True
    End If
    If Not Intersect(Target, Me.Range("B41:C1048576")) Is Nothing Then
        profileAffected = True
    End If

    For assetIndex = 1 To ASSET_CARD_COUNT
        If Not Intersect(Target, AssetInputRange(assetIndex)) Is Nothing Then
            assetAffected = True
            Exit For
        End If
    Next assetIndex

    Set lo = Me.ListObjects(TABLE_PEOPLE)
    If Not lo.DataBodyRange Is Nothing Then
        peopleAffected = Not Intersect(Target, lo.DataBodyRange) Is Nothing
    End If
    If Not extensionAffected And Not assetAffected And Not peopleAffected Then Exit Sub

    Application.EnableEvents = False
    If extensionAffected Then HandlePersonExtensionChange Target
    If assetAffected Then HandleTaiSanChange Target
    If peopleAffected Then HandlePeopleChange Target
    If profileAffected Then NormalizeProfileTextInputs Target

SafeExit:
    Application.EnableEvents = True
End Sub
