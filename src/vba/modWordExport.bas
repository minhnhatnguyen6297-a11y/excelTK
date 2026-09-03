Attribute VB_Name = "modWordExport"
Option Explicit

Public Sub XuatVanBan()
    Call RunWordExport(False)
End Sub

Public Sub ChonNoiLayMau()
    Dim picker As Object
    Dim selectedFolder As String
    Dim previousTemplate As String
    Dim candidateTemplate As String
    Dim templateName As String

    On Error GoTo PickerFailed
    Set picker = Application.FileDialog(4)
    picker.AllowMultiSelect = False
    picker.Title = "Chọn nơi lấy mẫu Word"
    If Len(GetConfigText("ThuMucMauWord")) > 0 Then picker.InitialFileName = GetConfigText("ThuMucMauWord")
    If picker.Show <> -1 Then Exit Sub
    selectedFolder = CStr(picker.SelectedItems(1))

    previousTemplate = GetConfigText("MauWordDangChon")
    If Len(previousTemplate) > 0 Then templateName = Dir$(previousTemplate)
    If Len(templateName) > 0 Then candidateTemplate = selectedFolder & Application.PathSeparator & templateName
    SetConfigText "ThuMucMauWord", selectedFolder
    If Len(candidateTemplate) > 0 And LCase$(Right$(candidateTemplate, 5)) = ".docx" And _
       Len(Dir$(candidateTemplate)) > 0 Then
        SetConfigText "MauWordDangChon", candidateTemplate
        UpdateTemplateCapacity
    Else
        SetConfigText "MauWordDangChon", vbNullString
        SetConfigNumber "SucChuaNguoiMau", 0
        SetConfigNumber "SucChuaTaiSanMau", 0
    End If
    RefreshWordSelectionDisplay
    Exit Sub

PickerFailed:
    MsgBox Err.Description, vbExclamation, "Không thể chọn nơi lấy mẫu"
End Sub

Public Sub ChonVanBan()
    Dim picker As Object
    Dim selectedFile As String
    Dim templateFolder As String
    Dim fso As Object

    On Error GoTo PickerFailed
    templateFolder = GetConfigText("ThuMucMauWord")
    If Len(templateFolder) = 0 Or Dir$(templateFolder, vbDirectory) = vbNullString Then
        MsgBox "Hãy chọn nơi lấy mẫu trước.", vbExclamation, "Chưa có thư mục mẫu"
        Exit Sub
    End If

    Set picker = Application.FileDialog(3)
    picker.AllowMultiSelect = False
    picker.Title = "Chọn văn bản Word"
    picker.InitialFileName = templateFolder & Application.PathSeparator
    picker.Filters.Clear
    picker.Filters.Add "Word Documents", "*.docx"
    If picker.Show <> -1 Then Exit Sub
    selectedFile = CStr(picker.SelectedItems(1))
    If LCase$(Right$(selectedFile, 5)) <> ".docx" Then
        MsgBox "Chỉ nhận file .docx.", vbExclamation, "Mẫu không hợp lệ"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If LCase$(fso.GetParentFolderName(selectedFile)) <> LCase$(templateFolder) Then
        MsgBox "Chỉ chọn file .docx nằm trực tiếp trong thư mục mẫu.", vbExclamation, "Mẫu không hợp lệ"
        Exit Sub
    End If

    SetConfigText "MauWordDangChon", selectedFile
    UpdateTemplateCapacity
    RefreshWordSelectionDisplay
    Exit Sub

PickerFailed:
    MsgBox Err.Description, vbExclamation, "Không thể chọn văn bản"
End Sub

Public Sub ChonNoiXuat()
    Dim picker As Object
    Dim selectedFolder As String

    On Error GoTo PickerFailed
    Set picker = Application.FileDialog(4)
    picker.AllowMultiSelect = False
    picker.Title = "Chọn nơi xuất văn bản"
    If Len(GetConfigText("ThuMucXuat")) > 0 Then picker.InitialFileName = GetConfigText("ThuMucXuat")
    If picker.Show <> -1 Then Exit Sub
    selectedFolder = CStr(picker.SelectedItems(1))
    SetConfigText "ThuMucXuat", selectedFolder
    RefreshWordSelectionDisplay
    Exit Sub

PickerFailed:
    MsgBox Err.Description, vbExclamation, "Không thể chọn nơi xuất"
End Sub

Public Sub RefreshWordSelectionDisplay()
    Dim ws As Worksheet
    Dim templatePath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT)
    templatePath = GetConfigText("MauWordDangChon")
    ws.Range("B3").Value2 = ShortPathText(GetConfigText("ThuMucMauWord"))
    If Len(templatePath) > 0 Then
        ws.Range("C5").Value2 = Dir$(templatePath)
    Else
        ws.Range("C5").Value2 = vbNullString
    End If
    ws.Range("B7").Value2 = ShortPathText(GetConfigText("ThuMucXuat"))
End Sub

Private Sub UpdateTemplateCapacity()
    Dim catalog As Worksheet
    Dim templatePath As String
    Dim templateName As String
    Dim rowNumber As Long
    Dim found As Boolean

    templatePath = GetConfigText("MauWordDangChon")
    If Len(templatePath) = 0 Then
        SetConfigNumber "SucChuaNguoiMau", 0
        SetConfigNumber "SucChuaTaiSanMau", 0
        Exit Sub
    End If
    templateName = Dir$(templatePath)
    Set catalog = ThisWorkbook.Worksheets("DanhMuc")
    For rowNumber = 4 To catalog.Cells(catalog.Rows.Count, 10).End(xlUp).Row
        If CStr(catalog.Cells(rowNumber, 10).Value2) = templateName Then
            SetConfigNumber "SucChuaNguoiMau", SafeLong(catalog.Cells(rowNumber, 11).Value2, 0)
            SetConfigNumber "SucChuaTaiSanMau", SafeLong(catalog.Cells(rowNumber, 12).Value2, 0)
            found = True
            Exit For
        End If
    Next rowNumber
    If Not found Then
        SetConfigNumber "SucChuaNguoiMau", 0
        SetConfigNumber "SucChuaTaiSanMau", 0
    End If
End Sub

Public Function RunWordExport(Optional ByVal silent As Boolean = False) As Boolean
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim errorText As String

    On Error GoTo ExportFailed

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    RefreshWordSelectionDisplay
    RefreshAllPeople

    Call ValidateWorkbook(Not silent)

    BuildExportData
    templatePath = GetConfigText("DuongDanMauWord")
    outputFolder = GetConfigText("ThuMucXuat")

    If Len(templatePath) = 0 Or Dir$(templatePath) = vbNullString Then
        errorText = "Không tìm thấy template Word: " & templatePath
        GoTo ExportFailed
    End If
    If Len(outputFolder) = 0 Or Dir$(outputFolder, vbDirectory) = vbNullString Then
        errorText = "Không tìm thấy thư mục xuất: " & outputFolder
        GoTo ExportFailed
    End If

    outputPath = outputFolder & Application.PathSeparator & _
                 SafeFileName("Ho_so_thua_ke_" & Format$(Now, "yyyymmdd_hhnnss")) & ".docx"

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    Set wordDoc = wordApp.Documents.Open(templatePath, False, True)

    ReplaceEverywhere wordDoc, "{{chu_dat.danh_sach}}", GroupNameList("ChuDat")
    ReplaceEverywhere wordDoc, "{{nguoi_da_chet.danh_sach}}", GroupNameList("NguoiDaChet")
    ReplaceEverywhere wordDoc, "{{nguoi_nhan_dat.danh_sach}}", GroupNameList("NguoiNhanDat")
    ReplaceEverywhere wordDoc, "{{nhom_tu_choi.TC_DEFAULT.danh_sach}}", GroupNameList("CacNhomTuChoi")
    ReplaceEverywhere wordDoc, "{{cay_nhanh.danh_sach}}", BranchSummary()
    ReplaceEverywhere wordDoc, "{{kiem_thu.ngay_sinh_nguoi_1}}", FirstPersonDate(COL_BIRTH)
    ReplaceEverywhere wordDoc, "{{kiem_thu.ngay_chet_nguoi_1}}", FirstPersonDate(COL_DEATH)
    ReplaceEverywhere wordDoc, "{{kiem_thu.tuoi_luc_chet_nguoi_1}}", FirstPersonAgeAtDeath()

    If DocumentContainsToken(wordDoc, "{{") Or DocumentContainsToken(wordDoc, "}}") Then
        errorText = "Template vẫn còn placeholder chưa được thay."
        GoTo ExportFailed
    End If

    wordDoc.SaveAs2 outputPath, 16
    wordDoc.Close False
    Set wordDoc = Nothing
    wordApp.Quit False
    Set wordApp = Nothing

    RunWordExport = True
    If Not silent Then
        MsgBox "Đã tạo văn bản mới:" & vbCrLf & outputPath, _
               vbInformation, "Xuất thành công"
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ExportFailed:
    If Len(errorText) = 0 Then errorText = Err.Description
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit False
    Set wordDoc = Nothing
    Set wordApp = Nothing
    If Len(outputPath) > 0 And Dir$(outputPath) <> vbNullString Then Kill outputPath
    On Error GoTo 0

    RunWordExport = False
    If Not silent Then MsgBox errorText, vbExclamation, "Không thể xuất văn bản"
    GoTo CleanExit
End Function

Private Sub ReplaceEverywhere(ByVal wordDoc As Object, ByVal findText As String, _
                              ByVal replacementText As String)
    Dim storyRange As Object
    Dim currentRange As Object
    Dim shapeItem As Object
    Dim sectionItem As Object
    Dim headerFooter As Object

    For Each storyRange In wordDoc.StoryRanges
        Set currentRange = storyRange
        Do While Not currentRange Is Nothing
            ReplaceInRange currentRange, findText, replacementText
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next storyRange

    For Each shapeItem In wordDoc.Shapes
        ReplaceInShape shapeItem, findText, replacementText
    Next shapeItem

    For Each sectionItem In wordDoc.Sections
        For Each headerFooter In sectionItem.Headers
            For Each shapeItem In headerFooter.Shapes
                ReplaceInShape shapeItem, findText, replacementText
            Next shapeItem
        Next headerFooter
        For Each headerFooter In sectionItem.Footers
            For Each shapeItem In headerFooter.Shapes
                ReplaceInShape shapeItem, findText, replacementText
            Next shapeItem
        Next headerFooter
    Next sectionItem
End Sub

Private Sub ReplaceInShape(ByVal shapeItem As Object, ByVal findText As String, _
                           ByVal replacementText As String)
    On Error Resume Next
    If shapeItem.TextFrame.HasText Then
        ReplaceInRange shapeItem.TextFrame.TextRange, findText, replacementText
    End If
    On Error GoTo 0
End Sub

Private Sub ReplaceInRange(ByVal sourceRange As Object, ByVal findText As String, _
                           ByVal replacementText As String)
    Dim searchRange As Object
    Dim originalEnd As Long
    Dim lengthChange As Long

    Set searchRange = sourceRange.Duplicate
    originalEnd = searchRange.End

    Do
        With searchRange.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = 0
            .Format = False
        End With
        If Not searchRange.Find.Execute Then Exit Do

        lengthChange = Len(replacementText) - Len(findText)
        searchRange.Text = replacementText
        originalEnd = originalEnd + lengthChange
        searchRange.Start = searchRange.End
        searchRange.End = originalEnd
    Loop
End Sub

Private Function DocumentContainsToken(ByVal wordDoc As Object, ByVal token As String) As Boolean
    Dim storyRange As Object
    Dim currentRange As Object
    Dim shapeItem As Object
    Dim sectionItem As Object
    Dim headerFooter As Object

    For Each storyRange In wordDoc.StoryRanges
        Set currentRange = storyRange.Duplicate
        Do While Not currentRange Is Nothing
            With currentRange.Find
                .ClearFormatting
                .Text = token
                .Forward = True
                .Wrap = 0
                .Format = False
            End With
            If currentRange.Find.Execute Then
                DocumentContainsToken = True
                Exit Function
            End If
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next storyRange

    For Each shapeItem In wordDoc.Shapes
        If ShapeContainsToken(shapeItem, token) Then
            DocumentContainsToken = True
            Exit Function
        End If
    Next shapeItem

    For Each sectionItem In wordDoc.Sections
        For Each headerFooter In sectionItem.Headers
            For Each shapeItem In headerFooter.Shapes
                If ShapeContainsToken(shapeItem, token) Then
                    DocumentContainsToken = True
                    Exit Function
                End If
            Next shapeItem
        Next headerFooter
        For Each headerFooter In sectionItem.Footers
            For Each shapeItem In headerFooter.Shapes
                If ShapeContainsToken(shapeItem, token) Then
                    DocumentContainsToken = True
                    Exit Function
                End If
            Next shapeItem
        Next headerFooter
    Next sectionItem
End Function

Private Function ShapeContainsToken(ByVal shapeItem As Object, ByVal token As String) As Boolean
    On Error Resume Next
    If shapeItem.TextFrame.HasText Then
        ShapeContainsToken = (InStr(1, shapeItem.TextFrame.TextRange.Text, token, vbTextCompare) > 0)
    End If
    On Error GoTo 0
End Function
