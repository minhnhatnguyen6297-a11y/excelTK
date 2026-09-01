Attribute VB_Name = "modXuatWord"
Option Explicit

Public Function ReplaceTemplateTokens(ByVal templatePath As String) As String
    Dim app As Object, doc As Object, outputPath As String, folder As String, failureText As String
    On Error GoTo Failed
    folder = GetConfigText("ThuMucXuat")
    outputPath = folder & Application.PathSeparator & SafeFileName("Ho_so_thua_ke_" & Format$(Now, "yyyymmdd_hhnnss")) & ".docx"
    If Dir$(outputPath) <> vbNullString Then outputPath = folder & Application.PathSeparator & SafeFileName("Ho_so_thua_ke_" & Format$(Now, "yyyymmdd_hhnnss") & "_1") & ".docx"
    FileCopy templatePath, outputPath
    Set app = CreateObject("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    Set doc = app.Documents.Open(outputPath, False, False)
    ReplaceDocumentRanges doc
    doc.Save
    doc.Close False
    app.Quit False
    Set doc = Nothing: Set app = Nothing
    ReplaceTemplateTokens = outputPath
    Exit Function
Failed:
    failureText = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not app Is Nothing Then app.Quit False
    If Len(outputPath) > 0 And Dir$(outputPath) <> vbNullString Then Kill outputPath
    Err.Raise vbObjectError + 1701, "modXuatWord.ReplaceTemplateTokens", failureText
End Function

Private Sub ReplaceDocumentRanges(ByVal doc As Object)
    Dim story As Object, current As Object, shapeItem As Object
    For Each story In doc.StoryRanges
        Set current = story
        Do While Not current Is Nothing
            ReplaceRangeTokens current
            Set current = current.NextStoryRange
        Loop
    Next story
    For Each shapeItem In doc.Shapes
        On Error Resume Next
        If shapeItem.TextFrame.HasText Then ReplaceRangeTokens shapeItem.TextFrame.TextRange
        On Error GoTo 0
    Next shapeItem
End Sub

Private Sub ReplaceRangeTokens(ByVal sourceRange As Object)
    ReplacePattern sourceRange, "\[*\]"
    ReplacePattern sourceRange, "\{\{*\}\}"
End Sub

Private Sub ReplacePattern(ByVal sourceRange As Object, ByVal pattern As String)
    Dim search As Object, token As String, replacement As String, originalEnd As Long, lengthChange As Long
    Set search = sourceRange.Duplicate
    originalEnd = search.End
    Do
        With search.Find
            .ClearFormatting: .Text = pattern: .Forward = True: .Wrap = 0: .Format = False: .MatchWildcards = True
        End With
        If Not search.Find.Execute Then Exit Do
        token = CStr(search.Text)
        replacement = ResolveToken(token)
        If replacement = token Then
            search.Start = search.End
            search.End = originalEnd
        Else
            lengthChange = Len(replacement) - Len(token)
            search.Text = vbNullString
            InsertReplacementText search, replacement
            originalEnd = originalEnd + lengthChange
            search.Start = search.End
            search.End = originalEnd
        End If
    Loop
End Sub

Private Function ResolveToken(ByVal token As String) As String
    Dim inner As String, parts() As String, slot As Long, field As String, i As Long
    If Left$(token, 1) = "[" Then
        inner = Mid$(token, 2, Len(token) - 2)
        If inner = "Nam chet" Or inner = UnicodeText("004E 0103 006D 0020 0063 0068 1EBF 0074") Then ResolveToken = LegacySlotValue(1, "Nam chet"): Exit Function
        If inner = "Nam chet 2" Or inner = UnicodeText("004E 0103 006D 0020 0063 0068 1EBF 0074 0020 0032") Then ResolveToken = LegacySlotValue(2, "Nam chet 2"): Exit Function
        parts = Split(inner, " "): slot = 1
        If IsNumeric(parts(UBound(parts))) Then slot = CLng(parts(UBound(parts))): ReDim Preserve parts(UBound(parts) - 1)
        For i = LBound(parts) To UBound(parts)
            If Len(field) > 0 Then field = field & " "
            field = field & parts(i)
        Next i
        If field = UnicodeText("0054 00EA 006E") Then
            field = "Ten"
        ElseIf field = UnicodeText("004E 0103 006D 0020 0073 0069 006E 0068") Then
            field = "Nam sinh"
        ElseIf field = UnicodeText("004E 0067 00E0 0079 0020 0063 1EA5 0070") Then
            field = "Ngay cap"
        ElseIf field = UnicodeText("0110 1ECB 0061 0020 0063 0068 1EC9") Then
            field = "Dia chi"
        ElseIf field = UnicodeText("004C 006F 1EA1 0069 0020 0043 0043") Then
            field = "Loai CC"
        ElseIf field = UnicodeText("004E 01A1 0069 0020 0063 1EA5 0070 0020 0043 0043") Then
            field = "Noi cap CC"
        ElseIf field = UnicodeText("0054 0068 01B0 1EDD 006E 0067 0020 0074 0072 00FA") Then
            field = "Thuong tru"
        End If
        ResolveToken = LegacySlotValue(slot, field)
    ElseIf Left$(token, 2) = "{{" Then
        inner = Mid$(token, 3, Len(token) - 4)
        If Left$(inner, 8) = "tai_san." Or Left$(inner, 6) = "ho_so." Then ResolveToken = NamespaceValue(inner) Else ResolveToken = token
    End If
End Function

Private Sub InsertReplacementText(ByVal target As Object, ByVal value As String)
    Dim startAt As Long, take As String
    If Len(value) <= 254 Then
        If Len(value) > 0 Then target.InsertAfter Replace$(value, vbCrLf, vbVerticalTab)
        Exit Sub
    End If
    startAt = 1
    Do While startAt <= Len(value)
        take = Mid$(value, startAt, 254)
        target.InsertAfter Replace$(take, vbCrLf, vbVerticalTab)
        target.Start = target.End
        startAt = startAt + 254
    Loop
End Sub

Private Function NamespaceValue(ByVal inner As String) As String
    Dim p() As String, lo As ListObject, i As Long
    p = Split(inner, ".")
    If p(0) = "ho_so" Then
        If UBound(p) = 1 Then NamespaceValue = ExportSectionValue("HoSo", HoSoFieldName(p(1)))
    ElseIf p(0) = "tai_san" And UBound(p) >= 2 Then
        Set lo = ThisWorkbook.Worksheets(SHEET_EXPORT).ListObjects(TABLE_ASSET_EXPORT)
        For i = 1 To lo.ListRows.Count
            If CLng(lo.DataBodyRange.Cells(i, 1).Value2) = CLng(p(1)) Then NamespaceValue = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(p(2)).Index).Value2): Exit Function
        Next i
    End If
End Function

Private Function HoSoFieldName(ByVal name As String) As String
    Select Case name
        Case "niem_yet": HoSoFieldName = "NiemYet"
        Case "so_cong_chung": HoSoFieldName = "SoCongChung"
        Case "nguoi_uy_quyen": HoSoFieldName = "NguoiUyQuyen"
        Case "nguoi_uy_quyen2": HoSoFieldName = "NguoiUyQuyen2"
    End Select
End Function

Private Function ExportSectionValue(ByVal sectionName As String, ByVal fieldName As String) As String
    Dim wsExport As Worksheet, sectionRow As Long, lastRow As Long, lastColumn As Long, columnIndex As Long, value As Variant
    If Len(fieldName) = 0 Then Exit Function
    Set wsExport = ThisWorkbook.Worksheets(SHEET_EXPORT)
    lastRow = wsExport.Cells(wsExport.Rows.Count, 1).End(xlUp).Row
    For sectionRow = 1 To lastRow - 2
        If CStr(wsExport.Cells(sectionRow, 1).Value2) = sectionName Then
            lastColumn = wsExport.Cells(sectionRow + 1, wsExport.Columns.Count).End(xlToLeft).Column
            For columnIndex = 1 To lastColumn
                If CStr(wsExport.Cells(sectionRow + 1, columnIndex).Value2) = fieldName Then
                    value = wsExport.Cells(sectionRow + 2, columnIndex).Value2
                    If Not IsError(value) And Not IsEmpty(value) Then ExportSectionValue = CStr(value)
                    Exit Function
                End If
            Next columnIndex
            Exit Function
        End If
    Next sectionRow
End Function
