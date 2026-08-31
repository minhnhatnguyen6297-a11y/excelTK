Attribute VB_Name = "modNgayThang"
Option Explicit

Public Function NormalizeChangedDateCells(ByVal changedRange As Range, _
                                           Optional ByVal showMessage As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim invalidCount As Long
    Dim firstInvalid As String

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then
        NormalizeChangedDateCells = True
        Exit Function
    End If

    NormalizeDateIntersection changedRange, COL_BIRTH, COL_BIRTH_RAW, COL_BIRTH_CALC, _
                              False, invalidCount, firstInvalid
    NormalizeDateIntersection changedRange, COL_DEATH, COL_DEATH_RAW, COL_DEATH_CALC, _
                              False, invalidCount, firstInvalid
    NormalizeDateIntersection changedRange, COL_ISSUED, COL_ISSUED_RAW, COL_ISSUED_CALC, _
                              False, invalidCount, firstInvalid

    NormalizeChangedDateCells = (invalidCount = 0)
    If showMessage And invalidCount > 0 Then
        MsgBox "Có " & CStr(invalidCount) & " ô ngày chưa hợp lệ." & vbCrLf & _
               "Ô đầu tiên: " & firstInvalid & vbCrLf & _
               "Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", _
               vbExclamation, "Ngày tháng chưa hợp lệ"
    End If
End Function

Public Function NormalizeAllDates(Optional ByVal showMessage As Boolean = False) As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim invalidCount As Long
    Dim firstInvalid As String

    Set lo = PeopleTable()
    If lo.DataBodyRange Is Nothing Then
        NormalizeAllDates = True
        Exit Function
    End If

    For rowIndex = 1 To lo.DataBodyRange.Rows.Count
        NormalizeOneDateCell PersonCell(rowIndex, COL_BIRTH), COL_BIRTH_RAW, COL_BIRTH_CALC, _
                             True, invalidCount, firstInvalid
        NormalizeOneDateCell PersonCell(rowIndex, COL_DEATH), COL_DEATH_RAW, COL_DEATH_CALC, _
                             True, invalidCount, firstInvalid
        NormalizeOneDateCell PersonCell(rowIndex, COL_ISSUED), COL_ISSUED_RAW, COL_ISSUED_CALC, _
                             True, invalidCount, firstInvalid
    Next rowIndex

    NormalizeAllDates = (invalidCount = 0)
    If showMessage And invalidCount > 0 Then
        MsgBox "Có " & CStr(invalidCount) & " ô ngày chưa hợp lệ." & vbCrLf & _
               "Ô đầu tiên: " & firstInvalid & vbCrLf & _
               "Chỉ nhập dd/mm/yyyy, mm/yyyy hoặc yyyy.", _
               vbExclamation, "Ngày tháng chưa hợp lệ"
    End If
End Function

Private Sub NormalizeDateIntersection(ByVal changedRange As Range, ByVal visibleHeader As String, _
                                      ByVal rawHeader As String, ByVal calcHeader As String, _
                                      ByVal preserveExistingRaw As Boolean, _
                                      ByRef invalidCount As Long, ByRef firstInvalid As String)
    Dim lo As ListObject
    Dim affected As Range
    Dim oneCell As Range

    Set lo = PeopleTable()
    Set affected = Intersect(changedRange, lo.ListColumns(visibleHeader).DataBodyRange)
    If affected Is Nothing Then Exit Sub

    For Each oneCell In affected.Cells
        NormalizeOneDateCell oneCell, rawHeader, calcHeader, preserveExistingRaw, _
                             invalidCount, firstInvalid
    Next oneCell
End Sub

Private Sub NormalizeOneDateCell(ByVal inputCell As Range, ByVal rawHeader As String, _
                                 ByVal calcHeader As String, ByVal preserveExistingRaw As Boolean, _
                                 ByRef invalidCount As Long, ByRef firstInvalid As String)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim rawCell As Range
    Dim calcCell As Range
    Dim normalizedText As String
    Dim originalText As String
    Dim calculatedDate As Date
    Dim errorReason As String
    Dim keepOldRaw As Boolean

    Set lo = PeopleTable()
    rowIndex = inputCell.Row - lo.DataBodyRange.Row + 1
    Set rawCell = PersonCell(rowIndex, rawHeader)
    Set calcCell = PersonCell(rowIndex, calcHeader)

    keepOldRaw = preserveExistingRaw And Len(Trim$(CStr(rawCell.Value2))) > 0 And _
                 IsDate(calcCell.Value) And _
                 (Trim$(CStr(inputCell.Value2)) = Trim$(CStr(rawCell.Value2)) Or _
                  Trim$(CStr(inputCell.Value2)) = Format$(CDate(calcCell.Value), "dd/mm/yyyy"))

    inputCell.NumberFormat = "@"
    If Len(Trim$(CStr(inputCell.Value2))) = 0 Then
        inputCell.ClearContents
        rawCell.ClearContents
        calcCell.ClearContents
        Exit Sub
    End If

    If TryParseLegalDate(inputCell.Value2, normalizedText, originalText, calculatedDate, errorReason) Then
        rawCell.NumberFormat = "@"
        If keepOldRaw Then
            originalText = CStr(rawCell.Value2)
        Else
            rawCell.Value2 = originalText
        End If
        inputCell.Value2 = originalText
        calcCell.Value = calculatedDate
        calcCell.NumberFormat = "dd/mm/yyyy"
    Else
        rawCell.NumberFormat = "@"
        If Not keepOldRaw Then rawCell.Value2 = originalText
        calcCell.ClearContents
        invalidCount = invalidCount + 1
        If Len(firstInvalid) = 0 Then firstInvalid = inputCell.Address(False, False)
    End If
End Sub

Public Function NormalizeStandaloneDateCell(ByVal inputCell As Range, ByVal rawCell As Range, _
                                            ByVal calcCell As Range) As Boolean
    Dim normalizedText As String
    Dim originalText As String
    Dim calculatedDate As Date
    Dim errorReason As String

    inputCell.NumberFormat = "@"
    If Len(Trim$(CStr(inputCell.Value2))) = 0 Then
        inputCell.ClearContents
        rawCell.ClearContents
        calcCell.ClearContents
        NormalizeStandaloneDateCell = True
        Exit Function
    End If

    rawCell.NumberFormat = "@"
    If TryParseLegalDate(inputCell.Value2, normalizedText, originalText, calculatedDate, errorReason) Then
        inputCell.Value2 = originalText
        rawCell.Value2 = originalText
        calcCell.Value = calculatedDate
        calcCell.NumberFormat = "dd/mm/yyyy"
        NormalizeStandaloneDateCell = True
    Else
        rawCell.Value2 = originalText
        calcCell.ClearContents
    End If
End Function

Public Function TryParseLegalDate(ByVal inputValue As Variant, ByRef normalizedText As String, _
                                  ByRef originalText As String, ByRef calculatedDate As Date, _
                                  ByRef errorReason As String) As Boolean
    Dim cleaned As String
    Dim parts As Variant
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim numericValue As Double

    On Error GoTo InvalidDate

    If IsError(inputValue) Or IsEmpty(inputValue) Then GoTo InvalidDate

    If IsNumeric(inputValue) And VarType(inputValue) <> vbString Then
        numericValue = CDbl(inputValue)
        If numericValue = Fix(numericValue) And numericValue >= 1000 And numericValue <= 9999 Then
            originalText = CStr(CLng(numericValue))
            dayValue = 1
            monthValue = 1
            yearValue = CLng(numericValue)
        ElseIf numericValue >= -657434 And numericValue <= 2958465 Then
            calculatedDate = CDate(numericValue)
            originalText = Format$(calculatedDate, "dd/mm/yyyy")
            normalizedText = originalText
            TryParseLegalDate = True
            Exit Function
        Else
            GoTo InvalidDate
        End If
    Else
        originalText = Trim$(CStr(inputValue))
        cleaned = Replace$(originalText, "-", "/")
        cleaned = Replace$(cleaned, ".", "/")
        cleaned = Replace$(cleaned, " ", vbNullString)
        parts = Split(cleaned, "/")

        Select Case UBound(parts) - LBound(parts) + 1
            Case 1
                If Len(CStr(parts(0))) <> 4 Or Not IsNumeric(parts(0)) Then GoTo InvalidDate
                dayValue = 1
                monthValue = 1
                yearValue = CLng(parts(0))
            Case 2
                If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then GoTo InvalidDate
                dayValue = 1
                monthValue = CLng(parts(0))
                yearValue = CLng(parts(1))
            Case 3
                If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or _
                   Not IsNumeric(parts(2)) Then GoTo InvalidDate
                dayValue = CLng(parts(0))
                monthValue = CLng(parts(1))
                yearValue = CLng(parts(2))
            Case Else
                GoTo InvalidDate
        End Select
    End If

    If yearValue < 1000 Or yearValue > 9999 Then GoTo InvalidDate
    If monthValue < 1 Or monthValue > 12 Then GoTo InvalidDate
    If dayValue < 1 Or dayValue > 31 Then GoTo InvalidDate

    calculatedDate = DateSerial(yearValue, monthValue, dayValue)
    If Year(calculatedDate) <> yearValue Or Month(calculatedDate) <> monthValue Or _
       Day(calculatedDate) <> dayValue Then GoTo InvalidDate

    normalizedText = Format$(calculatedDate, "dd/mm/yyyy")
    TryParseLegalDate = True
    Exit Function

InvalidDate:
    errorReason = "Chỉ dùng dd/mm/yyyy, mm/yyyy hoặc yyyy."
    TryParseLegalDate = False
End Function
