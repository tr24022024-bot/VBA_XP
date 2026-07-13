Attribute VB_Name = "modValueCodec"
Option Explicit

Private Const CODEBOOK_SHEET_NAME As String = "КОДИРОВКА"
Private Const CODE_GROUP_SIZE As Long = 9
Private Const MAX_ENCODED_VALUE As Long = 54

Public Sub CreateCodebookSheet()
    Dim TargetSheet As Worksheet
    Dim DefaultCodes As Variant
    Dim i As Long

    If Repository_WorksheetExists(CODEBOOK_SHEET_NAME, ThisWorkbook) Then
        Set TargetSheet = ThisWorkbook.Worksheets(CODEBOOK_SHEET_NAME)
    Else
        Set TargetSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        TargetSheet.Name = CODEBOOK_SHEET_NAME
    End If

    DefaultCodes = Array( _
        "Su", "Code2", "Code3", "Code4", "Code5", _
        "Code6", "Code7", "Code8", "Code9")

    TargetSheet.Cells.Clear

    TargetSheet.Cells(1, 1).Value = "Позиция"
    TargetSheet.Cells(1, 2).Value = "Буквенный код"
    TargetSheet.Cells(1, 3).Value = "Комментарий"

    For i = 1 To CODE_GROUP_SIZE
        TargetSheet.Cells(i + 1, 1).Value = i
        TargetSheet.Cells(i + 1, 2).Value = DefaultCodes(i - 1)
    Next i

    TargetSheet.Range("C2").Value = _
        "Замените Code2…Code9 собственными кодами."
    TargetSheet.Columns("A:C").AutoFit
End Sub

Public Sub EncodeValues_Default()
    EncodeValues_Range _
        SourceRange:=ActiveSheet.Range("D3:I3"), _
        LetterOutputFirstCell:=ActiveSheet.Range("L3"), _
        CoefficientOutputFirstCell:=ActiveSheet.Range("L4")
End Sub

Public Sub DecodeValues_Default()
    DecodeValues_Range _
        LetterRange:=ActiveSheet.Range("T3:Y3"), _
        CoefficientRange:=ActiveSheet.Range("T4:Y4"), _
        OutputFirstCell:=ActiveSheet.Range("AB3")
End Sub

Public Sub EncodeValues_Range( _
    ByVal SourceRange As Range, _
    ByVal LetterOutputFirstCell As Range, _
    ByVal CoefficientOutputFirstCell As Range)

    Dim TargetSheet As Worksheet
    Dim Cell As Range
    Dim PositionIndex As Long
    Dim NumericValue As Long
    Dim BasePosition As Long
    Dim Coefficient As Long
    Dim LetterCode As String
    Dim OutputIndex As Long

    Set TargetSheet = Repository_GetCodebookSheet()
    If TargetSheet Is Nothing Then Exit Sub

    OutputIndex = 0

    For Each Cell In SourceRange.Cells
        OutputIndex = OutputIndex + 1

        If IsNumeric(Cell.Value2) Then
            NumericValue = CLng(Cell.Value2)

            If NumericValue >= 1 And NumericValue <= MAX_ENCODED_VALUE Then
                BasePosition = ((NumericValue - 1) Mod CODE_GROUP_SIZE) + 1
                Coefficient = ((NumericValue - 1) \ CODE_GROUP_SIZE) + 1
                LetterCode = CStr(TargetSheet.Cells(BasePosition + 1, 2).Value2)

                LetterOutputFirstCell.Offset(0, OutputIndex - 1).Value = LetterCode
                CoefficientOutputFirstCell.Offset(0, OutputIndex - 1).Value = Coefficient
            Else
                LetterOutputFirstCell.Offset(0, OutputIndex - 1).Value = "#RANGE"
                CoefficientOutputFirstCell.Offset(0, OutputIndex - 1).ClearContents
            End If
        Else
            LetterOutputFirstCell.Offset(0, OutputIndex - 1).ClearContents
            CoefficientOutputFirstCell.Offset(0, OutputIndex - 1).ClearContents
        End If
    Next Cell
End Sub

Public Sub DecodeValues_Range( _
    ByVal LetterRange As Range, _
    ByVal CoefficientRange As Range, _
    ByVal OutputFirstCell As Range)

    Dim TargetSheet As Worksheet
    Dim i As Long
    Dim ItemCount As Long
    Dim LetterCode As String
    Dim Coefficient As Long
    Dim BasePosition As Long
    Dim DecodedValue As Long

    If LetterRange.Cells.Count <> CoefficientRange.Cells.Count Then
        MsgBox "Диапазоны букв и коэффициентов должны иметь одинаковый размер.", _
               vbExclamation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    Set TargetSheet = Repository_GetCodebookSheet()
    If TargetSheet Is Nothing Then Exit Sub

    ItemCount = LetterRange.Cells.Count

    For i = 1 To ItemCount
        LetterCode = Trim$(CStr(LetterRange.Cells(i).Value2))

        If Len(LetterCode) = 0 Or Not IsNumeric(CoefficientRange.Cells(i).Value2) Then
            OutputFirstCell.Offset(0, i - 1).ClearContents
        Else
            Coefficient = CLng(CoefficientRange.Cells(i).Value2)
            BasePosition = Repository_FindCodePosition(TargetSheet, LetterCode)

            If BasePosition = 0 Or Coefficient < 1 Or Coefficient > 6 Then
                OutputFirstCell.Offset(0, i - 1).Value = "#CODE"
            Else
                DecodedValue = BasePosition + _
                               ((Coefficient - 1) * CODE_GROUP_SIZE)
                OutputFirstCell.Offset(0, i - 1).Value = DecodedValue
            End If
        End If
    Next i
End Sub

Public Function EncodeNumberToCode( _
    ByVal NumericValue As Long, _
    ByRef LetterCode As String, _
    ByRef Coefficient As Long) As Boolean

    Dim TargetSheet As Worksheet
    Dim BasePosition As Long

    EncodeNumberToCode = False

    If NumericValue < 1 Or NumericValue > MAX_ENCODED_VALUE Then Exit Function

    Set TargetSheet = Repository_GetCodebookSheet()
    If TargetSheet Is Nothing Then Exit Function

    BasePosition = ((NumericValue - 1) Mod CODE_GROUP_SIZE) + 1
    Coefficient = ((NumericValue - 1) \ CODE_GROUP_SIZE) + 1
    LetterCode = CStr(TargetSheet.Cells(BasePosition + 1, 2).Value2)

    EncodeNumberToCode = (Len(LetterCode) > 0)
End Function

Public Function DecodeCodeToNumber( _
    ByVal LetterCode As String, _
    ByVal Coefficient As Long, _
    ByRef NumericValue As Long) As Boolean

    Dim TargetSheet As Worksheet
    Dim BasePosition As Long

    DecodeCodeToNumber = False

    Set TargetSheet = Repository_GetCodebookSheet()
    If TargetSheet Is Nothing Then Exit Function

    BasePosition = Repository_FindCodePosition(TargetSheet, LetterCode)

    If BasePosition = 0 Then Exit Function
    If Coefficient < 1 Or Coefficient > 6 Then Exit Function

    NumericValue = BasePosition + ((Coefficient - 1) * CODE_GROUP_SIZE)
    DecodeCodeToNumber = True
End Function

Private Function Repository_GetCodebookSheet() As Worksheet
    If Not Repository_WorksheetExists(CODEBOOK_SHEET_NAME, ThisWorkbook) Then
        MsgBox "Лист '" & CODEBOOK_SHEET_NAME & "' не найден." & vbCrLf & _
               "Сначала выполните CreateCodebookSheet.", _
               vbExclamation, MACRO_REPOSITORY_NAME
        Exit Function
    End If

    Set Repository_GetCodebookSheet = _
        ThisWorkbook.Worksheets(CODEBOOK_SHEET_NAME)
End Function

Private Function Repository_FindCodePosition( _
    ByVal CodebookSheet As Worksheet, _
    ByVal LetterCode As String) As Long

    Dim i As Long

    For i = 1 To CODE_GROUP_SIZE
        If StrComp( _
            Trim$(CStr(CodebookSheet.Cells(i + 1, 2).Value2)), _
            Trim$(LetterCode), _
            vbTextCompare) = 0 Then

            Repository_FindCodePosition = i
            Exit Function
        End If
    Next i
End Function
