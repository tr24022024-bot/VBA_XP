Attribute VB_Name = "modWordCorrection"
Option Explicit

Public Sub Word_CorrectSelection_FromFiles()
    Dim DictionaryPath As String
    Dim PairsPath As String
    Dim DictionaryWords As Object
    Dim ExactPairs As Object
    Dim SourceRange As Range
    Dim OriginalText As String
    Dim CorrectedText As String

    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "Сначала выделите текст для исправления.", vbInformation
        Exit Sub
    End If

    DictionaryPath = InputBox( _
        "Путь к текстовому словарю: одно правильное слово в строке.")

    If Len(DictionaryPath) = 0 Then Exit Sub
    If Dir$(DictionaryPath) = vbNullString Then
        MsgBox "Файл словаря не найден.", vbExclamation
        Exit Sub
    End If

    PairsPath = InputBox( _
        "Путь к CSV-файлу точных замен. Можно оставить пустым.")

    Set DictionaryWords = Word_LoadDictionaryFile(DictionaryPath)
    Set ExactPairs = Word_LoadPairsFile(PairsPath)

    Set SourceRange = Selection.Range.Duplicate
    OriginalText = SourceRange.Text
    CorrectedText = Word_CorrectTextPreserveSeparators( _
        OriginalText, DictionaryWords, ExactPairs, 2)

    SourceRange.Text = CorrectedText
End Sub

Private Function Word_LoadDictionaryFile(ByVal FilePath As String) As Object
    Dim ResultDictionary As Object
    Dim FileNumber As Integer
    Dim LineValue As String

    Set ResultDictionary = CreateObject("Scripting.Dictionary")
    ResultDictionary.CompareMode = vbTextCompare

    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber

    Do Until EOF(FileNumber)
        Line Input #FileNumber, LineValue
        LineValue = Trim$(LineValue)

        If Len(LineValue) > 0 Then
            ResultDictionary(LineValue) = LineValue
        End If
    Loop

    Close #FileNumber
    Set Word_LoadDictionaryFile = ResultDictionary
End Function

Private Function Word_LoadPairsFile(ByVal FilePath As String) As Object
    Dim ResultPairs As Object
    Dim FileNumber As Integer
    Dim LineValue As String
    Dim Parts As Variant

    Set ResultPairs = CreateObject("Scripting.Dictionary")
    ResultPairs.CompareMode = vbTextCompare

    If Len(FilePath) = 0 Or Dir$(FilePath) = vbNullString Then
        Set Word_LoadPairsFile = ResultPairs
        Exit Function
    End If

    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber

    Do Until EOF(FileNumber)
        Line Input #FileNumber, LineValue
        Parts = Split(LineValue, ",")

        If UBound(Parts) >= 1 Then
            If LCase$(Trim$(Parts(0))) <> "wrong" Then
                ResultPairs(Trim$(Parts(0))) = Trim$(Parts(1))
            End If
        End If
    Loop

    Close #FileNumber
    Set Word_LoadPairsFile = ResultPairs
End Function

Private Function Word_CorrectTextPreserveSeparators( _
    ByVal OriginalText As String, _
    ByVal DictionaryWords As Object, _
    ByVal ExactPairs As Object, _
    ByVal MaximumDistance As Long) As String

    Dim RegularExpression As Object
    Dim Matches As Object
    Dim MatchItem As Object
    Dim ResultText As String
    Dim OriginalWord As String
    Dim CorrectedWord As String
    Dim i As Long
    Dim StartPosition As Long

    Set RegularExpression = CreateObject("VBScript.RegExp")

    With RegularExpression
        .Global = True
        .IgnoreCase = True
        .Pattern = "[A-Za-zА-Яа-яЁё]+"
    End With

    ResultText = OriginalText
    Set Matches = RegularExpression.Execute(OriginalText)

    For i = Matches.Count - 1 To 0 Step -1
        Set MatchItem = Matches.Item(i)
        OriginalWord = CStr(MatchItem.Value)
        CorrectedWord = Word_CorrectSingleWord( _
            OriginalWord, DictionaryWords, ExactPairs, MaximumDistance)

        If CorrectedWord <> OriginalWord Then
            StartPosition = CLng(MatchItem.FirstIndex) + 1

            ResultText = Left$(ResultText, StartPosition - 1) & _
                         CorrectedWord & _
                         Mid$(ResultText, StartPosition + Len(OriginalWord))
        End If
    Next i

    Word_CorrectTextPreserveSeparators = ResultText
End Function

Private Function Word_CorrectSingleWord( _
    ByVal OriginalWord As String, _
    ByVal DictionaryWords As Object, _
    ByVal ExactPairs As Object, _
    ByVal MaximumDistance As Long) As String

    Dim Key As Variant
    Dim BestWord As String
    Dim BestDistance As Long
    Dim CurrentDistance As Long

    If ExactPairs.Exists(LCase$(OriginalWord)) Then
        Word_CorrectSingleWord = Word_ApplyCase( _
            OriginalWord, CStr(ExactPairs(LCase$(OriginalWord))))
        Exit Function
    End If

    If DictionaryWords.Exists(OriginalWord) Then
        Word_CorrectSingleWord = OriginalWord
        Exit Function
    End If

    BestDistance = MaximumDistance + 1

    For Each Key In DictionaryWords.Keys
        If Abs(Len(CStr(Key)) - Len(OriginalWord)) <= MaximumDistance Then
            CurrentDistance = Word_LevenshteinDistance( _
                OriginalWord, CStr(Key))

            If CurrentDistance < BestDistance Then
                BestDistance = CurrentDistance
                BestWord = CStr(Key)
            End If
        End If
    Next Key

    If BestDistance <= MaximumDistance And Len(BestWord) > 0 Then
        Word_CorrectSingleWord = Word_ApplyCase(OriginalWord, BestWord)
    Else
        Word_CorrectSingleWord = OriginalWord
    End If
End Function

Private Function Word_LevenshteinDistance( _
    ByVal FirstText As String, _
    ByVal SecondText As String) As Long

    Dim FirstLength As Long
    Dim SecondLength As Long
    Dim PreviousRow() As Long
    Dim CurrentRow() As Long
    Dim i As Long
    Dim j As Long
    Dim Cost As Long

    FirstText = LCase$(FirstText)
    SecondText = LCase$(SecondText)
    FirstLength = Len(FirstText)
    SecondLength = Len(SecondText)

    ReDim PreviousRow(0 To SecondLength)
    ReDim CurrentRow(0 To SecondLength)

    For j = 0 To SecondLength
        PreviousRow(j) = j
    Next j

    For i = 1 To FirstLength
        CurrentRow(0) = i

        For j = 1 To SecondLength
            If Mid$(FirstText, i, 1) = Mid$(SecondText, j, 1) Then
                Cost = 0
            Else
                Cost = 1
            End If

            CurrentRow(j) = Word_Minimum3( _
                PreviousRow(j) + 1, _
                CurrentRow(j - 1) + 1, _
                PreviousRow(j - 1) + Cost)
        Next j

        For j = 0 To SecondLength
            PreviousRow(j) = CurrentRow(j)
        Next j
    Next i

    Word_LevenshteinDistance = PreviousRow(SecondLength)
End Function

Private Function Word_Minimum3( _
    ByVal A As Long, _
    ByVal B As Long, _
    ByVal C As Long) As Long

    Word_Minimum3 = A
    If B < Word_Minimum3 Then Word_Minimum3 = B
    If C < Word_Minimum3 Then Word_Minimum3 = C
End Function

Private Function Word_ApplyCase( _
    ByVal OriginalWord As String, _
    ByVal CorrectedWord As String) As String

    If OriginalWord = UCase$(OriginalWord) Then
        Word_ApplyCase = UCase$(CorrectedWord)
    ElseIf Left$(OriginalWord, 1) = UCase$(Left$(OriginalWord, 1)) Then
        Word_ApplyCase = UCase$(Left$(CorrectedWord, 1)) & _
                         Mid$(CorrectedWord, 2)
    Else
        Word_ApplyCase = LCase$(CorrectedWord)
    End If
End Function
