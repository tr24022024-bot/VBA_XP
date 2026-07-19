Attribute VB_Name = "modTextCorrection"
Option Explicit

Private Const DICTIONARY_SHEET_NAME As String = "СЛОВАРЬ"

Public Sub CreateDictionarySheet()
    Dim TargetSheet As Worksheet

    If Repository_WorksheetExists(DICTIONARY_SHEET_NAME, ThisWorkbook) Then
        Set TargetSheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
    Else
        Set TargetSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        TargetSheet.Name = DICTIONARY_SHEET_NAME
    End If

    TargetSheet.Cells.Clear

    TargetSheet.Cells(1, 1).Value = "Правильные слова"
    TargetSheet.Cells(1, 2).Value = "Ошибочная форма"
    TargetSheet.Cells(1, 3).Value = "Точная замена"

    TargetSheet.Cells(2, 1).Value = "girl"
    TargetSheet.Cells(3, 1).Value = "truck"
    TargetSheet.Cells(4, 1).Value = "bicycle"

    TargetSheet.Cells(2, 2).Value = "girll"
    TargetSheet.Cells(2, 3).Value = "girl"
    TargetSheet.Cells(3, 2).Value = "truk"
    TargetSheet.Cells(3, 3).Value = "truck"
    TargetSheet.Cells(4, 2).Value = "biccicle"
    TargetSheet.Cells(4, 3).Value = "bicycle"

    TargetSheet.Columns("A:C").AutoFit
End Sub

Public Sub CorrectText_Default()
    CorrectText_Range _
        TargetRange:=ActiveSheet.Range("B385:G1000"), _
        MaximumDistance:=2
End Sub

Public Sub CorrectText_Range( _
    ByVal TargetRange As Range, _
    Optional ByVal MaximumDistance As Long = 2)

    Dim DictionarySheet As Worksheet
    Dim DictionaryWords As Object
    Dim ExactPairs As Object
    Dim Cell As Range
    Dim OriginalText As String
    Dim CorrectedText As String
    Dim OldScreenUpdating As Boolean
    Dim OldEnableEvents As Boolean

    On Error GoTo ErrorHandler

    If Not Repository_WorksheetExists(DICTIONARY_SHEET_NAME, ThisWorkbook) Then
        MsgBox "Лист '" & DICTIONARY_SHEET_NAME & "' не найден." & vbCrLf & _
               "Сначала выполните CreateDictionarySheet.", _
               vbExclamation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    Set DictionarySheet = ThisWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
    Set DictionaryWords = Repository_LoadDictionary(DictionarySheet)
    Set ExactPairs = Repository_LoadExactPairs(DictionarySheet)

    If DictionaryWords.Count = 0 And ExactPairs.Count = 0 Then
        MsgBox "Словарь пуст.", vbInformation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    OldScreenUpdating = Application.ScreenUpdating
    OldEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each Cell In TargetRange.Cells
        If Not Cell.HasFormula And Not IsError(Cell.Value2) Then
            OriginalText = CStr(Cell.Value2)

            If Len(OriginalText) > 0 Then
                CorrectedText = Repository_CorrectTextPreserveSeparators( _
                    OriginalText, DictionaryWords, ExactPairs, MaximumDistance)

                If CorrectedText <> OriginalText Then
                    Cell.Value = CorrectedText
                End If
            End If
        End If
    Next Cell

CleanExit:
    Application.ScreenUpdating = OldScreenUpdating
    Application.EnableEvents = OldEnableEvents
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка CorrectText_Range: " & _
           Err.Number & " — " & Err.Description, _
           vbExclamation, MACRO_REPOSITORY_NAME
    Resume CleanExit
End Sub

Private Function Repository_LoadDictionary( _
    ByVal DictionarySheet As Worksheet) As Object

    Dim ResultDictionary As Object
    Dim LastRow As Long
    Dim i As Long
    Dim WordValue As String

    Set ResultDictionary = CreateObject("Scripting.Dictionary")
    ResultDictionary.CompareMode = vbTextCompare

    LastRow = Repository_LastUsedRow(DictionarySheet, 1)

    For i = 2 To LastRow
        WordValue = Trim$(CStr(DictionarySheet.Cells(i, 1).Value2))

        If Len(WordValue) > 0 Then
            If Not ResultDictionary.Exists(WordValue) Then
                ResultDictionary.Add WordValue, WordValue
            End If
        End If
    Next i

    Set Repository_LoadDictionary = ResultDictionary
End Function

Private Function Repository_LoadExactPairs( _
    ByVal DictionarySheet As Worksheet) As Object

    Dim ResultPairs As Object
    Dim LastRow As Long
    Dim i As Long
    Dim WrongValue As String
    Dim ReplacementValue As String

    Set ResultPairs = CreateObject("Scripting.Dictionary")
    ResultPairs.CompareMode = vbTextCompare

    LastRow = Repository_LastUsedRow(DictionarySheet, 2)

    For i = 2 To LastRow
        WrongValue = Trim$(CStr(DictionarySheet.Cells(i, 2).Value2))
        ReplacementValue = Trim$(CStr(DictionarySheet.Cells(i, 3).Value2))

        If Len(WrongValue) > 0 And Len(ReplacementValue) > 0 Then
            ResultPairs(WrongValue) = ReplacementValue
        End If
    Next i

    Set Repository_LoadExactPairs = ResultPairs
End Function

Private Function Repository_CorrectTextPreserveSeparators( _
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
    Dim StartPosition As Long

    Set RegularExpression = CreateObject("VBScript.RegExp")

    With RegularExpression
        .Global = True
        .IgnoreCase = True
        .Pattern = "[A-Za-zА-Яа-яЁё]+"
    End With

    ResultText = OriginalText
    Set Matches = RegularExpression.Execute(OriginalText)

    For Each MatchItem In Repository_ReverseMatches(Matches)
        OriginalWord = CStr(MatchItem.Value)
        CorrectedWord = Repository_CorrectSingleWord( _
            OriginalWord, DictionaryWords, ExactPairs, MaximumDistance)

        If CorrectedWord <> OriginalWord Then
            StartPosition = CLng(MatchItem.FirstIndex) + 1

            ResultText = Left$(ResultText, StartPosition - 1) & _
                         CorrectedWord & _
                         Mid$(ResultText, StartPosition + Len(OriginalWord))
        End If
    Next MatchItem

    Repository_CorrectTextPreserveSeparators = ResultText
End Function

Private Function Repository_ReverseMatches(ByVal Matches As Object) As Collection
    Dim ResultCollection As New Collection
    Dim i As Long

    For i = Matches.Count - 1 To 0 Step -1
        ResultCollection.Add Matches.Item(i)
    Next i

    Set Repository_ReverseMatches = ResultCollection
End Function

Private Function Repository_CorrectSingleWord( _
    ByVal OriginalWord As String, _
    ByVal DictionaryWords As Object, _
    ByVal ExactPairs As Object, _
    ByVal MaximumDistance As Long) As String

    Dim Key As Variant
    Dim BestWord As String
    Dim BestDistance As Long
    Dim CurrentDistance As Long
    Dim NormalizedWord As String

    NormalizedWord = LCase$(OriginalWord)

    If ExactPairs.Exists(NormalizedWord) Then
        Repository_CorrectSingleWord = Repository_ApplyCase( _
            OriginalWord, CStr(ExactPairs(NormalizedWord)))
        Exit Function
    End If

    If DictionaryWords.Exists(NormalizedWord) Then
        Repository_CorrectSingleWord = OriginalWord
        Exit Function
    End If

    BestDistance = MaximumDistance + 1

    For Each Key In DictionaryWords.Keys
        If Abs(Len(CStr(Key)) - Len(OriginalWord)) <= MaximumDistance Then
            CurrentDistance = LevenshteinDistance(OriginalWord, CStr(Key))

            If CurrentDistance < BestDistance Then
                BestDistance = CurrentDistance
                BestWord = CStr(DictionaryWords(Key))

                If BestDistance = 1 Then Exit For
            End If
        End If
    Next Key

    If BestDistance <= MaximumDistance And Len(BestWord) > 0 Then
        Repository_CorrectSingleWord = Repository_ApplyCase( _
            OriginalWord, BestWord)
    Else
        Repository_CorrectSingleWord = OriginalWord
    End If
End Function

Private Function Repository_ApplyCase( _
    ByVal OriginalWord As String, _
    ByVal CorrectedWord As String) As String

    If OriginalWord = UCase$(OriginalWord) Then
        Repository_ApplyCase = UCase$(CorrectedWord)
    ElseIf Left$(OriginalWord, 1) = UCase$(Left$(OriginalWord, 1)) Then
        Repository_ApplyCase = UCase$(Left$(CorrectedWord, 1)) & _
                               Mid$(CorrectedWord, 2)
    Else
        Repository_ApplyCase = LCase$(CorrectedWord)
    End If
End Function
