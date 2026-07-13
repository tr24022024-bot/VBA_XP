Attribute VB_Name = "modOCR_Tesseract"
Option Explicit

Private Const DEFAULT_TESSERACT_PATH As String = _
    "C:\Program Files\Tesseract-OCR\tesseract.exe"

Public Sub OCR_Image_To_ActiveSheet()
    Dim ImagePath As String
    Dim TesseractPath As String
    Dim LanguageCode As String
    Dim OutputBase As String
    Dim OutputTextPath As String
    Dim CommandText As String
    Dim ShellObject As Object
    Dim ExitCode As Long
    Dim OCRText As String

    On Error GoTo ErrorHandler

    ImagePath = Repository_SelectImageFile()
    If Len(ImagePath) = 0 Then Exit Sub

    TesseractPath = DEFAULT_TESSERACT_PATH

    If Dir$(TesseractPath) = vbNullString Then
        TesseractPath = InputBox( _
            "Укажите полный путь к tesseract.exe:", _
            MACRO_REPOSITORY_NAME, _
            DEFAULT_TESSERACT_PATH)

        If Len(TesseractPath) = 0 Then Exit Sub
        If Dir$(TesseractPath) = vbNullString Then
            MsgBox "Файл tesseract.exe не найден.", _
                   vbExclamation, MACRO_REPOSITORY_NAME
            Exit Sub
        End If
    End If

    LanguageCode = InputBox( _
        "Языки Tesseract, например eng+rus+ukr:", _
        MACRO_REPOSITORY_NAME, _
        "eng+rus")

    If Len(LanguageCode) = 0 Then LanguageCode = "eng"

    OutputBase = Environ$("TEMP") & "\vba_repository_ocr_" & _
                 Format$(Now, "yyyymmdd_hhnnss")
    OutputTextPath = OutputBase & ".txt"

    CommandText = Repository_Quote(TesseractPath) & " " & _
                  Repository_Quote(ImagePath) & " " & _
                  Repository_Quote(OutputBase) & _
                  " -l " & LanguageCode & " --psm 6"

    Set ShellObject = CreateObject("WScript.Shell")
    ExitCode = ShellObject.Run(CommandText, 0, True)

    If ExitCode <> 0 Then
        MsgBox "Tesseract завершился с кодом " & ExitCode & ".", _
               vbExclamation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    If Dir$(OutputTextPath) = vbNullString Then
        MsgBox "Файл результата OCR не создан.", _
               vbExclamation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    OCRText = Repository_ReadUtf8Text(OutputTextPath)
    Repository_WriteOCRText ActiveSheet, OCRText

    On Error Resume Next
    Kill OutputTextPath
    On Error GoTo 0

    MsgBox "OCR завершён. Исходные строки записаны в столбец A.", _
           vbInformation, MACRO_REPOSITORY_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка OCR_Image_To_ActiveSheet: " & _
           Err.Number & " — " & Err.Description, _
           vbExclamation, MACRO_REPOSITORY_NAME
End Sub

Public Sub ParseDMS_FromColumnA()
    Dim TargetSheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim DegreesValue As Variant
    Dim MinutesValue As Variant
    Dim SecondsValue As Variant

    Set TargetSheet = ActiveSheet
    LastRow = Repository_LastUsedRow(TargetSheet, 1)

    TargetSheet.Cells(1, 2).Value = "Градусы"
    TargetSheet.Cells(1, 3).Value = "Минуты"
    TargetSheet.Cells(1, 4).Value = "Секунды"

    For i = 2 To LastRow
        If Repository_TryParseDMS( _
            CStr(TargetSheet.Cells(i, 1).Value2), _
            DegreesValue, MinutesValue, SecondsValue) Then

            TargetSheet.Cells(i, 2).Value = DegreesValue
            TargetSheet.Cells(i, 3).Value = MinutesValue
            TargetSheet.Cells(i, 4).Value = SecondsValue
        End If
    Next i
End Sub

Private Function Repository_SelectImageFile() As String
    Dim DialogObject As Object

    Set DialogObject = Application.FileDialog(3)

    With DialogObject
        .Title = "Выберите изображение для OCR"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Изображения", "*.png;*.jpg;*.jpeg;*.tif;*.tiff;*.bmp"

        If .Show = -1 Then
            Repository_SelectImageFile = .SelectedItems(1)
        End If
    End With
End Function

Private Function Repository_Quote(ByVal TextValue As String) As String
    Repository_Quote = Chr$(34) & TextValue & Chr$(34)
End Function

Private Function Repository_ReadUtf8Text(ByVal FilePath As String) As String
    Dim StreamObject As Object

    Set StreamObject = CreateObject("ADODB.Stream")

    With StreamObject
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile FilePath
        Repository_ReadUtf8Text = .ReadText
        .Close
    End With
End Function

Private Sub Repository_WriteOCRText( _
    ByVal TargetSheet As Worksheet, _
    ByVal OCRText As String)

    Dim Lines As Variant
    Dim i As Long
    Dim OutputRow As Long
    Dim DegreesValue As Variant
    Dim MinutesValue As Variant
    Dim SecondsValue As Variant

    Lines = Split(Replace(OCRText, vbCrLf, vbLf), vbLf)

    TargetSheet.Range("A:D").ClearContents
    TargetSheet.Cells(1, 1).Value = "Исходная строка OCR"
    TargetSheet.Cells(1, 2).Value = "Градусы"
    TargetSheet.Cells(1, 3).Value = "Минуты"
    TargetSheet.Cells(1, 4).Value = "Секунды"

    OutputRow = 2

    For i = LBound(Lines) To UBound(Lines)
        If Len(Trim$(CStr(Lines(i)))) > 0 Then
            TargetSheet.Cells(OutputRow, 1).Value = CStr(Lines(i))

            If Repository_TryParseDMS( _
                CStr(Lines(i)), _
                DegreesValue, MinutesValue, SecondsValue) Then

                TargetSheet.Cells(OutputRow, 2).Value = DegreesValue
                TargetSheet.Cells(OutputRow, 3).Value = MinutesValue
                TargetSheet.Cells(OutputRow, 4).Value = SecondsValue
            End If

            OutputRow = OutputRow + 1
        End If
    Next i

    TargetSheet.Columns("A:D").AutoFit
End Sub

Private Function Repository_TryParseDMS( _
    ByVal TextValue As String, _
    ByRef DegreesValue As Variant, _
    ByRef MinutesValue As Variant, _
    ByRef SecondsValue As Variant) As Boolean

    Dim RegularExpression As Object
    Dim Matches As Object

    Set RegularExpression = CreateObject("VBScript.RegExp")

    With RegularExpression
        .Global = False
        .IgnoreCase = True
        .Pattern = "(-?\d{1,3})\s*[" & _
                   ChrW$(176) & ChrW$(186) & "]\s*" & _
                   "(\d{1,2})\s*['" & ChrW$(8242) & "]\s*" & _
                   "(\d{1,2}([.,]\d+)?)\s*[" & _
                   ChrW$(34) & ChrW$(8243) & "]?"
    End With

    Set Matches = RegularExpression.Execute(TextValue)

    If Matches.Count = 0 Then Exit Function

    DegreesValue = CLng(Matches(0).SubMatches(0))
    MinutesValue = CLng(Matches(0).SubMatches(1))
    SecondsValue = Repository_ParseDecimal( _
        CStr(Matches(0).SubMatches(2)))

    Repository_TryParseDMS = True
End Function

Private Function Repository_ParseDecimal(ByVal TextValue As String) As Double
    Dim DecimalSeparator As String

    DecimalSeparator = Application.DecimalSeparator
    TextValue = Replace(TextValue, ".", DecimalSeparator)
    TextValue = Replace(TextValue, ",", DecimalSeparator)

    Repository_ParseDecimal = CDbl(TextValue)
End Function
