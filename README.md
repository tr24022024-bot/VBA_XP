# Примеры

## Форматирование

```vb
FormatTimeStatus_Default
```

или:

```vb
FormatTimeStatus_Range Sheet1.Range("B2:G100"), " - "
```

## Выбор без повторов

```vb
SelectWithoutRepeats_Range _
    SourceRange:=Sheet1.Range("D20:BR20"), _
    FirstOutputCell:=Sheet1.Range("D22"), _
    OutputCount:=10, _
    VerticalOutput:=False
```

## Кодирование

1. Выполните `CreateCodebookSheet`.
2. Заполните девять кодов в столбце B листа `КОДИРОВКА`.
3. Выполните `EncodeValues_Default`.

## Словарь

1. Выполните `CreateDictionarySheet`.
2. В столбце A разместите правильные слова.
3. В столбцах B:C разместите точные пары «ошибка → замена».
4. Выполните `CorrectText_Default`.

## OCR

```vb
OCR_Image_To_ActiveSheet
```
