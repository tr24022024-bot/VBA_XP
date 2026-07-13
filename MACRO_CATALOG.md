# Каталог процедур

| Модуль | Публичная процедура | Назначение |
|---|---|---|
| modCore | `MacroRepository_Info` | показывает версию библиотеки |
| modCore | `Repository_Initialize` | создаёт служебные листы |
| modFormatTimeStatus | `FormatTimeStatus_Default` | форматирует B892:G11000 |
| modFormatTimeStatus | `FormatTimeStatus_Range` | форматирует переданный диапазон |
| modRandomSelection | `SelectWithoutRepeats_Default` | выборка D20:BR20 в строку с D22 |
| modRandomSelection | `SelectWithoutRepeats_Range` | универсальная выборка |
| modValueCodec | `CreateCodebookSheet` | создаёт таблицу кодов |
| modValueCodec | `EncodeValues_Default` | D3:I3 → L3:Q4 |
| modValueCodec | `DecodeValues_Default` | T3:Y4 → начиная с AB3 |
| modTextCorrection | `CreateDictionarySheet` | создаёт лист словаря |
| modTextCorrection | `CorrectText_Default` | исправляет B385:G1000 |
| modOCR_Tesseract | `OCR_Image_To_ActiveSheet` | OCR выбранного изображения |
| modOCR_Tesseract | `ParseDMS_FromColumnA` | разбирает градусы, минуты, секунды |
| modWordCorrection | `Word_CorrectSelection_FromFiles` | исправляет выделенный текст Word |
| modSelfTests | `Repository_SelfTest` | запускает самопроверку |
