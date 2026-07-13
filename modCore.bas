Attribute VB_Name = "modCore"
Option Explicit

Public Const MACRO_REPOSITORY_NAME As String = "VBA Macro Repository"
Public Const MACRO_REPOSITORY_VERSION As String = "0.1.0"

Public Sub MacroRepository_Info()
    MsgBox MACRO_REPOSITORY_NAME & vbCrLf & _
           "Версия: " & MACRO_REPOSITORY_VERSION, _
           vbInformation, MACRO_REPOSITORY_NAME
End Sub

Public Sub Repository_Initialize()
    On Error GoTo ErrorHandler

    CreateCodebookSheet
    CreateDictionarySheet

    MsgBox "Начальная структура создана:" & vbCrLf & _
           "1. Лист КОДИРОВКА" & vbCrLf & _
           "2. Лист СЛОВАРЬ", _
           vbInformation, MACRO_REPOSITORY_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка Repository_Initialize: " & _
           Err.Number & " — " & Err.Description, _
           vbExclamation, MACRO_REPOSITORY_NAME
End Sub

Public Function Repository_NormalizeText(ByVal Value As Variant) As String
    Dim TextValue As String

    If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
        Repository_NormalizeText = vbNullString
        Exit Function
    End If

    TextValue = CStr(Value)
    TextValue = Replace(TextValue, ChrW(160), " ")
    TextValue = Replace(TextValue, vbTab, " ")
    TextValue = Trim$(TextValue)

    Do While InStr(1, TextValue, "  ", vbBinaryCompare) > 0
        TextValue = Replace(TextValue, "  ", " ")
    Loop

    Repository_NormalizeText = TextValue
End Function

Public Function Repository_WorksheetExists( _
    ByVal WorksheetName As String, _
    Optional ByVal WorkbookObject As Workbook = Nothing) As Boolean

    Dim TargetWorkbook As Workbook
    Dim TargetSheet As Worksheet

    If WorkbookObject Is Nothing Then
        Set TargetWorkbook = ThisWorkbook
    Else
        Set TargetWorkbook = WorkbookObject
    End If

    On Error Resume Next
    Set TargetSheet = TargetWorkbook.Worksheets(WorksheetName)
    On Error GoTo 0

    Repository_WorksheetExists = Not TargetSheet Is Nothing
End Function

Public Function Repository_LastUsedRow( _
    ByVal TargetSheet As Worksheet, _
    Optional ByVal ColumnNumber As Long = 1) As Long

    Dim LastRow As Long

    LastRow = TargetSheet.Cells(TargetSheet.Rows.Count, ColumnNumber).End(xlUp).Row

    If LastRow = 1 And Len(CStr(TargetSheet.Cells(1, ColumnNumber).Value2)) = 0 Then
        Repository_LastUsedRow = 0
    Else
        Repository_LastUsedRow = LastRow
    End If
End Function
