Attribute VB_Name = "modFormatTimeStatus"
Option Explicit

Public Sub FormatTimeStatus_Default()
    FormatTimeStatus_Range _
        TargetRange:=ActiveSheet.Range("B892:G11000"), _
        SeparatorText:=" - "
End Sub

Public Sub FormatTimeStatus_Range( _
    ByVal TargetRange As Range, _
    Optional ByVal SeparatorText As String = " - ")

    Dim Cell As Range
    Dim TextValue As String
    Dim FirstSeparator As Long
    Dim LastSeparator As Long
    Dim LeftLength As Long
    Dim RightStart As Long
    Dim RightLength As Long
    Dim OldScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    OldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For Each Cell In TargetRange.Cells
        If Not IsError(Cell.Value2) Then
            TextValue = CStr(Cell.Value2)

            If Len(TextValue) > 0 Then
                With Cell.Characters(1, Len(TextValue)).Font
                    .Color = vbBlack
                    .Bold = False
                End With

                FirstSeparator = InStr(1, TextValue, SeparatorText, vbBinaryCompare)
                LastSeparator = InStrRev(TextValue, SeparatorText, -1, vbBinaryCompare)

                If FirstSeparator > 1 Then
                    LeftLength = FirstSeparator - 1

                    With Cell.Characters(1, LeftLength).Font
                        .Color = RGB(0, 112, 192)
                        .Bold = True
                    End With
                End If

                If LastSeparator > 0 Then
                    RightStart = LastSeparator + Len(SeparatorText)
                    RightLength = Len(TextValue) - RightStart + 1

                    If RightLength > 0 Then
                        With Cell.Characters(RightStart, RightLength).Font
                            .Color = RGB(192, 0, 0)
                            .Bold = True
                        End With
                    End If
                End If
            End If
        End If
    Next Cell

CleanExit:
    Application.ScreenUpdating = OldScreenUpdating
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка FormatTimeStatus_Range: " & _
           Err.Number & " — " & Err.Description, _
           vbExclamation, MACRO_REPOSITORY_NAME
    Resume CleanExit
End Sub
