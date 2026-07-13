Attribute VB_Name = "modRandomSelection"
Option Explicit

Public Sub SelectWithoutRepeats_Default()
    SelectWithoutRepeats_Range _
        SourceRange:=ActiveSheet.Range("D20:BR20"), _
        FirstOutputCell:=ActiveSheet.Range("D22"), _
        OutputCount:=0, _
        VerticalOutput:=False
End Sub

Public Sub SelectWithoutRepeats_Range( _
    ByVal SourceRange As Range, _
    ByVal FirstOutputCell As Range, _
    Optional ByVal OutputCount As Long = 0, _
    Optional ByVal VerticalOutput As Boolean = False)

    Dim Values() As Variant
    Dim UniqueValues As Object
    Dim UniqueKey As String
    Dim Cell As Range
    Dim ValueCount As Long
    Dim i As Long
    Dim j As Long
    Dim TemporaryValue As Variant
    Dim ActualOutputCount As Long
    Dim OutputArray() As Variant

    Set UniqueValues = CreateObject("Scripting.Dictionary")
    UniqueValues.CompareMode = vbTextCompare

    For Each Cell In SourceRange.Cells
        If Not IsError(Cell.Value2) Then
            If Len(Trim$(CStr(Cell.Value2))) > 0 Then
                UniqueKey = Trim$(CStr(Cell.Value2))

                If Not UniqueValues.Exists(UniqueKey) Then
                    UniqueValues.Add UniqueKey, Cell.Value2
                    ValueCount = ValueCount + 1
                    ReDim Preserve Values(1 To ValueCount)
                    Values(ValueCount) = Cell.Value2
                End If
            End If
        End If
    Next Cell

    If ValueCount = 0 Then
        MsgBox "В исходном диапазоне нет непустых значений.", _
               vbInformation, MACRO_REPOSITORY_NAME
        Exit Sub
    End If

    Randomize

    For i = ValueCount To 2 Step -1
        j = Int(Rnd() * i) + 1

        TemporaryValue = Values(i)
        Values(i) = Values(j)
        Values(j) = TemporaryValue
    Next i

    If OutputCount <= 0 Or OutputCount > ValueCount Then
        ActualOutputCount = ValueCount
    Else
        ActualOutputCount = OutputCount
    End If

    If VerticalOutput Then
        ReDim OutputArray(1 To ActualOutputCount, 1 To 1)

        For i = 1 To ActualOutputCount
            OutputArray(i, 1) = Values(i)
        Next i

        FirstOutputCell.Resize(ActualOutputCount, 1).Value = OutputArray
    Else
        ReDim OutputArray(1 To 1, 1 To ActualOutputCount)

        For i = 1 To ActualOutputCount
            OutputArray(1, i) = Values(i)
        Next i

        FirstOutputCell.Resize(1, ActualOutputCount).Value = OutputArray
    End If
End Sub
