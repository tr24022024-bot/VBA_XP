Attribute VB_Name = "modLevenshtein"
Option Explicit

Public Function LevenshteinDistance( _
    ByVal FirstText As String, _
    ByVal SecondText As String) As Long

    Dim FirstLength As Long
    Dim SecondLength As Long
    Dim Matrix() As Long
    Dim i As Long
    Dim j As Long
    Dim Cost As Long
    Dim DeleteCost As Long
    Dim InsertCost As Long
    Dim ReplaceCost As Long

    FirstText = LCase$(FirstText)
    SecondText = LCase$(SecondText)

    FirstLength = Len(FirstText)
    SecondLength = Len(SecondText)

    If FirstLength = 0 Then
        LevenshteinDistance = SecondLength
        Exit Function
    End If

    If SecondLength = 0 Then
        LevenshteinDistance = FirstLength
        Exit Function
    End If

    ReDim Matrix(0 To FirstLength, 0 To SecondLength)

    For i = 0 To FirstLength
        Matrix(i, 0) = i
    Next i

    For j = 0 To SecondLength
        Matrix(0, j) = j
    Next j

    For i = 1 To FirstLength
        For j = 1 To SecondLength
            If Mid$(FirstText, i, 1) = Mid$(SecondText, j, 1) Then
                Cost = 0
            Else
                Cost = 1
            End If

            DeleteCost = Matrix(i - 1, j) + 1
            InsertCost = Matrix(i, j - 1) + 1
            ReplaceCost = Matrix(i - 1, j - 1) + Cost

            Matrix(i, j) = Repository_Minimum3( _
                DeleteCost, InsertCost, ReplaceCost)
        Next j
    Next i

    LevenshteinDistance = Matrix(FirstLength, SecondLength)
End Function

Private Function Repository_Minimum3( _
    ByVal FirstValue As Long, _
    ByVal SecondValue As Long, _
    ByVal ThirdValue As Long) As Long

    Dim ResultValue As Long

    ResultValue = FirstValue

    If SecondValue < ResultValue Then
        ResultValue = SecondValue
    End If

    If ThirdValue < ResultValue Then
        ResultValue = ThirdValue
    End If

    Repository_Minimum3 = ResultValue
End Function
