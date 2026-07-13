Attribute VB_Name = "modSelfTests"
Option Explicit

Public Sub Repository_SelfTest()
    Dim PassedCount As Long
    Dim FailedCount As Long

    Repository_AssertEqualLong _
        TestName:="Levenshtein: bicycle/biccicle", _
        ExpectedValue:=2, _
        ActualValue:=LevenshteinDistance("bicycle", "biccicle"), _
        PassedCount:=PassedCount, _
        FailedCount:=FailedCount

    Repository_AssertEqualLong _
        TestName:="Levenshtein: truck/truk", _
        ExpectedValue:=1, _
        ActualValue:=LevenshteinDistance("truck", "truk"), _
        PassedCount:=PassedCount, _
        FailedCount:=FailedCount

    Repository_AssertEqualString _
        TestName:="NormalizeText", _
        ExpectedValue:="one two", _
        ActualValue:=Repository_NormalizeText("  one   two  "), _
        PassedCount:=PassedCount, _
        FailedCount:=FailedCount

    MsgBox "Самопроверка завершена." & vbCrLf & _
           "Успешно: " & PassedCount & vbCrLf & _
           "Ошибок: " & FailedCount, _
           IIf(FailedCount = 0, vbInformation, vbExclamation), _
           MACRO_REPOSITORY_NAME
End Sub

Private Sub Repository_AssertEqualLong( _
    ByVal TestName As String, _
    ByVal ExpectedValue As Long, _
    ByVal ActualValue As Long, _
    ByRef PassedCount As Long, _
    ByRef FailedCount As Long)

    If ExpectedValue = ActualValue Then
        PassedCount = PassedCount + 1
        Debug.Print "PASS: " & TestName
    Else
        FailedCount = FailedCount + 1
        Debug.Print "FAIL: " & TestName & _
                    "; expected=" & ExpectedValue & _
                    "; actual=" & ActualValue
    End If
End Sub

Private Sub Repository_AssertEqualString( _
    ByVal TestName As String, _
    ByVal ExpectedValue As String, _
    ByVal ActualValue As String, _
    ByRef PassedCount As Long, _
    ByRef FailedCount As Long)

    If ExpectedValue = ActualValue Then
        PassedCount = PassedCount + 1
        Debug.Print "PASS: " & TestName
    Else
        FailedCount = FailedCount + 1
        Debug.Print "FAIL: " & TestName & _
                    "; expected=" & ExpectedValue & _
                    "; actual=" & ActualValue
    End If
End Sub
