'@Folder("VBAProject")
Option Explicit

Private this As RegExp

Public Sub init(Pattern As String, Optional isGlobal As Boolean = True, Optional IgnoreCase As Boolean = False)
    Set this = New RegExp
    With this
        .Pattern = Pattern
        .Global = isGlobal
        .IgnoreCase = IgnoreCase
    End With
End Sub

Public Function TestString(ByVal sourceString As String) As Boolean
    TestString = this.Test(sourceString)
End Function

Public Function TestArray(ByVal sourceArray As Variant, Optional getIdx As Boolean = False) As Variant
    Dim outputArr As Variant
    Dim i As Long
    For i = LBound(sourceArray) To UBound(sourceArray)
        If Not this.Test(sourceArray(i)) Then GoTo Continue
        If IsEmpty(outputArr) Then
            ReDim outputArr(1 To 1)
        Else
            ReDim Preserve outputArr(LBound(outputArr) To UBound(outputArr) + 1)
        End If
        outputArr(UBound(outputArr)) = True
        If getIdx Then outputArr(UBound(outputArr)) = i
Continue:
    Next
    TestArray = outputArr
End Function

Public Function ExecuteString(ByVal sourceString As String) As Object
    Set ExecuteString = this.Execute(sourceString)
End Function

Public Function ExecuteArray(ByVal sourceArray As Variant) As Variant
    Dim outputArr As Variant
    Dim i As Long
    For i = LBound(sourceArray) To UBound(sourceArray)
        If Not this.Test(sourceArray(i)) Then GoTo Continue
        If IsEmpty(outputArr) Then
            ReDim outputArr(1 To 1)
        Else
            ReDim Preserve outputArr(LBound(outputArr) To UBound(outputArr) + 1)
        End If
        Set outputArr(UBound(outputArr)) = this.Execute(sourceArray(i))
Continue:
    Next
    ExecuteArray = outputArr
End Function