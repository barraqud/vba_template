'@Folder("VBAProject")
Option Explicit

Private this As RegExp

Public Sub Init(Pattern As String, Optional isGlobal As Boolean = True, Optional IgnoreCase As Boolean = False)
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
    Dim j As Long
    For j = LBound(sourceArray) To UBound(sourceArray)
        If Not this.Test(sourceArray(j)) Then GoTo Continue
        On Error GoTo Continue
        i = i + 1
        ReDim Preserve outputArr(1 To i)
        outputArr(i) = True
        If getIdx Then outputArr(i) = j
Continue:
    Next
    TestArray = outputArr
End Function

Public Function ExecuteString(ByVal sourceString As String) As Object
    Set ExecuteString = this.Execute(sourceString)
End Function

Public Function ExecuteArray(ByVal sourceArray As Variant) As Variant
    Dim outputArr As Variant
    Dim curr As Variant
    Dim i As Long
    For Each curr In sourceArray
        If Not this.Test(curr) Then GoTo Continue
        On Error GoTo Continue
        i = i + 1
        ReDim Preserve outputArr(1 To i)
        Set outputArr(i) = this.Execute(curr)
Continue:
    Next
    ExecuteArray = outputArr
    On Error GoTo 0
End Function