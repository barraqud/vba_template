'@Folder("Options")
Option Explicit

Public Title As String
Public Data As String
Public SplitArgs As String
Public FilterArgs As String
Public GroupArgs As String
Public SortArgs As String

Friend Function Create(ParamArray Options() As Variant) As ClsOptionRow
    With Me
        .Title = Options(1)
        .Data = Options(2)
        .SplitArgs = Options(3)
        .FilterArgs = Options(4)
        .GroupArgs = Options(5)
        .SortArgs = Options(6)
    End With
    Set Create = Me
End Function
