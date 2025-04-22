'@Folder("VBAProject")
Option Explicit

Public ID As String
Public Range As Range
Public Value As Variant

Public Property Get Count() As Long
    Count = Me.Range.Rows.Count
End Property

Public Sub Create(ChunkRng As Range)
    With Me
        .ID = ChunkRng.Cells(1).Value2
        Set .Range = ChunkRng
        .Value = ChunkRng.Value2
    End With
End Sub
