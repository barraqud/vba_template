'@Folder("VBAProject")
Option Explicit

Public chunks As New Dictionary
Private p_range As Range

Public Property Get count() As Long
    count = chunks.count
End Property

Private Sub Class_Initialize()
    Dim headerRng As Range
    'todo: set range based on merge cells
    Set headerRng = RangeByName("ShtMainData_Header", ShtMainData.Cells(1).CurrentRegion.Rows(3))
    Set p_range = RangeExclude(headerRng.CurrentRegion, headerRng)
End Sub

Sub init(ColNum As Long, Optional ColSize As Long)
    Dim FilterRng As Range
    Dim ChunkRng As Range
    Dim lastRow As Long, i As Long

    With p_range
        lastRow = .Rows.count + .row - 1
        'note: counts rm columns from 1
        Set FilterRng = .Columns(ColNum)
        If ColSize > 1 Then Set FilterRng = FilterRng.Resize(ColumnSize:=ColSize)
        With FilterRng.SpecialCells(xlCellTypeConstants, 7)

            For i = .Areas.count To 1 Step -1
                Set ChunkRng = .Areas(i).Resize(lastRow - .Areas(i).row + 1, p_range.Columns.count)
                Call Add(ChunkRng)

                lastRow = .Areas(i).row - 1
            Next
        End With
    End With
End Sub

Private Sub Add(ChunkRng As Range)
    Dim ChunkNew As cls
    ChunkNew.ID = ChunkRng.Cells(1).Value2
    If chunks.Exists(ChunkNew.ID) Then GoTo MergeChunks
    ChunkNew.Range = ChunkRng
    ChunkNew.Value = ChunkRng.Value2
    'If Not ChunkNew Is Nothing Then
    chunks.Add ChunkNew.ID, ChunkNew
    Exit Sub
MergeChunks:
    MsgBox Join(Array("Есть несколько РМ с одним ID", ChunkRng.Address, chunks(ChunkNew.ID).Range.Address), vbNewLine), Title:="Дублирование РМ"
End Sub
