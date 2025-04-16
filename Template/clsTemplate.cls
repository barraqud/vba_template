'@Folder("Template")
Option Explicit

Public Name As String
Public path As String
Public Saved As Dictionary

Private RE_Prefix As New clsRE
Private RE_Variable As New clsRE
Private Variables As New Dictionary
Private Tables As New Collection

Public Sub Import(Optional ByVal filepath As String)
    If Len(filepath) = 0 Then
        SetFile
    Else
        path = filepath
    End If
    With RE_Prefix.ExecuteString(path)
        If .count = 0 Then GoTo Failed
        Name = .Item(0).submatches.Item(0)
    End With
        
    Exit Sub
Failed:
    Err.Raise Number:=vbNullChar + 1000, Description:="Path not match"
End Sub

'(Variables as Dictionary, Tables As Collection of Dictionary) -> Key = VarName, Item = index or column
Public Sub ParseFile()
    Dim outputDict As New Dictionary
    Dim tableDict As Dictionary
    Dim wApp As New clsWord
    Dim wDoc As Document
    Dim wRng As Word.Range
    Dim curr As Variant
    Dim i As Long
    On Error GoTo BeforeExit
    Set wDoc = wApp.Read(Me.path)
    SwitchAutomation
    With wDoc
        'Constant variables
        If .Shapes.count = 0 Then GoTo NoVariables
        For i = 1 To .Shapes.count
            On Error GoTo NextVar
            Set curr = .Shapes(i)
            With RE_Variable.ExecuteString(curr.TextFrame.TextRange.text)
                If .count = 0 Then GoTo NextVar
                Variables.Add .Item(0).submatches.Item(0), i
            End With
NextVar:
            If Err.Number <> 0 Then Err.Clear
        Next
        outputDict.Add "Variables", Variables
NoVariables:
        If .Tables.count = 0 Then GoTo NoTables
        Dim Tables As New Collection
        Set Tables = New Collection
        For Each curr In .Tables
            Set tableDict = New Dictionary
            For i = 1 To curr.Columns.count
                Set wRng = curr.cell(curr.Rows.count, i).Range
                With RE_Variable.ExecuteString(wRng.text)
                    If .count = 0 Then GoTo NextCol
                    tableDict.Add .Item(0).submatches.Item(0), i
                End With
NextCol:
                If Err.Number <> 0 Then Err.Clear
            Next
            If tableDict.count > 0 Then Tables.Add tableDict
        Next
        outputDict.Add "Tables", Tables
NoTables:
    End With
    Set Saved = outputDict
    Stop
BeforeExit:
    SwitchAutomation True
End Sub

Public Function ParseSaved() As Dictionary
Dim Chunks As New clsChunks
Dim rng As Range
With ShtOptionsMainData.Cells(1).CurrentRegion
    Set rng = RangeExclude(.Cells, .Rows(1))
End With
Chunks.Init 1, , rng
Set ParseSaved = Chunks.Chunks
End Function

Public Sub Fill()
    ShtOptionsMainData.Cells(1).CurrentRegion
End Sub

Private Sub SetFile(Optional filepath As String)
    Dim i As Long
    If Len(filepath) = 0 Then
        ' Open the file dialog
        With Application.FileDialog(msoFileDialogOpen)
            .InitialFileName = FilePathCurrent
            .AllowMultiSelect = False
            .Show
            path = .SelectedItems(1)
        End With
    End If
End Sub

Private Function OptionsCreate(OptsArr As Variant) As Collection
    Set OptionsCreate = New Collection
    Dim dict As Dictionary
    Dim args(1 To 5) As Variant
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(OptsArr)
        Set dict = New Dictionary
        dict.Add "VarName", OptsArr(i, 1)
        dict.Add "InitData", OptsArr(i, 2)
        dict.Add "Grouping", OptsArr(i, 3)
        dict.Add "AutoValue", OptsArr(i, 4)
        dict.Add "Modificator", OptsArr(i, 5)
        For j = 6 To UBound(OptsArr, 2)
            args(j - 5) = OptsArr(i, j)
        Next
        dict.Add "Args", args
        OptionsCreate.Add dict
    Next
End Function

Private Sub Class_Initialize()
    RE_Prefix.Init RePattern_prefix, False, True
    RE_Variable.Init RePattern_variable, False, True
End Sub
