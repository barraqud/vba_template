'@Folder("Template")
Option Explicit

Public Name As String
Public Path As String

Private RE_Prefix As New clsRE
Private RE_Variable As New clsRE
Private Variables As New Dictionary
Private tables As New Collection

Public Sub init(filepath As String)
    If IsMissing(filepath) Then
        SetFile
    Else
        Path = filepath
    End If
    With RE_Prefix.ExecuteString(Path)
        If .count = 0 Then GoTo Failed
        
        Name = .Item(0).submatches.Item(0)
    End With

    Exit Sub
Failed:
    Err.Raise Number:=vbNullChar + 1000, Description:="Path not match"
End Sub

'(Variables as Dictionary, Tables As Collection of Dictionary) -> Key = VarName, Item = index or column
Public Function Parse() As Dictionary
    Dim outputDict As New Dictionary
    Dim tableDict As Dictionary
    Dim wApp As New clsWord
    Dim wDoc As Document
    Dim wRng As Word.Range
    Dim Curr As Variant
    Dim i As Long
    On Error GoTo BeforeExit
    Set wDoc = wApp.Read(Me.Path)
    SwitchAutomation
    With wDoc
        'Constant variables
        If .Shapes.count = 0 Then GoTo NoVariables
        For i = 1 To .Shapes.count
            On Error GoTo NextVar
            Set Curr = .Shapes(i)
            With RE_Variable.ExecuteString(Curr.TextFrame.TextRange.text)
                If .count = 0 Then GoTo NextVar
                Variables.Add .Item(0).submatches.Item(0), i
            End With
NextVar:
            If Err.Number <> 0 Then Err.Clear
        Next
        outputDict.Add "Variables", Variables
NoVariables:
        If .tables.count = 0 Then GoTo NoTables
        Dim tables As New Collection
        Set tables = New Collection
        For Each Curr In .tables
            Set tableDict = New Dictionary
            For i = 1 To Curr.Columns.count
                Set wRng = Curr.cell(Curr.Rows.count, i).Range
                With RE_Variable.ExecuteString(wRng.text)
                    If .count = 0 Then GoTo NextCol
                    tableDict.Add .Item(0).submatches.Item(0), i
                End With
NextCol:
                If Err.Number <> 0 Then Err.Clear
            Next
            If tableDict.count > 0 Then tables.Add tableDict
        Next
        outputDict.Add "Tables", tables
NoTables:
    End With
    Set Parse = outputDict
BeforeExit:
    SwitchAutomation True
End Function

Public Sub Fill()
    With ShtMainData
        With RangeExclude(.Cells(1).CurrentRegion, .Names("").RefersToRange)
            .Select
        End With
    End With
End Sub

Private Sub SetFile(Optional filepath As String)
    Dim i As Long
    If Len(filepath) = 0 Then
        ' Open the file dialog
        With Application.FileDialog(msoFileDialogOpen)
            .InitialFileName = FilePathCurrent
            .AllowMultiSelect = False
            .Show
            Path = .SelectedItems(1)
        End With
    End If
End Sub

Private Sub Class_Initialize()
    RE_Prefix.init RePattern_prefix, False, True
    RE_Variable.init RePattern_variable, False, True
End Sub
