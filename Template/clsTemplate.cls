'@Folder("Template")
Option Explicit

Public name As String
Public Path As String

Private RE_Prefix As New clsRE
Private RE_Variable As New clsRE
Private Variables As New Dictionary
Private tables As New Collection

Public Sub Init(filepath As String)
    With RE_Prefix.ExecuteString(filepath)
        If .Count = 0 Then GoTo Failed
        Path = filepath
        name = .Item(0).submatches.Item(0)
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
    Dim curr As Variant
    Dim i As Long
    On Error GoTo BeforeExit
    Set wDoc = wApp.Read(Me.Path)
    SwitchAutomation
    With wDoc
        'Constant variables
        If .Shapes.Count = 0 Then GoTo NoVariables
        For i = 1 To .Shapes.Count
            On Error GoTo NextVar
            Set curr = .Shapes(i)
            With RE_Variable.ExecuteString(curr.TextFrame.TextRange.text)
                If .Count = 0 Then GoTo NextVar
                Variables.Add .Item(0).submatches.Item(0), i
            End With
NextVar:
            If Err.Number <> 0 Then Err.Clear
        Next
        outputDict.Add "Variables", Variables
NoVariables:
        If .tables.Count = 0 Then GoTo NoTables
        Dim tables As New Collection
        Set tables = New Collection
        For Each curr In .tables
            Set tableDict = New Dictionary
            For i = 1 To curr.Columns.Count
                Set wRng = curr.cell(curr.Rows.Count, i).Range
                With RE_Variable.ExecuteString(wRng.text)
                    If .Count = 0 Then GoTo NextCol
                    tableDict.Add .Item(0).submatches.Item(0), i
                End With
NextCol:
                If Err.Number <> 0 Then Err.Clear
            Next
            If tableDict.Count > 0 Then tables.Add tableDict
        Next
        outputDict.Add "Tables", tables
NoTables:
    End With
    Set Parse = outputDict
BeforeExit:
    SwitchAutomation True
End Function


Public Sub Fill()
'    ShtMainData
End Sub

Private Sub Class_Initialize()
    RE_Prefix.Init RePattern_prefix, False, True
    RE_Variable.Init RePattern_variable, False, True
End Sub

