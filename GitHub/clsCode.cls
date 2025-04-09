'@Folder("GitHub")
Option Explicit

Public Components As Collection

Private Const DirExclude = "Dev;_key"

Private Self As vbide.VBProject
Private Path As String
Private Extension As String

Private Function SetModuleDir(Line As String) As String
    Dim Match As String
    Dim Re As New clsRE
    Re.init "\'\@Folder\(\""(.+)\""\)", False, True
    If Not Re.TestString(Line) Then Exit Function
    On Error GoTo BeforeExit
    Match = Re.ExecuteString(Line).Item(0).submatches.Item(0)
    If Match <> "VBAProject" Then SetModuleDir = Match & "/"
BeforeExit:
End Function

Private Function SetModuleExt(ByVal vbMod As VBComponent) As String
    Select Case vbMod.Type
    Case vbext_ct_ClassModule: SetModuleExt = ".cls"
    Case vbext_ct_StdModule: SetModuleExt = ".bas"
    Case vbext_ct_MSForm: SetModuleExt = ".frm"
    Case Else: Exit Function
    End Select
End Function

Private Function GetModuleByName(moduleName As String) As vbide.CodeModule
    Dim project As vbide.VBProject
    Dim Curr As Variant
    Set project = ActiveWorkbook.VBProject
    For Each Curr In project.VBComponents
        If Curr.Name = moduleName And ( _
           Curr.Type = vbext_ct_ClassModule Or Curr.Type = vbext_ct_StdModule Or Curr.Type = vbext_ct_MSForm _
           ) Then
            Set GetModuleByName = Curr.CodeModule
            Exit For
        End If
    Next
End Function

Public Function Update(vbMod As vbide.CodeModule, searchStr As String, Optional replaceStr As String = vbNullString, Optional isGlobal As Boolean = False) As String
    Dim lineNum As Long
    lineNum = Find(vbMod, searchStr)
    On Error Resume Next
    With vbMod
        If Len(replaceStr) Then
            .ReplaceLine lineNum, replaceStr
        End If
    End With
End Function

Public Function Read(vbMod As vbide.CodeModule, Optional startPos As Long = 1, Optional endPos As Long = 0) As String
    If Err Then Exit Function
    With vbMod
        If Not endPos Then endPos = .CountOfLines
        Debug.Print .Lines(startPos, endPos - startPos)
    End With
End Function

Public Function Find(vbMod As vbide.CodeModule, search As String, Optional startPos As Long = 1) As Long
    Dim s As String
    Dim i As Long
    Dim Re As New RegExp
    Re.Pattern = search
    Re.Global = True
    Re.IgnoreCase = False
    
    With vbMod
        For i = startPos To .CountOfLines
            If Re.Test(.Lines(i, 1)) Then
                Find = i
                Exit For
            End If
        Next
    End With
End Function

Private Function DirPass(Dir As String) As Boolean
    Dim pass As Variant
    For Each pass In Split(DirExclude, ";")
        If Dir = pass & "/" Then
            DirPass = True
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    Dim dict As Dictionary
    Dim Dir As String
    Dim content As String
    Dim vMod As Variant
    Set Components = New Collection
    Set Self = ActiveWorkbook.VBProject
    For Each vMod In Self.VBComponents
        If vMod.Type = vbext_ct_ClassModule Or vMod.Type = vbext_ct_StdModule Or vMod.Type = vbext_ct_MSForm Then
            With vMod.CodeModule
                content = .Lines(1, .CountOfLines)
            End With
            Dir = SetModuleDir(content)
            If DirPass(Dir) Then GoTo Continue
            Set dict = New Dictionary
            dict.Add "path", Dir & vMod.Name & SetModuleExt(vMod)
            dict.Add "content", content
            dict.Add "mode", "100644"
            dict.Add "type", "blob"
            Components.Add dict
        End If
Continue:
    Next
End Sub

