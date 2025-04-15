'@Folder("GitHub")
Option Explicit

Private Const DirExclude = "Dev"

Public Components As Dictionary
'Только пути к старым файлам
Public OldTree As Dictionary

Private Self As vbide.VBProject

Private Function SetModuleDir(Line As String) As String
    Dim Match As String
    Dim Re As New clsRE
    Re.Init "\'\@Folder\(\""(.+)\""\)", False, True
    If Not Re.TestString(Line) Then Exit Function
    On Error GoTo BeforeExit
    Match = Re.ExecuteString(Line).Item(0).submatches.Item(0)
'    If Not Components.Exists(Match) Then
'        If DirPass(Match) Then GoTo BeforeExit
'        Dim dict As New Dictionary
'        dict.Add "path", Match
'        dict.Add "mode", "040000"
'        dict.Add "type", "tree"
'        Components.Add Match, dict
'    End If
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
    Dim curr As Variant
    Set project = ActiveWorkbook.VBProject
    For Each curr In project.VBComponents
        If curr.Name = moduleName And ( _
           curr.Type = vbext_ct_ClassModule Or curr.Type = vbext_ct_StdModule Or curr.Type = vbext_ct_MSForm _
           ) Then
            Set GetModuleByName = curr.CodeModule
            Exit For
        End If
    Next
End Function

Public Sub Update(vbMod As vbide.CodeModule, searchStr As String, Optional replaceStr As String = vbNullString, Optional isGlobal As Boolean = False)
    Dim lineNum As Long
    lineNum = Find(vbMod, searchStr)
    On Error Resume Next
    With vbMod
        If Len(replaceStr) Then
            .ReplaceLine lineNum, replaceStr
        End If
    End With
End Sub

Public Function Read(vbMod As vbide.CodeModule, Optional startPos As Long = 1, Optional endPos As Long = 0) As String
    If Err Then Exit Function
    With vbMod
        If Not endPos Then endPos = .CountOfLines
        Read = .Lines(startPos, endPos - startPos)
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

Public Sub Compare(Optional Other As Dictionary, Optional isPushing As Boolean = True)
    If isPushing Then
        Set Components = DropOld(Other, Components)
    Else
        '        Set Compare = DropOld(Components, Other)
        '        Set Components = Other
    End If
End Sub

Private Function DropOld(oldDict As Dictionary, newDict As Dictionary) As Dictionary
    Dim dict As Dictionary
    Dim key As Variant
    Set DropOld = New Dictionary
    For Each key In oldDict
        If InStr(1, key, "README", vbTextCompare) Or Not InStr(1, key, "/", vbTextCompare) Then GoTo Continue
        If newDict.Exists(key) Then
            Set dict = newDict(key)
            DropOld.Add key, dict
        Else
            Set dict = New Dictionary
            dict.Add "path", key
            dict.Add "content", Null
            dict.Add "mode", "100644"
            dict.Add "type", "blob"
            DropOld.Add key, dict
        End If
Continue:
    Next
    For Each key In newDict
        If Not DropOld.Exists(key) And Not InStr(1, key, "README", vbTextCompare) Then DropOld.Add key, newDict(key)
    Next
End Function

Private Function DirPass(dir As String) As Boolean
    Dim pass As Variant
    For Each pass In Split(DirExclude, ";")
        If dir = pass & "/" Or dir = pass Then
            DirPass = True
            Exit Function
        End If
    Next
End Function

Private Sub Class_Initialize()
    Dim dict As Dictionary
    Dim dir As String
    Dim path As String
    Dim content As String
    Dim vMod As Variant
    Set Components = New Dictionary
    Components.CompareMode = TextCompare
    Set Self = ActiveWorkbook.VBProject
    For Each vMod In Self.VBComponents
        If vMod.Type = vbext_ct_ClassModule Or vMod.Type = vbext_ct_StdModule Or vMod.Type = vbext_ct_MSForm Then
            With vMod.CodeModule
                content = .Lines(1, .CountOfLines)
            End With
            dir = SetModuleDir(content)
            If DirPass(dir) Then GoTo Continue
            Set dict = New Dictionary
            path = dir & vMod.Name & SetModuleExt(vMod)
            dict.Add "path", path
            dict.Add "content", content
            dict.Add "mode", "100644"
            dict.Add "type", "blob"
            Components.Add path, dict
        End If
Continue:
    Next
End Sub

