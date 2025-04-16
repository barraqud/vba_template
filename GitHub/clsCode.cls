'@Folder("GitHub")
Option Explicit

Private Const DirExclude = "Dev"
Private Const FileExclude = "README"

Public Components As Dictionary
'Только пути к старым файлам
Public OldTree As Dictionary

Private Self As VBIDE.VBProject

Private Function Create(ByVal ModName As String, ByVal content As String) As VBIDE.CodeModule
    Dim vbComp As VBIDE.VBComponent
    Dim vbType As vbext_ComponentType
    vbType = GetModuleType(ModName)
    Set vbComp = Self.VBComponents.Add(vbType)
    Dim Name As String
    Name = GetModuleName(ModName)
    vbComp.Name = Name
    Set Create = vbComp.CodeModule
End Function

Public Sub Update(ByVal ModName As String, ByVal content As String, Optional clearContent As Boolean = True)
    Dim vbMod As VBIDE.CodeModule
    Set vbMod = GetModuleByName(ModName)
    If vbMod Is Nothing Then Set vbMod = Create(ModName, content)
    If Err.Number <> 0 Then Exit Sub
    With vbMod
        If clearContent Then Call .DeleteLines(1, .CountOfLines)
        .AddFromString content
    End With
End Sub

Public Function Read(ByVal ModName As String, Optional startPos As Long = 1, Optional endPos As Long = 0) As String
    Dim vbMod As VBIDE.CodeModule
    Set vbMod = GetModuleByName(ModName)
    If Err.Number <> 0 Then Exit Function
    With vbMod
        If Not endPos Then endPos = .CountOfLines
        Read = .Lines(startPos, endPos)
    End With
End Function

Public Function Find(vbMod As VBIDE.CodeModule, search As String, Optional startPos As Long = 1) As Long
    Dim s As String
    Dim i As Long
    Dim RE As New RegExp
    RE.Pattern = search
    RE.Global = True
    RE.IgnoreCase = False
    
    With vbMod
        For i = startPos To .CountOfLines
            If RE.Test(.Lines(i, 1)) Then
                Find = i
                Exit For
            End If
        Next
    End With
End Function

Public Sub Compare(Optional NewData As Dictionary, Optional isPushing As Boolean = True)
    If isPushing Then
        Set Components = DropOld(NewData, Components)
    Else
        UpdateAll NewData
                Init
    End If
End Sub

Private Sub UpdateAll(ByVal ModData As Dictionary)
    Dim key As Variant
    Dim ModName As String
    Dim vbMod As VBIDE.CodeModule
    For Each key In ModData.Keys
        If CheckExclude(key) Then GoTo Continue
        Update key, ModData(key)
Continue:
    Next
    For Each key In Components.Keys
        If ModData.Exists(key) Or CheckExclude(key) Then GoTo Continue2
        ModName = GetModuleName(key)
        Set vbMod = GetModuleByName(ModName)
        Self.VBComponents.Remove vbMod.Parent
Continue2:
    Next
End Sub

Private Sub Init()
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
            path = dir & vMod.Name & SetModuleExt(vMod)
            If CheckExclude(path) Then GoTo Continue
            Set dict = New Dictionary
            dict.Add "path", path
            dict.Add "content", content
            dict.Add "mode", "100644"
            dict.Add "type", "blob"
            Components.Add path, dict
        End If
Continue:
    Next
End Sub

Private Function CheckExclude(ByVal ModName As String) As Boolean
    Dim RE As New clsRE
    Dim dir As String
    Dim filename As String
    Dim curr As Variant
    Dim ExcludeList As Variant
    RE.Init "(([a-z0-9_-]+)\/)?([a-z0-9_-]+)(\.(bas|cls|frm))", False, True
    On Error GoTo BeforeExit
    With RE.ExecuteString(ModName).Item(0).submatches
        dir = .Item(1)
        filename = .Item(2)
    End With
    For Each curr In Split(DirExclude, ";")
        If InStr(1, dir, curr, vbTextCompare) Then GoTo BeforeExit
    Next
    For Each curr In Split(FileExclude, ";")
        If InStr(1, filename, curr, vbTextCompare) Then GoTo BeforeExit
    Next
    Exit Function
BeforeExit:
    CheckExclude = True
    Exit Function
End Function

Private Function SetModuleDir(Line As String) As String
    Dim Match As String
    Dim RE As New clsRE
    RE.Init "\'\@Folder\(\""(.+)\""\)", False, True
    If Not RE.TestString(Line) Then Exit Function
    On Error GoTo BeforeExit
    Match = RE.ExecuteString(Line).Item(0).submatches.Item(0)
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

Private Function GetModuleType(ByVal ModName As String) As vbext_ComponentType
    Dim Match As String
    Dim RE As New clsRE
    RE.Init ".*(\.(bas|cls|frm))", False, True
    Match = RE.ExecuteString(ModName).Item(0).submatches.Item(0)
    Select Case Match
    Case ".cls":  GetModuleType = vbext_ct_ClassModule
    Case ".bas":  GetModuleType = vbext_ct_StdModule
    Case ".frm":  GetModuleType = vbext_ct_MSForm
    Case Else: Err.Raise vbObjectError + 500, "clsCode.SetModuleType", "Неправильное расширение файла"
    End Select
End Function

Private Function GetModuleName(ByVal ModName As String) As String
    Dim Match As String
    Dim RE As New clsRE
    RE.Init "([a-z0-9_-]+\/)?([a-z0-9_-]+)(\.(bas|cls|frm))", False, True
    GetModuleName = RE.ExecuteString(ModName).Item(0).submatches.Item(1)
End Function

Private Function GetModuleByName(moduleName As String) As VBIDE.CodeModule
    Dim project As VBIDE.VBProject
    Dim curr As Variant
    Set project = ActiveWorkbook.VBProject
    For Each curr In project.VBComponents
        If StrComp(curr.Name, moduleName, vbTextCompare) And ( _
           curr.Type = vbext_ct_ClassModule Or curr.Type = vbext_ct_StdModule Or curr.Type = vbext_ct_MSForm _
           ) Then
            Set GetModuleByName = curr.CodeModule
            Exit For
        End If
    Next
End Function

Private Function DropOld(oldDict As Dictionary, newDict As Dictionary) As Dictionary
    Dim dict As Dictionary
    Dim key As Variant
    Set DropOld = New Dictionary
    For Each key In oldDict
        If InStr(1, key, "README", vbTextCompare) Then GoTo Continue
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

Private Sub Class_Initialize()
    Init
End Sub

