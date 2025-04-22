'@Folder("GitHub")
Option Explicit

Private Type TComponentProps
    cmpName As String
    cmpType As vbext_ComponentType
    cmpExt As String
    dir As String
    filePath As String
    isInExclude As Boolean
    isInComponents As Boolean
End Type

Private Self As vbide.VBProject
Private Const DirExclude = "Dev"
Private Const FileExclude = "README;JsonConverter"

Public Components As Dictionary
'Только пути к старым файлам
Public OldTree As Dictionary

Public Function Read(ByVal ModName As String, Optional startPos As Long = 1, Optional endPos As Long = 0) As String
    Dim vbMod As vbide.CodeModule
    Set vbMod = Self.VBComponents(ModName).CodeModule
    If Err.Number <> 0 Then Exit Function
    With vbMod
        If Not endPos Then endPos = .CountOfLines
        Read = .Lines(startPos, endPos)
    End With
End Function

Public Function Find(vbMod As vbide.CodeModule, search As String, Optional startPos As Long = 1) As Long
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

Public Sub Compare(Optional newData As Dictionary, Optional isPushing As Boolean = True)
    If isPushing Then
        Set Components = ComponentsDropOld(newData, Components)
    Else
        UpdateComponents newData
        componentsRemove newData
    End If
End Sub

'Private
Private Sub Init()
    Dim dict As Dictionary
    Dim dir As String
    Dim path As String
    Dim content As String
    Dim vMod As vbide.VBComponent
    Set Components = New Dictionary
    Components.CompareMode = TextCompare
    Set Self = ActiveWorkbook.VBProject
    For Each vMod In Self.VBComponents
        If vMod.Type <> vbext_ct_ClassModule _
           And vMod.Type <> vbext_ct_StdModule _
           And vMod.Type <> vbext_ct_MSForm Then GoTo Continue
        With vMod.CodeModule
            content = .Lines(1, .CountOfLines)
        End With
        dir = ComponentGetDir(content)
        path = dir & vMod.Name & ComponentGetExt(vMod)
        If ComponentIgnore(path) Then GoTo Continue
        Set dict = ComponentCreatePayload(content, path)
        Components.Add path, dict
Continue:
    Next
End Sub

Private Function ComponentCreatePayload(ByVal content As String, ByVal path As String) As Dictionary
    Set ComponentCreatePayload = New Dictionary
    ComponentCreatePayload.Add "path", path
    ComponentCreatePayload.Add "content", content
    ComponentCreatePayload.Add "mode", "100644"
    ComponentCreatePayload.Add "type", "blob"
End Function

Private Sub UpdateComponents(ByVal newData As Dictionary)
    Dim key As Variant
    Dim cmp As vbide.VBComponent
    Dim newCmp As vbide.VBComponent
    Dim tempProject As Workbook
    Dim props As TComponentProps
    On Error GoTo BeforeExit
    SwitchAutomation
    
    Set tempProject = Application.Workbooks.Add
    For Each key In newData.Keys
        props = ComponentGetProps(key)
        If props.isInExclude Then GoTo Continue
        If props.cmpType = vbext_ct_MSForm Then
        Set cmp = Self.VBComponents(cmp.Name)
        With cmp.CodeModule
            .DeleteLines 1, .CountOfLines
            .AddFromString newData(key)
        End With
        GoTo Continue
        End If
        Set cmp = tempProject.VBProject.VBComponents.Add(props.cmpType)
        With cmp
            .Name = props.cmpName
            With .CodeModule
                .DeleteLines 1, .CountOfLines
                .AddFromString newData(key)
            End With
            .Export props.filePath
        End With
        Set cmp = Nothing
        On Error Resume Next
        With Self.VBComponents
            Set cmp = .Item(props.cmpName)
            If cmp Is Nothing Then GoTo NoComponent
            cmp.Name = "DEL_" & props.cmpName
            DoEvents
            .Remove cmp
            DoEvents
NoComponent:
            .Import props.filePath
        End With
        Kill props.filePath
        Debug.Print "updated: " & props.cmpName
        ComponentClearProps props
Continue:
    Next
    tempProject.Close False
    DoEvents

BeforeExit:
    If Err.Number <> 0 Then Debug.Print "error: "; Err.Number & ", " & Err.Source & ": " & Err.Description
    SwitchAutomation True
End Sub

Private Sub componentsRemove(newData As Dictionary)
    Dim key As Variant
    Dim cmp As vbide.VBComponent
    Dim props As TComponentProps
    
    On Error GoTo Continue
    SwitchAutomation
    For Each key In Components.Keys
        props = ComponentGetProps(key)
        If props.isInExclude Then GoTo Continue
        Debug.Print "not exist: " & props.cmpName
        With ThisWorkbook.VBProject.VBComponents
            Set cmp = .Item(props.cmpName)
            .Remove cmp
        End With
        Debug.Print "deleted: " & props.cmpName
        ComponentClearProps props
Continue:
    Next
BeforeExit:
    If Err.Number <> 0 Then Debug.Print "error: "; Err.Number & ", " & Err.Source & ": " & Err.Description
    SwitchAutomation True
End Sub

Private Function ComponentIgnore(ByVal cmpPath As String) As Boolean
    Dim RE As New clsRE
    Dim dir As String
    Dim fileName As String
    Dim curr As Variant
    Dim ExcludeList As Variant
    RE.Init "(([a-z0-9_-]+)\/)?([a-z0-9_-]+)(\.(bas|cls|frm))", False, True
    On Error GoTo Exclude
    With RE.ExecuteString(cmpPath).Item(0).submatches
        dir = .Item(1)
        fileName = .Item(2)
    End With
    ComponentIgnore = ComponentIsExclude(dir, fileName)
    Exit Function
Exclude:
    ComponentIgnore = True
End Function

Private Function ComponentIsExclude(ByVal dir As String, ByVal cmpName As String) As Boolean
    Dim curr As Variant
    On Error GoTo Exclude
    For Each curr In Split(DirExclude, ";")
        If InStr(1, dir, curr, vbTextCompare) Then GoTo Exclude
    Next
    For Each curr In Split(FileExclude, ";")
        If InStr(1, cmpName, curr, vbTextCompare) Then GoTo Exclude
    Next
    Exit Function
Exclude:
    ComponentIsExclude = True
End Function

Private Function ComponentGetProps(ByVal ComponentPath As String) As TComponentProps
    Dim RE As New clsRE
    RE.Init "(([a-z0-9_-]+)\/)?([a-z0-9_-]+)(\.(bas|cls|frm))", False, True
    With RE.ExecuteString(ComponentPath).Item(0).submatches
        ComponentGetProps.dir = .Item(1)
        ComponentGetProps.cmpName = .Item(2)
        ComponentGetProps.filePath = FilePathCurrent(.Item(2) & .Item(3))
        ComponentGetProps.cmpExt = .Item(3)
        ComponentGetProps.cmpType = ComponentGetType(.Item(3))
        ComponentGetProps.isInExclude = ComponentIsExclude(.Item(1), .Item(2))
        ComponentGetProps.isInComponents = ComponentIsInOld(.Item(2))
    End With
End Function

Private Sub ComponentClearProps(ByRef props As TComponentProps)
    With props
        .cmpName = vbNullString
        .cmpType = Nothing
        .cmpExt = vbNullString
        .dir = vbNullString
        .fileName = vbNullString
        .isInExclude = False
        .isInComponents = False
    End With
End Sub

Private Function ComponentIsInOld(ByVal cmpName As String) As Boolean
    Dim cmp As vbide.VBComponent
    On Error Resume Next
    Set cmp = Self.VBComponents(cmpName)
    If Not cmp Is Nothing Then ComponentIsInOld = True
    Err.Clear
End Function

Private Function ComponentGetDir(Line As String) As String
    Dim Match As String
    Dim RE As New clsRE
    RE.Init "\'\@Folder\(\""(.+)\""\)", False, True
    If Not RE.TestString(Line) Then Exit Function
    On Error GoTo BeforeExit
    Match = RE.ExecuteString(Line).Item(0).submatches.Item(0)
    If Match <> "VBAProject" Then ComponentGetDir = Match & "/"
BeforeExit:
End Function

Private Function ComponentGetExt(ByVal vbMod As VBComponent) As String
    Select Case vbMod.Type
    Case vbext_ct_ClassModule: ComponentGetExt = ".cls"
    Case vbext_ct_StdModule: ComponentGetExt = ".bas"
    Case vbext_ct_MSForm: ComponentGetExt = ".frm"
    Case Else: Exit Function
    End Select
End Function

Private Function ComponentGetType(ByVal cmpType As String) As vbext_ComponentType
    Select Case cmpType
    Case ".cls":  ComponentGetType = vbext_ct_ClassModule
    Case ".bas":  ComponentGetType = vbext_ct_StdModule
    Case ".frm":  ComponentGetType = vbext_ct_MSForm
    Case Else: Err.Raise vbObjectError + 500, "clsCode.SetModuleType", "Неправильное расширение файла"
    End Select
End Function

Private Function ComponentsDropOld(oldDict As Dictionary, newDict As Dictionary) As Dictionary
    Dim dict As Dictionary
    Dim key As Variant
    Set ComponentsDropOld = New Dictionary
    For Each key In oldDict
        If ComponentIgnore(key) Then GoTo Continue
        If newDict.Exists(key) Then
            Set dict = newDict(key)
            ComponentsDropOld.Add key, dict
        Else
            Set dict = New Dictionary
            dict.Add "path", key
            dict.Add "content", Null
            dict.Add "mode", "100644"
            dict.Add "type", "blob"
            ComponentsDropOld.Add key, dict
        End If
Continue:
    Next
    For Each key In newDict
        If Not ComponentsDropOld.Exists(key) And Not InStr(1, key, "README", vbTextCompare) Then ComponentsDropOld.Add key, newDict(key)
    Next
End Function

Private Sub Class_Initialize()
    Init
End Sub

