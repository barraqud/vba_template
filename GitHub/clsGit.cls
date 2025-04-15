'@Folder("GitHub")
Option Explicit

Private Type TCommitter
    Name As String
    email As String
End Type

Private Type TCommit
    branch As String
    message As String
    SHA As String
    treeSHA As String
End Type

Public Enum HttpMethod
    MGET
    MPOST
    MPATCH
End Enum

Private Const DefaultHeadSHA As String = "419ea2b11b9419885471b6fdc52994c1adbd20a5" 'start point
Private Const BranchNameCell As String = "A2:C6"

Private Const baseHeaders As String = "Accept: application/vnd.github+json;X-GitHub-Api-Version: 2022-11-28"
Private Const baseURL As String = "https://api.github.com/repos/barraqud"
Private Const key As String = ""
Private Head As TCommit
Private committer As TCommitter
Private CodeBase As clsCode

Public Property Get Author() As Dictionary
    Set Author = New Dictionary
    With committer
        Author.Add "name", .Name
        Author.Add "email", .email
    End With
End Property

Public Property Get Settings() As Dictionary
    Set Settings = New Dictionary
    With Head
        Settings.Add "Ветка", .branch
        Settings.Add "Описание закрепления", .message
        Settings.Add "ID закрепления", .SHA
    End With
    With committer
        Settings.Add "Имя автора", .Name
        Settings.Add "Почта автора", .email
    End With
End Property

Public Property Get SHADefault() As String
    SHADefault = DefaultHeadSHA
End Property

Public Property Get Status() As Object
    If Len(Head.branch) > 0 Then Set Status = fetchAuthorized(setURL("git", "refs", "heads", Head.branch), method:=MGET)
End Property

'Sub Functions
'Request
Private Function setMethod(Optional method As HttpMethod = MGET) As String
    Select Case method
    Case MGET: setMethod = "GET"
    Case MPOST: setMethod = "POST"
    Case MPATCH: setMethod = "PATCH"
    End Select
End Function

Private Function parseHeaders(headers As String) As Variant
    Dim outputArr As Variant
    Dim headersArr As Variant
    Dim i As Long
    If Len(baseHeaders) Then headers = baseHeaders & ";" & headers
    headersArr = Split(headers, ";")

    ReDim outputArr(0 To UBound(headersArr))
    For i = 0 To UBound(outputArr)
        outputArr(i) = Split(headersArr(i), ": ")
    Next
    parseHeaders = outputArr
End Function

Private Function setURL(ParamArray urlParts()) As String
    setURL = baseURL & "/vba_template/" & Join(urlParts, "/")
End Function

Private Function fetch(strUrl As String, Optional method As HttpMethod = MGET, Optional ReqHeaders As String = vbNullString, Optional Payload As Object = Nothing) As Object
    Dim hReq As Object
    Dim response As String
    Dim HeadersList As Variant
    Dim CurrHeader As Variant
    Dim payloadStr As String

    If method <> MGET And Not Payload Is Nothing Then
        payloadStr = JsonConverter.ConvertToJson(Payload)
    End If
    HeadersList = parseHeaders(ReqHeaders)
    On Error Resume Next
    Set hReq = GetObject(, "MSXML2.XMLHTTP")
    If hReq Is Nothing Then Set hReq = CreateObject("MSXML2.XMLHTTP")

    With hReq
        .Open setMethod(method), strUrl, False
        If Not HeadersList Is Nothing Then
            For Each CurrHeader In HeadersList
                .SetRequestHeader CurrHeader(0), CurrHeader(1)
            Next
        End If
        .Send payloadStr
        response = .ResponseText
    End With

    Set fetch = JsonConverter.ParseJson(response)
End Function

Private Function fetchAuthorized(strUrl As String, Optional method As HttpMethod = MGET, Optional ReqHeaders As String = vbNullString, Optional Payload As Object = Nothing) As Object
    Dim headers As String
    headers = "Authorization: Bearer " & GITHUB_TOKEN
    If Len(ReqHeaders) Then headers = headers & ";" & ReqHeaders
    Set fetchAuthorized = fetch(strUrl, method, headers, Payload)
End Function

'Commit and Committer
Private Sub SettingsSave(Optional branch As String = vbNullString, Optional oldSHA As String = vbNullString, Optional message As String = vbNullString, Optional Name As String = vbNullString, Optional email As String = vbNullString)
    Dim Re As New clsRE
    On Error GoTo BeforeExit
    With ShtOptionsProject.Range(BranchNameCell)
        If Len(branch & message & oldSHA) <> 0 Then
            Re.Init "^[A-Z0-9\_]+$", True, True
            'todo: запретить master
            If Not Re.TestString(branch) Then Err.Raise vbObjectError + 101, "Настройка Проект.Git/ClsGit.SettingsSave", "Ветка должна быть в одно слово," & vbNewLine & "англ буквами, цифрами или ""_"""
            If Len(oldSHA) = 0 Then Err.Raise vbObjectError + 102, "Настройка Проект.Git/ClsGit.SettingsSave", "ID пусто"
            If Len(message) = 0 Then Err.Raise vbObjectError + 103, "Настройка Проект.Git/ClsGit.SettingsSave", "Сообщение пусто"
            .Cells(1, 3).Value2 = branch
            .Cells(2, 3).Value2 = oldSHA
            .Cells(3, 3).Value2 = message
            With Head
                .branch = branch
                .SHA = oldSHA
                .message = message
            End With
        End If
        If Len(Name & email) <> 0 Then
            Re.Init RePattern_email, False, True
            If Len(Name) = 0 Then Err.Raise vbObjectError + 104, "Настройка Проект.Git/ClsGit.SettingsSave", "Имя пусто"
            If Not Re.TestString(email) Then Err.Raise vbObjectError + 105, "Настройка Проект.Git/ClsGit.SettingsSave", "Почта неверна"
            .Cells(4, 3).Value2 = Name
            .Cells(5, 3).Value2 = email
            With committer
                .Name = Name
                .email = email
            End With
        End If
    End With
BeforeExit:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Source
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub SettingsParse(Optional NeedMessage As Boolean = False)
    Dim values As Variant
    Dim Re As New clsRE
    On Error GoTo BeforeExit
    values = Array2DSlice(ShtOptionsProject.Range(BranchNameCell).Value2, ECol, 3)
    
    Re.Init "^[A-Z0-9\_]+$", True, True
    'todo: запретить master
    If Not Re.TestString(values(1)) Then Err.Raise vbObjectError + 101, "Настройка Проект.Git/ClsGit.SettingsParse", "Ветка должна быть в одно слово," & vbNewLine & "англ буквами, цифрами или ""_"""
    If Len(values(2)) = 0 Then Err.Raise vbObjectError + 102, "Настройка Проект.Git/ClsGit.SettingsParse", "ID пусто"
    If Len(values(3)) = 0 And NeedMessage Then Err.Raise vbObjectError + 103, "Настройка Проект.Git/ClsGit.SettingsParse", "Сообщение пусто"
    With Head
        .branch = values(1)
        .SHA = values(2)
        .message = values(3)
    End With
    Re.Init RePattern_email, False, True
    If Len(values(4)) = 0 Then Err.Raise vbObjectError + 104, "Настройка Проект.Git/ClsGit.SettingsSave", "Имя пусто"
    If Not Re.TestString(values(5)) Then Err.Raise vbObjectError + 105, "Настройка Проект.Git/ClsGit.SettingsSave", "Почта неверна"
    With committer
        .Name = values(4)
        .email = values(5)
    End With
BeforeExit:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Source
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function CommitCreate(message As String, treeSHA As String, parentSHA As String) As String
    Dim Payload As New Dictionary
    Dim response As Dictionary
    Dim parents As New Collection
    parents.Add parentSHA
    Payload.Add "message", message
    If committer.Name <> vbNullString Then Payload.Add "author", Author
    Payload.Add "tree", treeSHA
    Payload.Add "parents", parents
    Set response = fetchAuthorized(setURL("git", "commits"), MPOST, , Payload)
    Debug.Assert response.Exists("sha")
    Head.SHA = response("sha")
    '        If response.Exists("message") Then CommitCreate = response("message")
End Function

'Git submethods
'BRANCH
Public Function BranchGetList() As Collection
    Set BranchGetList = fetchAuthorized(setURL("git", "refs", "heads"), method:=MGET)
End Function

Public Function BranchGet() As Dictionary
    Dim branch As Dictionary
    
    Set branch = fetchAuthorized(setURL("commits", LCase(Head.branch)), method:=MGET)
    If branch.Exists("message") Then Exit Function
    Set BranchGet = New Dictionary
    BranchGet.Add "sha", branch("sha")
    BranchGet.Add "tree", branch("commit")("tree")("sha")
End Function

Private Function BranchCreate(Optional SHA = DefaultHeadSHA) As String
    Dim branch As Variant
    Dim Payload As New Dictionary
    Payload.Add "ref", "refs/heads/" & LCase(Head.branch)
    Payload.Add "sha", SHA
    
    Set branch = fetchAuthorized(setURL("git", "refs"), method:=MPOST, Payload:=Payload)
    Debug.Assert Not branch.Exists("message")
    BranchCreate = branch("object")("sha")
End Function

Private Function BranchSetRef(SHA As String) As String
    Dim branch As Variant
    Dim Payload As New Dictionary
    Payload.Add "sha", SHA
    Set branch = fetchAuthorized(setURL("git", "refs", "heads", Head.branch), method:=MPATCH, Payload:=Payload)
    If branch.Exists("message") And Not branch.Exists("object") Then BranchSetRef = branch("message")
    If branch.Exists("object") Then Head.SHA = branch("object")("sha")
End Function

'BLOB
Private Function BlobGet(ByVal url As String) As String
    Dim response As Object
    Set response = fetchAuthorized(url, MGET)
    If response.Exists("message") And Not response.Exists("content") Then Exit Function
    If InStr(response("encoding"), "base64") Then
        BlobGet = DecodeBase64(response("content"))
    Else
        BlobGet = response("content")
    End If
End Function

Private Function BlobCreate(ByVal content As String) As String
    Dim Payload As New Dictionary
    Dim response As Object
    Payload.Add "content", content
    Payload.Add "encoding", "utf-8"
    Set response = fetchAuthorized(setURL("git", "blobs"), MPOST, , Payload)
    Debug.Assert response.Exists("sha")
    BlobCreate = response("sha")
End Function

'CONTENT
Private Function ContentPush(ByVal Comp As Dictionary) As String
    Dim Payload As New Dictionary
    Dim response As Object
    Payload.Add "message", Head.message
    Payload.Add "committer", Author
    Payload.Add "content", Comp("content")
    Set response = fetchAuthorized(setURL("contents", Comp("path")), MPOST, , Payload)
    If response.Exists("sha") Then ContentPush = response("sha")
End Function

'TREE
Private Function TreeGet(ByVal treeSHA As String) As Dictionary
    Set TreeGet = New Dictionary
    Dim dir As String
    Dim path As String
    Dim curr As Variant
    Dim response As Dictionary
    Set response = fetchAuthorized(setURL("git", "trees", treeSHA & "?recursive=true"), MGET)
    If response.Exists("message") And Not response.Exists("tree") Then Exit Function
    On Error GoTo BeforeExit
    For Each curr In response("tree")
        If curr("mode") = "100644" And Not InStr(curr("path"), "README") Then
            TreeGet.Add curr("path"), BlobGet(curr("url"))
        End If
    Next
BeforeExit:
End Function

Private Function TreeCreate() As String
    Dim response As Dictionary
    Dim Payload As New Dictionary
    Dim content As String, currSHA As String
    Dim Comp As Variant
    Dim dict As Dictionary
    Dim comps As New Collection
    For Each Comp In CodeBase.Components.Items
        Set dict = New Dictionary
        If VBA.VarType(Comp("content")) = vbNull Then
            dict.Add "sha", Null
        Else
            content = Comp("content")
            currSHA = BlobCreate(content)
            If Len(currSHA) = 0 Then GoTo Continue
            dict.Add "sha", currSHA
        End If
        dict.Add "path", Comp("path")
        dict.Add "mode", Comp("mode")
        dict.Add "type", Comp("type")
        comps.Add dict
Continue:
    Next
    Payload.Add "base_tree", Head.treeSHA
    Payload.Add "tree", comps
    Set response = fetchAuthorized(setURL("git", "trees"), MPOST, , Payload)
    Head.treeSHA = response("sha")
    If response.Exists("message") And Not response.Exists("sha") Then TreeCreate = response("message")
End Function

'Git methods
Public Function Push() As Boolean
    On Error GoTo BeforeExit
    Dim state As String
    Dim CurrHead As Dictionary
    Set CodeBase = New clsCode
    SettingsParse True
    Set CurrHead = BranchGet
    If Not CurrHead.Exists("tree") Then Err.Raise vbObjectError + 106, "Git Push/ClsGit.BranchGet", "Нет Tree SHA"
    Head.SHA = CurrHead("sha")
    Head.treeSHA = CurrHead("tree")
    CodeBase.Compare TreeGet(Head.treeSHA)
    'Получить ветку
    state = TreeCreate
    If Len(state) <> 0 Then Err.Raise vbObjectError + 106, "Git Push/ClsGit.TreeCreate", state
    'Подготовить коммит
    state = CommitCreate(Head.message, Head.treeSHA, Head.SHA)
    If Len(state) <> 0 Then Err.Raise vbObjectError + 106, "Git Push/ClsGit.CommitCreate", state
    'Обновить состояние к новому
    state = BranchSetRef(Head.SHA)
    If Len(state) <> 0 Then Err.Raise vbObjectError + 106, "Git Push/ClsGit.BranchSetRef", state
    SettingsSave
    Push = True
BeforeExit:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Source
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function Pull(Optional path As String = vbNullString) As Boolean
    Dim response As Object
    Dim headers As Variant
    Set CodeBase = New clsCode
    Dim CurrHead As Dictionary
    Set CurrHead = BranchGet
    If Not CurrHead.Exists("tree") Then Err.Raise vbObjectError + 106, "Git Pull/ClsGit.BranchGet", "Нет Tree SHA"
    Head.treeSHA = CurrHead("tree")
    'Получить ветку
    Set response = TreeGet(Head.treeSHA)
    If response.count = 0 Then Err.Raise vbObjectError + 106, "Git Pull/ClsGit.TreeGet", "Нет файлов"
    CodeBase.Compare response
    'Получить файлы
End Function

Private Sub Class_Initialize()
    SettingsParse
End Sub
