'@Folder("VBAProject")
Option Explicit

Private elements As New Collection
Private Const PageList As String = "Главная;Настройки;Проект"
Private Const PageMain As String = "PageMainMenu"
Private Const PageSettings As String = "PageSettings"
Private Const PageProject As String = "PageProject"

Private Enum NumPage
    pMain = 0
    pSettings = 1
    pProject = 2
End Enum

'USERFORM
Private Sub UserForm_Initialize()
    CtrlDefaultParams Me, 330, 480
    With Me
        CtrlDefaultParams .MultiPages, 330, 405, 48
        With .MultiPages
            'Стартовая страница
            .Value = NumPage.pMain

        End With
        
    End With
End Sub

'GLOBAL
Private Sub MultiPages_Layout(ByVal Index As Long)
    SetPageTitle Index
    RenderContent Index
End Sub

'TOOLBAR
Private Sub Button_Main_Click()
    Me.MultiPages.Value = NumPage.pMain
End Sub

Private Sub Button_Settings_Click()
    Me.MultiPages.Value = NumPage.pSettings
End Sub

Private Sub Button_Project_Click()
    Me.MultiPages.Value = NumPage.pProject
End Sub

'MAINPAGE

Private Sub Button_TemplateAdd_Click()
    AddTemplate
End Sub

'SETTINGS
Private Sub Button_Dev_Pull_Click()
    Dim Re As New clsRE
    Dim Git As New ClsGit
    Dim Name As String
    Dim email As String
    Dim password As String
    Re.init RePattern_email, False, True
    With Me
        If isControlEmpty(.TextBox_DevName) Or isControlEmpty(.TextBox_DevPassword) Or .TextBox_DevPassword.Value <> DEV_PASSWORD Or isControlEmpty(.TextBox_DevEmail) Or Not Re.TestString(.TextBox_DevEmail.Value) Then GoTo BeforeExit
        Name = .TextBox_DevName.Value
        email = .TextBox_DevEmail.Value
        password = .TextBox_DevPassword.Value
    End With
    'Git pull
    Exit Sub
BeforeExit:
    MsgBox "Данные введены неверно!", vbOKOnly, "Имя, email и пароль"
End Sub

'SUBFUNCS
Private Sub AddTemplate()
    Me.Hide
    ufTemplate.ReadTemplates
    Me.Show
End Sub

Private Sub SetPageTitle(ByVal Index As Long)
    Dim pList As Variant
    pList = Split(PageList, ";")
    With Me.Toolbar
        .Controls("Label_PageTitle").Caption = pList(Index)
    End With
End Sub

Private Sub DrawTemplates(Page As MSForms.Page)
    Dim Git As New ClsGit
    Dim Settings As Dictionary
    Dim ctrl As clsModArg
    Dim i As Long
    Dim keys As Variant
    Set Settings = Git.Settings
    For i = 0 To Settings.count - 1
        keys = Settings.keys
        Set ctrl = New clsModArg
        ctrl.Add Page.Controls("BG_Settings"), keys(i), keys(i), "String", True, i, 0, , , , 50
        ctrl.Value = Settings(keys(i))
        elements.Add ctrl
    Next

End Sub

Private Sub DrawSettings(Page As MSForms.Page)
    Dim Git As New ClsGit
    Dim Settings As Dictionary
    Dim ctrl As clsModArg
    Dim i As Long
    Dim keys As Variant
    Set Settings = Git.Settings
    For i = 0 To Settings.count - 1
        keys = Settings.keys
        Set ctrl = New clsModArg
        ctrl.Add Page.Controls("BG_Settings"), keys(i), keys(i), "String", True, i, 0, , , , 50
        ctrl.Value = Settings(keys(i))
        elements.Add ctrl
    Next

End Sub
Private Sub RenderContent(ByVal Index As Long)
    With Me.MultiPages
        Select Case Index
        Case NumPage.pMain
            Debug.Print ""
            '        .Pages (PageMain)
        Case NumPage.pProject
            Debug.Print ""
            '        .Pages (PageProject)
        Case NumPage.pSettings
            DrawSettings .Pages(PageSettings)
    
        End Select
    End With
End Sub
