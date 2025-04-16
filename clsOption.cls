'@Folder("VBAProject")
Option Explicit

Private Enum NumPage
    pMain = 0
    pSettings = 1
    pProject = 2
End Enum

Private Const PageNames As String = "MainMenu;Settings;Project"
Private Const PageTitles As String = "Главная;Настройки;Проект"
Const Prefix_Git As String = "TextBox_Git_"

Private CurrentLayout As Long
Private elements As New Collection

Private Property Get Title() As Long
    Dim pageIndex As Long
    Dim pList As Variant
    pList = Split(PageTitles, ";")
    With Me.Toolbar
        For pageIndex = 0 To UBound(pList)
            If .Controls("Label_PageTitle").Caption = pList(pageIndex) Then Title = pageIndex
        Next
    End With
End Property

Private Property Let Title(pageIndex As Long)
    Dim pNames As Variant
    Dim pList As Variant
    pNames = Split(PageNames, ";")
    pList = Split(PageTitles, ";")
    With Me.Toolbar
        .Controls("Label_PageTitle").Caption = pList(pageIndex)
        .Controls("Button_" & pNames(pageIndex)).Visible = False
        If CurrentLayout > 0 Then .Controls("Button_" & pNames(CurrentLayout - 1)).Visible = True
    End With
End Property

Private Property Get Status() As String
    Status = Me.Toolbar.Controls("Label_Status").Caption
End Property

Private Property Let Status(s As String)
    With Me.Toolbar
        .Controls("Label_Status").Caption = s
    End With
End Property

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
    ToolbarUpdate Index
    With Me.MultiPages
        Select Case Index
        Case NumPage.pMain
            MainTemplatesList
        Case NumPage.pProject
            ProjectInit
        Case NumPage.pSettings
            SettingsDetails
        End Select
    End With
End Sub

'TOOLBAR
Private Sub Button_MainMenu_Click()
    Me.MultiPages.Value = NumPage.pMain
End Sub

Private Sub Button_Settings_Click()
    Me.MultiPages.Value = NumPage.pSettings
End Sub

Private Sub Button_Project_Click()
    Me.MultiPages.Value = NumPage.pProject
End Sub

Private Sub ToolbarUpdate(ByVal Idx As Long, Optional ByVal statusText As String = vbNullString)
    Title = Idx
    Status = statusText
    'Обновляет текущую страницу
    CurrentLayout = Idx + 1
End Sub

'MAINPAGE
Private Sub MainTemplatesList()
    Dim Saved As Dictionary
    Dim templ As New clsTemplate
    Set Saved = templ.ParseSaved
End Sub

Private Sub Button_TemplateAdd_Click()
    AddTemplate
End Sub

'PROJECT
Public Sub ProjectInit(Optional ByVal ProjectName As String)
    
End Sub

'SETTINGS
Private Sub SettingsDetails()
    Dim Git As New clsGit
    On Error Resume Next
    With Me.MultiPages.Pages(NumPage.pSettings)
        .Controls(Prefix_Git & "Branch").Value = Git.Settings("Ветка")
        .Controls(Prefix_Git & "Message").Value = Git.Settings("Описание закрепления")
        .Controls(Prefix_Git & "SHA").Value = Git.Settings("ID закрепления")
        .Controls(Prefix_Git & "AuthorName").Value = Git.Settings("Имя автора")
        .Controls(Prefix_Git & "AuthorEmail").Value = Git.Settings("Почта автора")
    End With
End Sub

Private Sub Button_Git_Push_Click()
    Dim Git As New clsGit
    With Me.MultiPages.Pages(NumPage.pSettings).Controls(Prefix_Git & "Message")
        If Len(.Value) <> 0 Then
            Git.Push .Value
        Else
            MsgBox "Нужно указать описание", vbOKOnly, "Описание закрепления"
        End If
    End With
    Exit Sub
BeforeExit:
    MsgBox "Данные введены неверно!", vbOKOnly, "Имя, email и пароль"
End Sub

Private Sub Button_Git_Refresh_Click()
    Dim Git As New clsGit
    Dim done As Boolean
    done = Git.Pull
    If done Then Status = "Код обновлен"
End Sub

Private Sub TextBox_Git_Access_Change()
    If TextBox_Git_Access.Value <> GITHUB_SETTINGS_ACCESS Then Exit Sub
    On Error Resume Next
    With Me.MultiPages.Pages(NumPage.pSettings)
        .Controls(Prefix_Git & "Branch").Locked = False
        With .Controls(Prefix_Git & "Message")
            .Locked = False
            .Value = vbNullString
        End With
        .Controls(Prefix_Git & "SHA").Locked = False
        .Controls(Prefix_Git & "AuthorName").Locked = False
        .Controls(Prefix_Git & "AuthorEmail").Locked = False
        With TextBox_Git_Access
            .Locked = True
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbGreen
            .ForeColor = FOREGRAY
        End With
    End With
    Button_Git_Push.Enabled = True
    Status = "Можно изменять"
End Sub

Private Sub TextBox_Git_RefreshAccess_Change()
    If TextBox_Git_RefreshAccess.Value <> GITHUB_REFRESH_ACCESS Then Exit Sub
    On Error Resume Next
    With TextBox_Git_RefreshAccess
        .Locked = True
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = vbGreen
        .ForeColor = FOREGRAY
    End With
    Button_Git_Refresh.Enabled = True
    Status = "Можно обновлять"
End Sub

'SUBFUNCS
Private Sub AddTemplate()
    Me.Hide
    ufTemplate.ReadTemplates
    Me.Show
End Sub

Private Sub ProjectSetParameters()
    Dim templ As New clsTemplate
End Sub
