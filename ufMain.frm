'@Folder("VBAProject")
Option Explicit

Private Sub Button_Main1_Click()
    Me.MultiPageMenu.value = 0
End Sub

Private Sub Button_Main2_Click()
    Me.MultiPageMenu.value = 0
End Sub

Private Sub Button_Settings_Click()
    Me.MultiPageMenu.value = 1
End Sub

Private Sub Button_Template_Click()
    Me.MultiPageMenu.value = 2
End Sub


Private Sub AddTemplate()

End Sub

Private Sub Button_Dev_Pull_Click()
    Dim Re As New clsRE
    Dim git As New ClsGit
    Dim name As String
    Dim email As String
    Dim password As String
    Re.Init RePattern_email, False, True
    With Me
        If isControlEmpty(.TextBox_DevName) Or isControlEmpty(.TextBox_DevPassword) Or .TextBox_DevPassword.value <> DEV_PASSWORD Or isControlEmpty(.TextBox_DevEmail) Or Not Re.TestString(.TextBox_DevEmail.value) Then GoTo BeforeExit
        name = .TextBox_DevName.value
        email = .TextBox_DevEmail.value
        password = .TextBox_DevPassword.value
    End With
    'Git pull
    Exit Sub
BeforeExit:
    MsgBox "Данные введены неверно!", vbOKOnly, "Имя, email и пароль"
End Sub

Private Sub UserForm_Initialize()
    
    With Me
        .KeepScrollBarsVisible = fmScrollBarsNone
        .Height = 480
        .Width = 320
        .BackColor = RGB(256, 256, 256)
        .ForeColor = FOREGRAY
        .StartUpPosition = 1
    End With
End Sub
