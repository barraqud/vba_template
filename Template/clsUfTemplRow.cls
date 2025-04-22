'@Folder("Template")
Option Explicit

Public Title As String
Public Parent As String
Public index As Long
Public InitColumn As Long
Public Func As String

Private args As Dictionary
Private Page As MSForms.Page
Private TitleLabel As MSForms.Label
Private WithEvents InitData As MSForms.TextBox
Private WithEvents Modifier As MSForms.ComboBox
Private modifiersList As Dictionary
Private modArgs As Collection
Private ColIndex As Long
Private Top As Long
Private Const Width As Long = 60
Private Const Height As Long = 20

Public Sub Init(CurrPage As MSForms.Page, ByVal RowTitle As String, ByVal RowParent As String, ByVal RowIndex As Long)
    index = RowIndex
    Title = RowTitle
    Parent = RowParent
    Set Page = CurrPage
    Set modifiersList = ModList
    AddLabel
    Set InitData = AddTextBox(Page, "InitData_" & Parent, index, Width, Height, , "Номер колонки", True)
    AddModifier
End Sub

Private Sub AddLabel(Optional Left As Long = 10)
    With Page
        Set TitleLabel = .Controls.Add("Forms.Label.1", "Label_" & Parent)
        CtrlDefaultParams TitleLabel, Width, Height, index * 30 + 10, Left, Title, index
        TitleLabel.ForeColor = FOREPURPLE
    End With
End Sub

Private Sub AddModifier(Optional Left As Long = 150)
    With Page
        Set Modifier = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Parent)
        Modifier.List = modifiersList.Keys
        CtrlDefaultParams Modifier, Width, Height, index * 30 + 10, Left, , index
        Modifier.TextAlign = fmTextAlignLeft
    End With
End Sub

Public Function ParseModArg() As Dictionary
    Set ParseModArg = New Dictionary
    Dim params As Variant
    Dim ctrl As Variant
    With Page
        For Each ctrl In .Controls
            If InStr(ctrl.Name, "Arg_" & index) = 1 Or InStr(ctrl.Name, "ComboBox_Arg_" & index) = 1 Then
                ParseModArg.Add Split(ctrl.Name, "_")(3), ctrl.text
            End If
        Next
    End With
End Function

Private Sub clearModArgs()
    Dim ctrl As Variant
    With Page
        For Each ctrl In .Controls
            If InStr(ctrl.Name, "Arg_" & index) Then
                .Controls.Remove ctrl.Name
            End If
        Next
    End With
    ColIndex = 0
End Sub

Private Sub CheckInitData()
    Dim RE As New clsRE
    RE.Init RePattern_rngstring, False, True
    On Error GoTo BeforeExit
    Select Case True
    Case RE.TestString(InitData.Value)
        InitColumn = Range(InitData.Value).Column
    Case InitData.Value = vbNullString: Exit Sub
    Case Else
        InitColumn = CInt(InitData.Value)
    End Select
    Exit Sub
BeforeExit:
    MsgBox "Введите номер колонки (Исходная информация) или выберите диапазон колонок"
End Sub

Private Sub InitData_Change()
    InitData.BackStyle = IIf(Len(InitData.text) = 0, fmBackStyleTransparent, fmBackStyleOpaque)
End Sub

Private Sub InitData_DropButtonClick()
    Dim rng As Range
    Page.Parent.Parent.Hide
    ShtMainData.Activate
    On Error GoTo NoValue
    Set rng = Application.InputBox("Выберите колонку", "Указатель номера колонки", InitData.text, Type:=8)
    InitData.Value = "'" & rng.Parent.Name & "'!" & rng.Address(False, False)
NoValue:
    Page.Parent.Parent.Show vbModal
    On Error GoTo 0
End Sub

Private Sub Modifier_Change()
    Dim key As Variant
    Dim FnData As Dictionary
    Dim modArg As clsModArg
    
    CheckInitData
    clearModArgs
    Set modArgs = New Collection
    If Not modifiersList.Exists(Modifier.Value) Then Exit Sub
    Set FnData = modifiersList(Modifier.Value)
    Func = FnData("Name")
    Set args = FnData("Args")
    On Error GoTo BeforeNext
    'Argument title
    For Each key In args.Keys
        Set modArg = New clsModArg
        modArg.Add Page, Func, key, args(key)("Type"), args(key)("Optional"), index, ColIndex, modifiersList, Width, Height
        modArgs.Add modArg
BeforeNext:
        ColIndex = ColIndex + 1
    Next
    If Page.Parent.Parent.Width < 220 + 70 * ColIndex Then
        Page.Parent.Parent.Width = 220 + 70 * ColIndex + 20
        Page.Parent.Width = 220 + 70 * ColIndex + 20
    End If
    On Error GoTo 0
End Sub

