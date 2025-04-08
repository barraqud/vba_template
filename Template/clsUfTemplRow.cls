'@Folder("Template")
Option Explicit

Public Title As String
Public Parent As String
Public Index As Long
Public InitColumn As Long
Public Func As String

Private Args As Dictionary
Private Page As MSForms.Page
Private TitleLabel As MSForms.label
Private WithEvents InitData As MSForms.textbox
Private WithEvents Modifier As MSForms.ComboBox
Private modifiersList As Dictionary
Private ColIndex As Long
Private Top As Long
Private Const Width As Long = 60
Private Const Height As Long = 20

Public Sub Init(CurrPage As MSForms.Page, ByVal RowTitle As String, ByVal RowParent As String, ByVal RowIndex As Long)
    Index = RowIndex
    Title = RowTitle
    Parent = RowParent
    Set Page = CurrPage
    Set modifiersList = ModList
    AddLabel
    Set InitData = AddTextBox("InitData_" & Parent, , "Номер колонки", True)
    AddModifier
End Sub

Private Sub AddLabel(Optional Left As Long = 10)
    With Page
        Set TitleLabel = .Controls.Add("Forms.Label.1", "Label_" & Parent)
        CtrlDefaultParams TitleLabel, Width, Height, Index * 30 + 10, Left, Title, Index
        TitleLabel.ForeColor = FOREPURPLE
    End With
End Sub

Private Function AddTextBox(name As String, Optional Left As Long = 80, Optional helper As String, Optional isDropDown As Boolean = False) As MSForms.textbox
    Dim Placeholder As MSForms.textbox
    Dim Box As MSForms.textbox
    With Page
        Set Placeholder = .Controls.Add("Forms.TextBox.1", "Placeholder_" & name)
        CtrlDefaultParams Placeholder, Width, Height, Index * 30 + 10, Left, , Index
        Placeholder.text = helper
        Placeholder.MultiLine = True
        Placeholder.TextAlign = fmTextAlignLeft
        Placeholder.Enabled = False
        
        Set Box = .Controls.Add("Forms.TextBox.1", name)
        CtrlDefaultParams Box, Width, Height, Index * 30 + 10, Left, , Index
        If isDropDown Then
            Box.DropButtonStyle = fmDropButtonStyleReduce
            Box.ShowDropButtonWhen = fmShowDropButtonWhenAlways
        End If
        Box.BackStyle = fmBackStyleTransparent
    End With
    Set AddTextBox = Box
End Function

Private Sub AddModifier(Optional Left As Long = 150)
    With Page
        Set Modifier = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Parent)
        Modifier.List = modifiersList.Keys
        CtrlDefaultParams Modifier, Width, Height, Index * 30 + 10, Left, , Index
        Modifier.TextAlign = fmTextAlignLeft
    End With
End Sub

Public Function ParseModArg() As Dictionary
    Set ParseModArg = New Dictionary
    Dim params As Variant
    Dim ctrl As Variant
    With Page
        For Each ctrl In .Controls
            If InStr(ctrl.name, "Arg_" & Index) = 1 Or InStr(ctrl.name, "ComboBox_Arg_" & Index) = 1 Then
                ParseModArg.Add Split(ctrl.name, "_")(3), ctrl.text
            End If
        Next
    End With
End Function

Private Sub AddModArg(ByVal arg As String, ByVal IsOptional As Boolean)
    Dim ctrl As Control
    Dim helper As String
    
    Select Case arg
    Case "Long"
        helper = "Число"
    Case "String"
        helper = "Текст"
    Case "Variant"
        helper = "Текст"
    Case "Boolean"
        helper = "Текст"
    Case "FuncName"
        With Page
            Set ctrl = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
            ctrl.List = modifiersList.Keys
            CtrlDefaultParams ctrl, Width, Height, Index * 30 + 10, 220 + 70 * ColIndex, , Index
            ctrl.TextAlign = fmTextAlignLeft
        End With
    Case Else
        Debug.Print "else "; arg
    End Select
    If Len(helper) > 0 Then Set ctrl = AddTextBox(Join(Array("Arg", Index, Func, ColIndex), "_"), 220 + 70 * ColIndex, helper, False)
End Sub

Private Sub clearModArgs()
    Dim ctrl As Variant
    With Page
        For Each ctrl In .Controls
            If InStr(ctrl.name, "Arg_" & Index) Then
                .Controls.Remove ctrl.name
                ColIndex = 0
            End If
        Next
    End With
End Sub


Private Sub CheckInitData()
    Dim Re As New clsRE
    Re.Init RePattern_rngstring, False, True
    On Error GoTo BeforeExit
    Select Case True
    Case Re.TestString(InitData.value)
        InitColumn = Range(InitData.value).Column
    Case InitData.value = vbNullString: Exit Sub
    Case Else
        InitColumn = CInt(InitData.value)
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
    Set rng = Application.InputBox("Select the range", "Range Picker", InitData.text, Type:=8)
    InitData.value = "'" & rng.Parent.name & "'!" & rng.Address(False, False)
    Page.Parent.Parent.Show vbModal
    Exit Sub
NoValue:
    Page.Parent.Parent.Show vbModal
    On Error GoTo 0
End Sub

Private Sub Modifier_Change()
    Dim key As Variant
    Dim FnData As Dictionary
    
    CheckInitData
    clearModArgs
    If Not modifiersList.Exists(Modifier.value) Then Exit Sub
    Set FnData = modifiersList(Modifier.value)
    Func = FnData("Name")
    Set Args = FnData("Args")
    On Error GoTo BeforeNext
    'Argument title
    For Each key In Args.Keys
        AddModArg Args(key)("Type"), Args(key)("Optional")
BeforeNext:
        ColIndex = ColIndex + 1
    Next
    If Page.Parent.Parent.Width < 220 + 70 * ColIndex Then
    Page.Parent.Parent.Width = 220 + 70 * ColIndex + 20
    Page.Parent.Width = 220 + 70 * ColIndex + 20
    End If
    On Error GoTo 0
End Sub
