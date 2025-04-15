'@Folder("UFComponents")
Option Explicit

Private Page As Variant
Private WithEvents ctrl As MSForms.TextBox
Private WithEvents checker As MSForms.CheckBox
Private WithEvents dropdown As MSForms.ComboBox
Private modifiersList As Dictionary
Private m_IsOptional As Boolean
Private m_ArgType As String
Private m_ArgPlaceholder As String

Public Property Get ArgType() As String
    ArgType = m_ArgType
End Property
Public Property Get Value() As String
    Select Case True
    Case m_ArgType = "Boolean"
        Value = checker.Value
    Case m_ArgType = "FuncName"
        Value = dropdown.Value
    Case Else
        Value = ctrl.text
    End Select
End Property
Public Property Let Value(v As String)
    Select Case True
    Case m_ArgType = "Boolean"
        checker.Value = v
    Case m_ArgType = "FuncName"
        dropdown.Value = v
    Case Else
        ctrl.text = v
    End Select
End Property

Public Sub Add(Wrap As Variant, ByVal Func As String, ByVal ArgName As String, ByVal arg As String, _
ByVal IsOptional As Boolean, Index As Long, ColIndex As Long, Optional modifiers As Dictionary, _
Optional Width As Long = 60, Optional Height As Long = 20, Optional loffset As Long = 220)
    If Not Len(ArgName) > 0 Then Exit Sub
    Set Page = Wrap
    m_ArgType = arg
    m_IsOptional = IsOptional
    m_ArgPlaceholder = ArgName
    If m_ArgType = "FuncName" And Not modifiers Is Nothing Then Set modifiersList = modifiers
    With Page
    Select Case m_ArgType
        Case "Boolean"
            Set checker = .Controls.Add("Forms.CheckBox.1", "CheckBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
            CtrlDefaultParams checker, Width, Height, Index * 30 + 10, loffset + (ColIndex * 70), m_ArgPlaceholder, Index
            checker.TextAlign = fmTextAlignLeft
        Case "FuncName"
            Set dropdown = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
            dropdown.List = modifiersList.Keys
            CtrlDefaultParams dropdown, Width, Height, Index * 30 + 10, loffset + (ColIndex * 70), , Index
            dropdown.TextAlign = fmTextAlignLeft
        Case "Variant"
            Set ctrl = AddTextBox(Page, Join(Array("Arg", Index, Func, ColIndex), "_"), Index, Width, Height, loffset + (ColIndex * 70), m_ArgPlaceholder, True)
        Case Else
            Set ctrl = AddTextBox(Page, Join(Array("Arg", Index, Func, ColIndex), "_"), Index, Width, Height, loffset + (ColIndex * 70), m_ArgPlaceholder, False)
        End Select
    End With

End Sub

Public Function Validate() As Boolean
    Dim s As MSForms.CheckBox
    Select Case True
    Case m_ArgType = "Long"
        Validate = IsNumeric(ctrl.text)
    Case m_ArgType = "String"
        Validate = Len(ctrl.text)
    Case m_ArgType = "FuncName"
        Validate = modifiersList.Exists = dropdown.Value
    Case Else
        Validate = True
    End Select
    If m_IsOptional And Not Validate Then Validate = (Len(Value) = 0)
End Function

Public Sub Clear()
    With Page
        .Controls.Remove ctrl.Name
    End With
End Sub

Private Sub ctrl_Change()
ctrl.BackStyle = IIf(Len(ctrl.text) = 0, fmBackStyleTransparent, fmBackStyleOpaque)
End Sub

Private Sub ctrl_DropButtonClick()
Dim formula As Variant
Page.Parent.Parent.Hide
    On Error GoTo NoValue
    Set formula = Application.InputBox("Введите параметры через "",""", "Список параметров", ctrl.text, Type:=64)
    ctrl.Value = Join(formula, ", ")
NoValue:
    Page.Parent.Parent.Show vbModal
    On Error GoTo 0
End Sub

Private Sub dropdown_Change()
    Debug.Print "dropdown_Change: "; dropdown.Value
End Sub

Private Sub checker_Change()
    Debug.Print "checker_Change: "; checker.Value
End Sub
