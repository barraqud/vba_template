'@Folder("Template")
Option Explicit

Public UF As UserForm
Public wrap As Variant
Public Name As String
Public IdxHorizontal As Long
Public IdxVertical As Long

Private Self As MSForms.Frame
Private WithEvents TextBlock As MSForms.textBox
Private WithEvents CheckBlock As MSForms.CheckBox
Private WithEvents SelectBlock As MSForms.ComboBox

Private m_loc As TLoc
Private m_BlockType As String
Private m_SelectList As Dictionary
Private m_IsOptional As Boolean
Private m_ValueType As String
Private m_Placeholder As String

Public Property Get blockType() As String
    ValType = m_ArgType
End Property

Public Property Get ValType() As String
    ValType = m_ArgType
End Property

Public Property Get Top() As String
    Top = Self.Top
End Property

Public Property Let Top(Value As Long)
    Self.Top = Value
End Property

Public Property Get Left() As String
    Left = Self.Left
End Property

Public Property Let Left(Value As Long)
    Self.Left = Value
End Property

Public Property Get Width() As String
    Width = Self.Width
End Property

Public Property Let Width(Value As Long)
    Self.Width = Value
End Property

Public Property Get Height() As String
    Height = Self.Height
End Property

Public Property Let Height(Value As Long)
    Self.Height = Value
End Property

Private Sub addType(blockType As String)
    Select Case blockType
    Case "Boolean"
        m_BlockType = "CheckBox"
        Set CheckBlock = Self
    Case "FuncName"
        m_BlockType = "ComboBox"
        Set SelectBlock = Self
    Case Else
        m_BlockType = "TextBox"
        Set TextBlock = Self
    End Select
End Sub

Public Sub Add(UForm As UserForm, SubName As String, Optional helper As String, _
               Optional Vertical As Long = 0, Optional Horizontal As Long = 0, _
               Optional blockType As String, Optional Parent As MSForms.Control, _
               Optional TopOffset As Long = 10, Optional LeftOffset As Long = 10, _
               Optional Width As Long = 60, Optional Height As Long = 20)
    Set UF = UForm
    Name = SubName
    IdxVertical = Vertical
    IdxHorizontal = Horizontal
    m_Placeholder = helper
    If IsMissing(Parent) Then
        Set wrap = UF
    Else
        Set wrap = Parent
    End If
    Set Frame = ControlCreate(wrap, "Frame", Width, Height * 2, Vertical, Horizontal, TopOffset, LeftOffset)
    With wrap.Controls

        Set Self = ControlCreate(Frame, m_BlockType, Width, Height, Vertical, Horizontal, TopOffset + Height, LeftOffset, m_Placeholder)
    End With
    
End Sub

'
'Public Sub Add(Wrap As Variant, ByVal Func As String, ByVal ArgName As String, ByVal arg As String, ByVal IsOptional As Boolean, Index As Long, ColIndex As Long, Optional modifiers As Dictionary, Optional Width As Long = 60, Optional Height As Long = 20)
'    If Not Len(ArgName) > 0 Then Exit Sub
'    Set Page = Wrap
'    m_ArgType = arg
'    m_IsOptional = IsOptional
'    m_ArgPlaceholder = ArgName
'    If m_ArgType = "FuncName" And Not modifiers Is Nothing Then Set modifiersList = modifiers
'    With Page
'    Select Case m_ArgType
'        Case "Boolean"
'            Set checker = .Controls.Add("Forms.CheckBox.1", "CheckBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
'            CtrlDefaultParams checker, Width, Height, Index * 30 + 10, 220 + 70 * ColIndex, m_ArgPlaceholder, Index
'            checker.TextAlign = fmTextAlignLeft
'        Case "FuncName"
'            Set dropdown = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
'            dropdown.List = modifiersList.keys
'            CtrlDefaultParams dropdown, Width, Height, Index * 30 + 10, 220 + 70 * ColIndex, , Index
'            dropdown.TextAlign = fmTextAlignLeft
'        Case "Variant"
'            Set ctrl = AddTextBox(Page, Join(Array("Arg", Index, Func, ColIndex), "_"), Index, Width, Height, 220 + 70 * ColIndex, m_ArgPlaceholder, True)
'        Case Else
'            Set ctrl = AddTextBox(Page, Join(Array("Arg", Index, Func, ColIndex), "_"), Index, Width, Height, 220 + 70 * ColIndex, m_ArgPlaceholder, False)
'        End Select
'    End With
'
'End Sub

Public Function Validate() As Boolean
    Dim s As MSForms.CheckBox
    Select Case True
    Case m_ArgType = "Long"
        Validate = IsNumeric(ctrl.text)
    Case m_ArgType = "String"
        Validate = Len(ctrl.text)
    Case "FuncName"
        Validate = modifiersList.Exists = dropdown.Value
    Case Else
        Validate = True
    End Select
    If m_opt And Not Validate Then Validate = (Len(Value) = 0)
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
    Debug.Print dropdown.Value
End Sub

Private Sub checker_Change()
    Debug.Print checker.Value
End Sub

