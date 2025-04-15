'@Folder("UFComponents")
Option Explicit

Public Parent As Variant
Public Name As String
Public Row As Long
Public Column As Long
Public IsOrientVert As Boolean

Private Self As MSForms.Control
Private WithEvents LabelBlock As MSForms.Label
Private WithEvents TextBlock As MSForms.TextBox
Private WithEvents CheckBlock As MSForms.CheckBox
Private WithEvents SelectBlock As MSForms.ComboBox

Public Tip As Boolean
Private m_TipText As String
Private m_BlockType As String
Private m_SelectList As Dictionary
Private m_IsOptional As Boolean
Private m_ArgType As String
Private m_Placeholder As String

Public Property Get Value() As String
    Value = Self.Value
End Property

Public Property Let Value(v As String)
    Self.Value = v
End Property

Public Property Get blockType() As String
    blockType = m_ArgType
End Property

Public Property Get Top() As Long
    Top = Self.Top
End Property

Public Property Let Top(v As Long)
    Self.Top = v
End Property

Public Property Get Left() As Long
    Left = Self.Left
End Property

Public Property Let Left(v As Long)
    Self.Left = v
End Property

Public Property Get Width() As Long
    Width = Self.Width
End Property

Public Property Let Width(v As Long)
    Self.Width = v
End Property

Public Property Get Height() As Long
    Height = Self.Height
End Property

Public Property Let Height(v As Long)
    Self.Height = v
End Property

Public Sub Add(ParentBlock As Variant, ByVal SubName As String, Optional InitValue As String, _
               Optional ByVal helper As String, Optional ByVal IsOrientVertically As Boolean = True, _
               Optional ByVal IdxVertical As Long = 0, Optional ByVal IdxHorizontal As Long = 0, _
               Optional ByVal blockType As String, Optional ByVal isOpt As Boolean, _
               Optional ByVal TopOffset As Long = 10, Optional ByVal LeftOffset As Long = 10, _
               Optional ByVal Width As Long = 60, Optional ByVal Height As Long = 20, _
               Optional ValList As Dictionary, Optional TipText As String)
    Set Parent = ParentBlock
    
    Name = SubName
    m_ArgType = blockType
    Row = IdxVertical
    Column = IdxHorizontal
    m_Placeholder = helper
    m_IsOptional = isOpt
    IsOrientVert = IsOrientVertically
    If Not ValList Is Nothing Then Set m_SelectList = ValList
    If TipText Then m_TipText = TipText
    
    On Error GoTo BeforeExit
    Set LabelBlock = CreateLabel(Name, Width, Height, TopOffset, LeftOffset)
    CreateInput
    Self.Value = InitValue
BeforeExit:

End Sub

Public Function Validate() As Boolean
    Dim v As String
    Dim s As MSForms.CheckBox
    Select Case True
    Case m_ArgType = "Long"
        v = TextBlock.text
        Validate = IsNumeric(v)
    Case m_ArgType = "String"
        v = TextBlock.text
        Validate = Len(v)
    Case "FuncName"
        v = SelectBlock.Value
        Validate = m_SelectList.Exists(v)
    Case Else
        v = TextBlock.text
        Validate = True
    End Select
    If m_IsOptional And Not Validate Then Validate = (Len(v) = 0)
End Function

Public Sub Clear()
    Parent.Controls.Remove LabelBlock.Name
    Parent.Controls.Remove Self.Name
End Sub

Private Sub LabelBlock_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Tip Then Exit Sub
    Debug.Print Button, Shift
End Sub

Private Sub TextBlock_Change()
    TextBlock.BackStyle = IIf(Len(TextBlock.text) = 0, fmBackStyleTransparent, fmBackStyleOpaque)
End Sub

Private Sub TextBlock_DropButtonClick()
    Dim UF As UserForm
    If TypeOf Parent Is UserForm Then
        UF = Parent
    Else
        UF = Parent.Parent
    End If
    Dim formula As Variant
    UF.Hide
    On Error GoTo NoValue
    Set formula = Application.InputBox("Введите параметры через "",""", "Список параметров", TextBlock.text, Type:=64)
    TextBlock.Value = Join(formula, ", ")
NoValue:
    UF.Show vbModal
    On Error GoTo 0
End Sub

Private Sub SelectBlock_Change()
    Debug.Print SelectBlock.Value
End Sub

Private Sub CheckBlock_Change()
    Debug.Print CheckBlock.Value
End Sub

'=========================================================== CREATE ===========================================================
Public Sub CreateLabel(ByVal Caption As String, Width As Long, Height As Long, _
                       Optional TopOffset As Long = 0, Optional LeftOffset As Long = 0)
    With Parent.Controls
    Set LabelBlock = .Add("Forms.Label.1", Join(Array("Label", Name, Row, Column), "_"), True)
    End With
    On Error Resume Next
    With LabelBlock
        .Top = TopOffset + ((Height + 10) * Row)
        .Left = LeftOffset + ((Width + 10) * Column)
        .Height = Height
        .Width = Width
        .Wrap = True
        .Caption = Caption
        .BackColor = BACKGRAY
        .ForeColor = FOREPURPLE
        .SpecialEffect = fmSpecialEffectFlat
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .KeepScrollBarsVisible = fmScrollBarsNone
        .ZOrder = fmZOrderFront
    If IsOrientVert Then
        .Top = TopOffset + ((Height * 2 + 10) * Row)
    Else
        .Left = LeftOffset + ((Width * 2 + 10) * Column)
    End If
    End With
End Sub

Private Sub setInput()
    With Parent.Controls
    Select Case m_ArgType
    Case "Boolean"
        Set Self = .Add("Forms.CheckBox.1", Join(Array("Label", Name, Row, Column), "_"), True)
        Set CheckBlock = Self
    Case "FuncName"
        m_BlockType = "ComboBox"
        Set Self = .Add("Forms.ComboBox.1", Join(Array("Label", Name, Row, Column), "_"), True)
        Set SelectBlock = Self
    Case Else
        m_BlockType = "TextBox"
        Set Self = .Add("Forms.TextBox.1", Join(Array("Label", Name, Row, Column), "_"), True)
        Set TextBlock = Self
    End Select
    End With
End Sub

Public Sub CreateInput(ByVal Caption As String, ByVal Place As String, _
                       Width As Long, Height As Long, Optional IdxVert As Long = 0, Optional IdxHoriz As Long = 0, _
                       Optional TopOffset As Long = 0, Optional LeftOffset As Long = 0)
    On Error Resume Next
    setInput
    With Self
        .Top = TopOffset + ((Height + 10) * IdxVert)
        .Left = LeftOffset + ((Width + 10) * IdxHoriz)
        .Height = Height
        .Width = Width
        .Wrap = True
        .Caption = Caption
        .BackColor = BACKGRAY
        .ForeColor = FOREPURPLE
        .SpecialEffect = fmSpecialEffectFlat
        .TextAlign = fmTextAlignLeft
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .KeepScrollBarsVisible = fmScrollBarsNone
        .ZOrder = fmZOrderFront
        Select Case Place
        Case "Left"
            .Left = LeftOffset + ((Width + 10) * 2 * IdxHoriz)
        Case "Top"
            .Top = TopOffset + ((Height + 10) * 2 * IdxVert)
        End Select
    End With
End Sub

