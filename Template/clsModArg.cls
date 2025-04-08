'@Folder("Template")
Option Explicit

Private WithEvents ctrl As MSForms.textbox
Private m_IsOptional As Boolean
Private m_ArgType As String
Private m_ArgPlaceholder As String

Public Property Get ArgType() As String
    ArgType = m_ArgType
End Property


Public Sub Add(ByVal arg As String, ByVal IsOptional As Boolean)
    m_ArgType = arg
    m_IsOptional = IsOptional
    Select Case m_ArgType
    Case "Long"
        m_ArgPlaceholder = "Число"
    Case "String"
        m_ArgPlaceholder = "Текст"
    Case "Variant"
        m_ArgPlaceholder = "Текст"
    Case "Boolean"
        m_ArgPlaceholder = "Текст"
    Case "FuncName"
        With Page
            Set ctrl = .Controls.Add("Forms.ComboBox.1", "ComboBox_" & Join(Array("Arg", Index, Func, ColIndex), "_"))
            ctrl.List = modifiersList.Keys
            CtrlDefaultParams ctrl, Width, Height, Index * 30 + 10, 220 + 70 * ColIndex, , Index
            ctrl.TextAlign = fmTextAlignLeft
        End With
    Case Else
        Debug.Print "Тип переменной "; arg
    End Select
    If Len(helper) > 0 Then Set ctrl = AddTextBox(Join(Array("Arg", Index, Func, ColIndex), "_"), 220 + 70 * ColIndex, m_ArgPlaceholder, False)
End Sub
