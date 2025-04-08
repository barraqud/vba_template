'@Folder("Template")
Option Explicit
Private WithEvents PagesList As MSForms.MultiPage
Private WithEvents successBtn As MSForms.CommandButton
Private TemplateTitle As String
Private elements As Variant

Private Sub successBtn_Click()
    Dim ctrl As clsUfTemplRow
    Dim curr As Variant
    Dim output() As Variant
    ReDim output(1 To UBound(elements))
    Dim i As Long, j As Long, lastCol As Long
    For i = 1 To UBound(output)
        Set ctrl = elements(i)
        If lastCol < ctrl.ParseModArg.Count + 4 Then lastCol = ctrl.ParseModArg.Count + 4
        ReDim curr(1 To lastCol)
        If ctrl.ParseModArg.Count = 0 Then GoTo Continue
        curr(1) = TemplateTitle
        curr(2) = ctrl.Title
        curr(3) = ctrl.Func
        curr(4) = ctrl.InitColumn
        For j = 1 To ctrl.ParseModArg.Count
            curr(j + 4) = ctrl.ParseModArg.Items(j - 1)
        Next
Continue:
        output(i) = curr
    Next
    
End Sub

Private Sub UserForm_Initialize()
    Dim dict As New Dictionary
    With Me
        .KeepScrollBarsVisible = fmScrollBarsNone
        .BackColor = RGB(256, 256, 256)
        .StartUpPosition = 1
    End With
End Sub

Public Sub DrawTemplate(Title As String, VarDict As Dictionary)
    Dim row As clsUfTemplRow
    Dim i As Long, j As Long
    Dim VarName As Variant
    Dim Table As Variant
    Dim CurrPage As MSForms.Page
    With Me
        With .Controls
            Set PagesList = .Add("Forms.MultiPage.1", "MultiPage")
            PagesList.Pages.Clear
            CtrlDefaultParams PagesList, 480, 390, 0, 0
            PagesList.BackColor = RGB(256, 256, 30)
            Set CurrPage = PagesList.Pages.Add("Vars_1", "Переменные " & 1, 0)
            If Not VarDict.Exists("Variables") Then GoTo NoVariables
            ReDim elements(1 To VarDict("Variables").Count)
            For Each VarName In VarDict("Variables").Keys
                Set row = New clsUfTemplRow
                row.Init CurrPage, VarName, Join(Array("Var", VarDict("Variables")(VarName), i), "_"), i
                i = i + 1
                Set elements(i) = row
            Next
NoVariables:
            If Not VarDict.Exists("Tables") Then GoTo NoTables
            For Each Table In VarDict("Tables")
                j = j + 1
                Set CurrPage = PagesList.Pages.Add("Tbl_" & j, "Таблица " & j, j)
                If IsEmpty(elements) Then
                    ReDim elements(1 To Table.Count)
                Else
                    ReDim Preserve elements(1 To UBound(elements) + Table.Count)
                End If
                For Each VarName In Table.Keys
                    Set row = New clsUfTemplRow
                    row.Init CurrPage, VarName, Join(Array("Table", j, Table(VarName), i), "_"), Table(VarName)
                    i = i + 1
                    Set elements(i) = row
                Next
            Next
NoTables:
            Set successBtn = .Add("Forms.CommandButton.1", "CommandButton_Success")
            CtrlDefaultParams successBtn, 90, 40, i * 30 + 20, 150, "Сохранить"
        End With
        
        .Width = 480
        .Height = i * 30 + 90
    End With
End Sub

Private Sub GenerateVariable(Frame As MSForms.Frame, ByVal var As String, ByVal VarName As String, ByVal idx As Long)
    Dim label As MSForms.label
    Dim textbox As MSForms.textbox
    Dim cellBtn As refEdit.refEdit
    With Frame
        Set label = .Controls.Add("Forms.Label.1", "Label_" & var)
        With label
            .Caption = VarName
            .Tag = var
            .TextAlign = fmTextAlignCenter
            .SpecialEffect = fmSpecialEffectFlat
            .BackStyle = fmBackStyleTransparent
            .Height = 20
            .Width = 70
            .Left = 10
            .Top = idx * 30 + 10
        End With
        Set cellBtn = .Controls.Add("RefEdit.Ctrl", "RefEdit_" & var, True)
        With cellBtn
            .SpecialEffect = fmSpecialEffectFlat
            .BackColor = RGB(256, 256, 256)
            .Height = 20
            .Width = 70
            .Left = 100
            .Top = idx * 30 + 10
        End With
        Set textbox = .Controls.Add("Forms.TextBox.1", "TextBox_" & var)
        With textbox
            .Tag = VarName
            .SpecialEffect = fmSpecialEffectFlat
            .BackStyle = fmBackStyleOpaque
            .BackColor = RGB(256, 256, 256)
            .Font.Size = 7
            .Height = 20
            .Width = 70
            .Left = 180
            .Top = idx * 30 + 10
        End With
    End With
End Sub
