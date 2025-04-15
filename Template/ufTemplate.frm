'@Folder("Template")
Option Explicit
Private WithEvents PagesList As MSForms.MultiPage
Private WithEvents successBtn As MSForms.CommandButton
Private TemplateTitle As String
Private elements() As Variant
Private Filename As String

Private Sub successBtn_Click()
    Dim ctrl As clsUfTemplRow
    Dim curr As Variant
    Dim output() As Variant
    ReDim output(1 To UBound(elements))
    Dim i As Long, j As Long, lastCol As Long
    For i = 1 To UBound(output)
        Set ctrl = elements(i)
        If lastCol < ctrl.ParseModArg.count + 4 Then lastCol = ctrl.ParseModArg.count + 4
        ReDim curr(1 To lastCol)
        If ctrl.ParseModArg.count = 0 Then GoTo Continue
        curr(1) = TemplateTitle
        curr(2) = ctrl.Title
        curr(3) = ctrl.Func
        curr(4) = ctrl.InitColumn
        For j = 1 To ctrl.ParseModArg.count
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
    Dim Row As clsUfTemplRow
    Dim i As Long, j As Long
    Dim VarName As Variant
    Dim Table As Variant
    Dim CurrPage As MSForms.Page
    With Me
        .Caption = Title
        With .Controls
            Set PagesList = .Add("Forms.MultiPage.1", "MultiPage")
            PagesList.Pages.Clear
            Set CurrPage = PagesList.Pages.Add("Vars_1", "Переменные " & 1, 0)
            If Not VarDict.Exists("Variables") Then GoTo NoVariables
            ReDim elements(1 To VarDict("Variables").count)
            For Each VarName In VarDict("Variables").Keys
                Set Row = New clsUfTemplRow
                Row.Init CurrPage, VarName, Join(Array("Var", VarDict("Variables")(VarName), i), "_"), i
                i = i + 1
                Set elements(i) = Row
            Next
            
NoVariables:
            If Not VarDict.Exists("Tables") Then GoTo NoTables
            For Each Table In VarDict("Tables")
                j = j + 1
                Set CurrPage = PagesList.Pages.Add("Tbl_" & j, "Таблица " & j, j)
                If IsEmpty(elements) Then
                    ReDim elements(1 To Table.count)
                Else
                    ReDim Preserve elements(1 To UBound(elements) + Table.count)
                End If
                For Each VarName In Table.Keys
                    Set Row = New clsUfTemplRow
                    Row.Init CurrPage, VarName, Join(Array("Table", j, Table(VarName), i), "_"), Table(VarName)
                    i = i + 1
                    Set elements(i) = Row
                Next
            Next
NoTables:
            Set successBtn = .Add("Forms.CommandButton.1", "CommandButton_Success")
            CtrlDefaultParams successBtn, 90, 40, i * 30 + 20, 150, "Сохранить"
            CtrlDefaultParams PagesList, 480, i * 30 + 20, 0, 0
        End With
        
        .Width = 480
        .Height = i * 30 + 90
    End With
End Sub

Sub ReadTemplates()
    Dim templ As New clsTemplate
    Dim dict As Dictionary
    
    On Error GoTo BeforeExit
'    templ.Import
'    TemplateTitle = templ.Name
'    Set dict = templ.ParseFile
        TemplateTitle = "Мок"
        Set dict = MockTemplate
    With ufTemplate
        .DrawTemplate TemplateTitle, dict
        .Show vbModal
    End With
BeforeExit:
End Sub
