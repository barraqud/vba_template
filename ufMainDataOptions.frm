'@Folder "Options"
Option Explicit



Private Sub UserForm_Initialize()
With Me
    .KeepScrollBarsVisible = fmScrollBarsNone
    .Width = 640
    .Height = 480
    .BackColor = RGB(256, 256, 256)
    .StartUpPosition = 1
End With
End Sub


Private Sub DrawGroup(UF As UserForm, Title As String, Blocks As Dictionary)

Dim block As Variant
    With UF.Controls
        For Each block In Blocks.keys
            block
        Next
    End With
End Sub


Private Sub DrawColumn(UF As UserForm)
    Dim dataOpts As Dictionary
    Dim elem As Variant
    Set dataOpts = HeaderParse
    Debug.Print Join(dataOpts.keys, vbNewLine)
End Sub

