'@Folder("VBAProject")
Option Explicit

Dim App As Word.Application
Dim Doc As Document

Private Sub Class_Initialize()

    On Error Resume Next
    Set App = GetObject(, "Word.Application")
    If App Is Nothing Then Set App = New Word.Application
    On Error GoTo 0

End Sub

Private Sub Class_Terminate()
    Dim el As Variant
    With App
        .Visible = True
        .DisplayAlerts = -1
        For Each el In .Documents
            el.Close False
        Next
        .Quit
    End With
    Set Doc = Nothing
    Set App = Nothing
End Sub

Public Sub Generate(rng As Range, FileName As String)
    
    On Error GoTo BeforeExit
    With App
        .Visible = True
        .DisplayAlerts = 0
        Set Doc = .Documents.Add
        With Doc
            rng.Copy
            .Paragraphs(1).Range.PasteExcelTable _
        LinkedToExcel:=False, _
        WordFormatting:=False, _
        RTF:=False
            .SaveAs "D:\WORK\VBA\Макет\" & FileName & ".docx"
            .Close
        End With
        .Visible = True
    End With
BeforeExit:
    Application.CutCopyMode = False
    App.DisplayAlerts = -1
End Sub

Public Function Read(FileName As String) As Document
        On Error GoTo BeforeExit
    With App
        .Visible = True
        .DisplayAlerts = 0
        Set Read = .Documents.Open(FileName)
    End With
BeforeExit:
    App.DisplayAlerts = -1
End Function
