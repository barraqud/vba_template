'@Folder("VBAProject")
Option Explicit


Public Sub ValidateSubStrExist(ByRef TargetRng As Range)
    Dim errorsArr() As Variant
    Dim elem As Variant
    Dim Value As String
    Dim i As Long

    For Each elem In TargetRng.Cells
        Value = elem.Value2
        On Error GoTo addErr
        If InStr(Value, "Существующие: ") = 0 _
           Or UBound(Split(Value, Chr(10))) <> 1 _
           Or InStr(Value, "Отсутствующие: ") = 0 Then
addErr:
            i = i + 1
            ReDim Preserve errorsArr(1 To i)
            errorsArr(i) = elem.Address
        End If
    Next
    If i <> 0 Then MsgBox "Не хватает ключевых слов" & vbCrLf & "'Существующие: ' и 'Отсутствующие: ', " & vbCrLf & "между которыми должен быть один перенос строк(Alt+Enter)" & vbCrLf & "В этих ячейках: " & vbCrLf & Join(errorsArr, ", "), Title:="Неверно введены данные"
End Sub

Public Sub ValidationDependentUpdate(TargetRng As Range, OffsetFromParent As Long, KeyRngName As String, ValRngName As String)
    Dim Grouped As Object
    Set Grouped = CreateObject("Scripting.Dictionary")
    
    Dim KeyArr As Variant
    Dim ValArr As Variant
    Dim key As String
    Dim elem As Variant
    
    Dim validationString As String
    
    KeyArr = ThisWorkbook.Names(KeyRngName).RefersToRange.Value2
    KeyArr = ArrayTranspose(KeyArr)
    ValArr = ThisWorkbook.Names(ValRngName).RefersToRange.Value2
    ValArr = ArrayTranspose(ValArr)
    
    'note: ВАЖНО!!!!!! Меняет "," на Chr(130) (Возможны ошибки)
    Set Grouped = ArrayToDict(KeyArr, ValArr, "StringEscComma")
    
    For Each elem In TargetRng
        key = Trim(elem.Value2)
        If Not Grouped.Exists(key) Then GoTo Continue
        validationString = Join(Grouped(key), ",")
        
        With elem.offset(columnoffset:=OffsetFromParent).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=validationString
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
        End With
Continue:
    Next
End Sub

Public Sub ValidationUpdate(Payload As String, columnNum As Long, Optional RowsCount As Long, Optional ByRangeName As Boolean = True)
    Dim validationString As String
    
    If ByRangeName Then
        With Evaluate("=" & Payload)
            validationString = "='" & .Parent.Name & "'!" & .Address(External:=False)
        End With
    Else: validationString = Payload
    End If
    With ShtMainData.Columns(columnNum)
        If RowsCount = 0 Then RowsCount = .Rows.count
        With .Resize(RowsCount - 3).offset(3).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=validationString
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
        End With
    End With
End Sub

Public Sub ValidationClear(columnNum As Long)
    With ShtMainData.Columns(columnNum)
        .Resize(.Rows.count - 3).offset(3).Validation.Delete
    End With
End Sub

Public Function getLastRow(Optional withScroll As Boolean = True) As Range
    With ShtMainData.Cells(1, 1).CurrentRegion
        Set getLastRow = .Rows(.Rows.count + 1)
        If withScroll Then Application.Goto getLastRow, False
    End With
End Function

Public Sub FormatByPreset(presetRng As Range, TargetRng As Range)
    Dim tCell As Variant
    Dim ValArr As Variant
    Dim i As Long
    ValArr = Application.Transpose(presetRng.Cells.Value)
    For Each tCell In TargetRng.Cells
        tCell.Interior.ColorIndex = xlNone
        For i = 1 To UBound(ValArr)
            If tCell.Value <> ValArr(i) Then GoTo Continue
            tCell.Interior.Color = presetRng.Cells(i).Interior.Color
            GoTo NextCell
Continue:
        Next
NextCell:
    Next
    
End Sub

