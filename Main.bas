'@Folder("VBAProject")
Option Explicit
'=========================================================== CONSTANTS ===========================================================
Public Const BACKGRAY As Long = &HE0E0E0
Public Const FOREPURPLE As Long = &H8000000D
Public Const FOREGRAY As Long = &H80000006
Public Const DEV_PASSWORD As String = "qwerty"

Public Const RePattern_prefix = "ШАБЛОН_(.+)\.docx?$"
Public Const RePattern_variable = "<<(.+)>>"
Public Const RePattern_email As String = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}"
Public Const RePattern_rngstring As String = "(\'[A-ZА-ЯЁ\s0-9]+\'\!)?\$?[A-Z]{1,3}\$?\d+"
Public Const RePattern_cell As String = "\$?[A-Z]{1,3}\$?\d+"
'=========================================================== TYPE ===========================================================
'=========================================================== ENUM ===========================================================
Public Enum EORientation
    Vertical = 0
    Horizontal = 1
End Enum

Public Enum EDimention
    ERow = 2
    ECol = 1
End Enum

'=========================================================== OPERATIONS ===========================================================
'вкл/выкл различные функции отрисовки
'Вызывать до и после скрипта
Public Sub SwitchAutomation(Optional state As Boolean = False)
    If state = True Then Application.Calculation = xlCalculationManual
    If state = False Then Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = state
    Application.DisplayStatusBar = state
    Application.DisplayAlerts = state
    Application.EnableEvents = state
End Sub

Public Function MainDataHeaderToDict() As Dictionary
    Dim headerRng As Range
    Dim HeaderArr As Variant
    Set headerRng = RangeGroupMerged(ShtMainData.Cells(1).CurrentRegion)
    HeaderArr = headerRng.Value2
    Set MainDataHeaderToDict = ArrayToDictByCol(HeaderArr)
End Function

'=========================================================== WRAPPERS ===========================================================
'Обертка для запуска скриптов
Public Sub ScriptRunner(ByVal SubName As String, ParamArray SubArguments() As Variant)
    On Error GoTo BeforeExit
    SwitchAutomation

    Application.Run SubName, SubArguments
BeforeExit:
    If Err.Number <> 0 Then Debug.Print SubName; ": "; Err.Number; vbNewLine; Err.Description
    SwitchAutomation True

End Sub

'Обертка для замера скорости работы скриптов скриптов
Public Sub Stopwatch(ByVal SubName As String, ParamArray SubArguments() As Variant)
    Dim X As Single
    X = Timer
    ScriptRunner SubName
    Debug.Print MySub; "executed in "; Timer - X
End Sub

'Обертка для запуска модификаторов
Public Function FnWrap(ByVal FnName As String, ByVal arr As Variant, ByVal Args As Variant) As Variant
    FnWrap = Application.Run(FnName, arr, Args)
End Function

'=========================================================== UserForm ===========================================================
'Если поле ввода пустое
Public Function isControlEmpty(Control As Object) As Boolean
    isControlEmpty = (Len(Trim(Control.value)) = 0)
End Function

'=========================================================== BASE64 ===========================================================
'String -> BASE64
Public Function EncodeBase64(ByVal text$)
    Dim b
    With CreateObject("ADODB.Stream")
        .Open: .Type = 2: .Charset = "utf-8"
        .WriteText text: .Position = 0: .Type = 1: b = .Read
        With CreateObject("Microsoft.XMLDOM").createElement("b64")
            .DataType = "bin.base64": .nodeTypedValue = b
            EncodeBase64 = Replace(Mid(.text, 5), vbLf, "")
        End With
        .Close
    End With
End Function

'Base64 -> String
Public Function DecodeBase64(ByVal b64$)
    Dim b
    With CreateObject("Microsoft.XMLDOM").createElement("b64")
        .DataType = "bin.base64": .text = b64
        b = .nodeTypedValue
        With CreateObject("ADODB.Stream")
            .Open: .Type = 1: .Write b: .Position = 0: .Type = 2: .Charset = "utf-8"
            DecodeBase64 = .ReadText
            .Close
        End With
    End With
End Function

'=========================================================== SHEET ===========================================================
'Codename назначается в настройках worksheet('(Name)')
Public Function SheetByCodename(ByVal codename As String) As Worksheet
    Dim Index As Long

    Index = ThisWorkbook.VBProject.VBComponents(codename).Properties("Index").value
    Set SheetByCodename = Worksheets(Index)
End Function

'=========================================================== RANGE ===========================================================

Public Function RangeByName(name As String, Optional DefaultRng As Range) As Range
    On Error GoTo Default
    Set RangeByName = ThisWorkbook.Names(name).RefersToRange
    Exit Function
Default:
    If Not DefaultRng Is Nothing Then Set RangeByName = DefaultRng
End Function

'Достать Range по координатам, нужно указать конечные строка,колонка, начальные по умолчанию 1
Public Function RangeGetSubRange(rng As Range, endRow As Long, endCol As Long, Optional startRow As Long = 1, Optional startCol As Long = 1) As Range
    With rng
        Set RangeGetSubRange = .Range(.Cells(startRow, startCol), .Cells(endRow, endCol))
    End With
End Function

'Вырезать из Range SubRange
Public Function RangeExclude(rng As Range, Optional RngMinus As Range, Optional cutRow As Long, Optional cutCol As Long) As Range
    If Not RngMinus Is Nothing Then
        If rng.Rows.Count > RngMinus.Rows.Count Then cutRow = RngMinus.Rows.Count
        If rng.Columns.Count > RngMinus.Columns.Count Then cutCol = RngMinus.Columns.Count
    End If
    With rng
        Set RangeExclude = .Resize(.Rows.Count - cutRow, .Rows.Columns.Count - cutCol).offset(cutRow, cutCol)
    End With
End Function

'Проходит по первому ряду и возвращает весь range заголовка(включая объединения)
Public Function RangeGetMerged(rng As Range) As Range
    Dim col As Long
    Dim row As Long
    With rng
        With .Rows(1)
            For col = 1 To .Columns.Count
                If Not .Columns(col).MergeCells Then GoTo Continue
                col = col + .Columns(col).MergeArea.Columns.Count - 1
                If .Columns(col).MergeArea.Rows.Count > row Then row = .Columns(col).MergeArea.Rows.Count
Continue:
            Next
        End With
        Set RangeGetMerged = .Range(.Cells(1, 1), .Cells(row, col - 1))
    End With
End Function

'Использовать аккуратно, предпочтительно в режиме разработчика
Public Sub RangeJumpAndSelect(rng As Range)
    Application.Goto rng.Cells(1, 1), Scroll:=True
    rng.Select
    Stop
End Sub

'=========================================================== FILE ===========================================================
'Если AskUser=True то спрашивает где сохранить файл(при отмене укажет путь в текущую папку)
Public Function FileNameCreate(Optional ByVal FileName As String, Optional AskUser As Boolean = False, Optional Title As String = "Выберите путь для сохранения") As String
    Dim Dir As FileDialog
    Dim sItem As String
    If AskUser = True Then
        Set Dir = Application.FileDialog(msoFileDialogFolderPicker)
        With Dir
            .Title = Title
            .AllowMultiSelect = False
            .InitialFileName = FilePathCurrent(FileName)
            If .Show <> -1 Then GoTo NoPath
            FileNameCreate = .SelectedItems(1)

            Exit Function
        End With
    End If
NoPath:
    Set Dir = Nothing
    FileNameCreate = FilePathCurrent(FileName)
End Function

Public Function FilePathCurrent(Optional FileName As String) As String
    Dim Path As String
    Path = ThisWorkbook.Path
    If Len(FileName) > 0 Then Path = Path & "\" & FileName
    FilePathCurrent = Application.Clean(Path)
End Function

'=========================================================== Chunks ===========================================================
Public Function ChunksParse(rng As Range, ColNum As Long, Optional ColSize As Long) As Dictionary
    Set ChunksParse = New Dictionary
    Dim FilterRng As Range
    Dim ChunkRng As Range
    Dim Chunk As clsChunk
    Dim lastrow As Long, i As Long

    With rng
        lastrow = .Rows.Count + .row - 1
        'note: counts rm columns from 1
        Set FilterRng = .Columns(ColNum)
        If ColSize > 1 Then Set FilterRng = FilterRng.Resize(ColumnSize:=ColSize)
        With FilterRng.SpecialCells(xlCellTypeConstants, 7)

            For i = .Areas.Count To 1 Step -1
                Set ChunkRng = .Areas(i).Resize(lastrow - .Areas(i).row + 1, rng.Columns.Count)
                Set Chunk = New clsChunk
                Chunk.Create ChunkRng
                If ChunksParse.Exists(Chunk.ID) Then GoTo ChunksDuplicated
                ChunksParse.Add Chunk.ID, Chunk
                lastrow = .Areas(i).row - 1
            Next
        End With
    End With
    Exit Function
ChunksDuplicated:
    Dim oldChunk As clsChunk
    Set oldChunk = ChunksParse.Item(Chunk.ID)
    MsgBox Join(Array("Есть несколько РМ с одним ID", ChunkRng.Address, oldChunk.Range.Address), vbNewLine), Title:="Дублирование РМ"
End Function

'=========================================================== ARRAY ===========================================================
Public Function ArrayLen(ByRef arr As Variant) As Long
    LenArray = UBound(arr) - LBound(arr) + 1
End Function

Public Function ArrayContains(SearchVal As Variant, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = SearchVal Then
            ArrayContains = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Function ArraySlice(ByRef arr As Variant, Optional startPos As Long, Optional endPos As Long) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim diff As Long
    diff = 1 - LBound(arr)
    If startPos = 0 Then startPos = 1
    If endPos = 0 Then endPos = UBound(arr) + diff - startPos
    If startPos >= UBound(arr) + diff Or endPos > UBound(arr) + diff Then Exit Function
    ReDim outputArr(endPos - diff)
    For i = 1 To endPos
        outputArr(i - 1) = arr(startPos + i - diff)
    Next
    ArraySlice = outputArr
End Function

Public Function Array2DSlice(ByRef arr As Variant, dimention As EDimention, idx As Long) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim diff As Long
    diff = 1 - LBound(arr, dimention)
    ReDim outputArr(1 To UBound(arr, dimention) + diff)
    For i = 1 To UBound(outputArr)
        If dimention = ERow Then
            outputArr(i) = arr(idx, i)
        Else
            outputArr(i) = arr(i, idx)
        End If
    Next
    Array2DSlice = outputArr
End Function

Public Function ArrayUpdateColumns(arr As Variant, ByVal FnDict As Dictionary) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim colIdx As Variant
    Dim offset As Long
    Dim Done As New Dictionary
    If FnDict.Count = 0 Then
        ArrayUpdateColumns = arr
        Exit Function
    End If
    ReDim outputArr(LBound(arr) To UBound(arr))
    For i = LBound(outputArr) To UBound(outputArr)
        If FnDict.Exists(i) And Not Done.Exists(i) Then
            outputArr(i + offset) = Application.Run(FnDict(i)("Name"), arr, FnDict(i)("Args"))
            If FnDict(i)("NewCol") Then
                offset = offset + 1
                ReDim Preserve outputArr(LBound(arr) To UBound(arr) + offset)
                Done.Add i, True
                i = i - 1
            End If
        Else
            On Error GoTo BeforeExit
            outputArr(i + offset) = arr(i)
        End If

    Next
BeforeExit:
    Done.RemoveAll
    Err.Clear
    On Error GoTo 0
    ArrayUpdateColumns = outputArr
End Function

Public Function Array2DIterRows(ByRef arr As Variant, Optional ByVal offset As Long, Optional UpdateFuncs As Dictionary) As Variant
    Dim outputArr As Variant
    Dim rowArr As Variant
    Dim col As Variant
    Dim FnArgs As Variant
    Dim i As Long
    ReDim outputArr(1 To 1)
    Do
        i = i + 1
        On Error GoTo Break
        rowArr = Array2DSlice(arr, ERow, offset + i)
        If Not UpdateFuncs Is Nothing Then rowArr = ArrayUpdateColumns(rowArr, UpdateFuncs)
        
        ReDim Preserve outputArr(1 To i)
        outputArr(i) = rowArr
Break:
    Loop Until Err.Number <> 0
    Err.Clear
    Array2DIterRows = outputArr
End Function

'Группирует по первой колонке и вкладывает словари (ключ = Текст, значение = номер колонки)
Public Function Array2DToDictByCol(arr As Variant) As Dictionary
    Dim outputDict As New Dictionary
    Dim dict As Dictionary
    Dim col As Long
    Dim row As Long
    Dim Title As Variant
    
    For col = 1 To UBound(arr, 2)
        If Not IsEmpty(arr(1, col)) Then Title = arr(1, col)
        For row = 1 To UBound(arr, 1)
            If IsEmpty(arr(row, col)) Then GoTo Continue2
            If Not outputDict.Exists(Title) Then
                Set dict = New Dictionary
            Else
                Set dict = outputDict(Title)
            End If
            If Not dict.Exists(arr(row, col)) Then dict.Add arr(row, col), col
Continue2:
        Next
        Set outputDict(Title) = dict
    Next
    
    Set Array2DToDictByCol = outputDict
End Function

'Объединяет список со списками в 2мерный список
Public Function ArrayMergeTo2D(arr As Variant, ParamArray exclude() As Variant) As Variant
    Dim outputArr As Variant
    Dim i As Long, j As Long
    i = LBound(arr)
    ReDim outputArr(LBound(arr) To UBound(arr), LBound(arr(i)) To UBound(arr(i)))
    For i = LBound(outputArr, 1) To UBound(outputArr, 1)
        For j = LBound(outputArr, 2) To UBound(outputArr, 2)
            outputArr(i, j) = arr(i)(j)
        Next
    Next
    ArrayMergeTo2D = outputArr
End Function

'===================================================== MODIFICATORS SUBS/FUNCS =====================================================
'Добавление в словарь функции для обработки данных
Public Sub ModAddParam(ByRef dict As Dictionary, ByVal FnName As String, ByVal FnDesc As String, ByVal FnArgs As Variant)
    Dim subDict As New Dictionary
    Dim argslistDict As New Dictionary
    Dim argDict As Dictionary
    Dim argParams As Variant
    Dim curr As Variant
     
    If IsEmpty(FnArgs) Then GoTo BeforeExit
    For Each curr In FnArgs
        Set argDict = New Dictionary
        argParams = Split(curr, "=")
        argDict.Add "Optional", CBool(InStr(argParams(1), "Optional "))
        argDict.Add "Type", Replace(argParams(1), "Optional ", vbNullString)
        argslistDict.Add argParams(0), argDict
    Next
BeforeExit:
    subDict.Add "Name", FnName
    subDict.Add "Args", argslistDict
    
    dict.Add FnDesc, subDict
End Sub

'Получить список актуальных Функций
Public Function ModList() As Dictionary
    Dim Fn As Variant
    Dim subDict As Dictionary
    Set ModList = New Dictionary
    'Список аргументов является списком строк с маской "ИмяАргумента=[Optional ]Тип",
    'например Array("Аргумент 1=String", "Аргумент 2=Optional Variant", "Аргумент 3=Optional Long")
    For Each Fn In Array( _
        Array("autoIncrement", "Номер по порядку", Array("Начальное значение=Optional Long")), _
        Array("updateValue", "Фиксированное значение", Array("Номер Колонки=Long", "Значение=Variant", "Менять не пустые=Boolean")), _
        Array("addPrefix", "Добавить префикс", Array("Номер Колонки=Long", "Префикс=String")), _
        Array("checkCondition", "Если(Условие как в формуле, вместо ячейки значения target - ""AND(target > 1,target < 5)"")", Array("Номер Колонки=Long", "Условие=String", "Если ИСТИНА=FuncName", "Параметры=Variant")) _
        )
        ModAddParam ModList, Fn(0), Fn(1), Fn(2)
    Next
End Function

'Добавить настройки функции
Public Sub FnUpdateAdd(ByRef dict As Dictionary, ByVal ColNum As Long, ByVal NewCol As Boolean, ByVal FnName As String, ByVal FnArgs As Variant)
    Dim subDict As New Dictionary
    subDict.Add "Name", FnName
    subDict.Add "NewCol", NewCol
    subDict.Add "Args", FnArgs
    dict.Add ColNum, subDict
End Sub

'=========================================================== MODIFICATORS ===========================================================
'Это функции, которые нужны для работы со значениями внутри таблицы
'Передаются строкой с названием
'ВАЖНО: Принимают аргументами 2 массива - текущий массив и массив аргументов

''
'Note: Форма написания функции:
'Public Function [Name](ByRef rowArr As Variant, Args As Variant) as [Type]
' Dim [ArgName] as [ArgType]                    '[index]
' ArgName = Args(index)
' --- do something
'End Function
''

'Инкрементирует автоматически(статическая функция запоминает значения внутри runtime)
Static Function autoIncrement(ByRef rowArr As Variant, Args As Variant) As Long
    Dim n As Long
    Dim start As Long                            '0
    If Not IsEmpty(Args) Then start = Args(0)
    If start > n Then n = start
    n = n + 1
    autoIncrement = n
End Function

'Назначаем значение по умолчанию если ячейка пуста
Public Function updateValue(ByRef rowArr As Variant, Args As Variant) As Variant
    Dim ColNum As Long                           '0
    Dim value As Variant                         '1
    Dim NotEmptyUpdate As Boolean                '2
    ColNum = Args(0)
    value = Args(1)
    NotEmptyUpdate = Args(2)
    If NotEmptyUpdate = True Or Len(rowArr(ColNum)) = 0 Then
        updateValue = value
    Else
        updateValue = rowArr(ColNum)
    End If
End Function

'Добавляем префикс с возможным условием
Public Function addPrefix(ByRef rowArr As Variant, Args As Variant) As String
    Dim ColNum As Long                           '0
    Dim Prefix As String                         '1
    On Error GoTo Done
    ColNum = Args(0)
    Prefix = Args(1)
    If VarType(rowArr(ColNum)) <> vbString Then
        rowArr(ColNum) = CStr(rowArr(ColNum))
    End If
    addPrefix = Prefix & rowArr(ColNum)
Done:
End Function

'Проверяем условие для значения и назначаем положительные и отрицательные значения
Public Function checkCondition(ByRef rowArr As Variant, Args As Variant) As Variant
    Const keyword As String = "target"
    Dim output As Variant
    Dim ColNum As Long                           '0
    Dim condition As String                      '1
    Dim FnNameOnTrue As String                   '2
    Dim FnArgs As Variant                        '3
    ColNum = Args(0)
    condition = Args(1)
    If UBound(Args) > 1 Then
        FnNameOnTrue = Args(2)
        FnArgs = Args(3)
    End If
    
    condition = "=IF(" & condition & ", True, False)"
    condition = WorksheetFunction.Substitute(condition, keyword, rowArr(ColNum))
    output = Evaluate(condition)
    If Len(FnNameOnTrue) > 0 And output = True Then checkCondition = Application.Run(FnNameOnTrue, rowArr, FnArgs)
    If Len(FnNameOnTrue) = 0 Then checkCondition = output
End Function

Public Sub CtrlDefaultParams(ctrl As Control, ByVal Width As Long, ByVal Height As Long, ByVal Top As Long, ByVal Left As Long, Optional ByVal Caption As String = vbNullString, Optional ByVal Tag As String = vbNullString)
    On Error Resume Next
    With ctrl
        .Top = Top
        .Left = Left
        .Height = Height
        .Width = Width
        .Wrap = True
        .Caption = Caption
        .BackColor = BACKGRAY
        .ForeColor = FOREGRAY
        .SpecialEffect = fmSpecialEffectFlat
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleOpaque
        .BorderStyle = fmBorderStyleNone
        .ListWidth = 400 'combobox ширина выпадающего списка
    End With
    On Error GoTo 0
End Sub

'*********************************************************** TEST ***********************************************************
Sub testArr()
    Dim rng As Range
    Dim i As Long
    Dim arr As Variant
    Dim iterArr As Variant
    Dim result As Variant
    Dim FnDict As Dictionary
    Set FnDict = ModList
    Dim UpdateDict As New Dictionary
    'Нужно учитывать, что если добавляем колонку, индекс последующих увеличивается
    For Each iterArr In Array( _
        Array(1, True, "autoIncrement", Empty), _
        Array(2, False, "addPrefix", Array(2, "Pre_")), _
        Array(4, True, "checkCondition", Array(4, "target = 1", "updateValue", Array(2, "Бюро", True))) _
        )
        FnUpdateAdd UpdateDict, iterArr(0), iterArr(1), iterArr(2), iterArr(3)
    Next
    With ShtMainData.Cells(1).CurrentRegion
        '        Set rng = RangeExclude(.Cells, RangeGetMerged(.Cells))
        arr = Array2DIterRows(.Cells.Value2, RangeGetMerged(.Cells).Rows.Count, UpdateDict)
    End With
    result = ArrayMergeTo2D(arr)

End Sub

Sub testrun()
    Dim dataRng As Range
    Dim chunks As Dictionary
    With ShtMainData.Cells(1).CurrentRegion
        Set dataRng = RangeExclude(.Cells, RangeGetMerged(.Cells))
    End With
    
    Set chunks = ChunksParse(dataRng, 1)
    Stop
End Sub

'*********************************************************** TEST ***********************************************************










