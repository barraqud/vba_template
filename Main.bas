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
    Set headerRng = RangeGetMerged(ShtMainData.Cells(1).CurrentRegion)
    HeaderArr = headerRng.Value2
    Set MainDataHeaderToDict = ArrayToDictByCol(HeaderArr)
End Function

'=========================================================== WRAPPERS ===========================================================
'Обертка для запуска скриптов
'Первые 10 аргументов передаются
Public Sub ScriptRunner(ByVal SubName As String, isInitial As Boolean, ParamArray args() As Variant)
    Dim i As Long
    On Error Resume Next
    i = UBound(args) - LBound(args) + 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo BeforeExit
    If isInitial Then SwitchAutomation
    Select Case True
    Case IsMissing(args)
        Application.Run SubName
    Case i = 10
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)
    Case i = 9
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)
    Case i = 8
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)
    Case i = 7
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4), args(5), args(6)
    Case i = 6
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4), args(5)
    Case i = 5
        Application.Run SubName, args(0), args(1), args(2), args(3), args(4)
    Case i = 4
        Application.Run SubName, args(0), args(1), args(2), args(3)
    Case i = 3
        Application.Run SubName, args(0), args(1), args(2)
    Case i = 2
        Application.Run SubName, args(0), args(1)
    Case i = 1
        Application.Run SubName, args(0)
    Case Else
        Err.Raise vbObjectError + 500, "ScriptRunner", ">10 args"
    End Select
BeforeExit:
    If Err.Number <> 0 Then Debug.Print SubName; ": "; Err.Number; vbNewLine; Err.Description
    If isInitial Then SwitchAutomation True

End Sub

'Обертка для замера скорости работы скриптов скриптов
Public Sub Stopwatch(ByVal SubName As String, ParamArray SubArguments() As Variant)
    Dim X As Single
    X = Timer
    ScriptRunner SubName, True, SubArguments
    Debug.Print MySub; "executed in "; Timer - X
End Sub

'Обертка для запуска модификаторов
Public Function FnWrap(ByVal FnName As String, ByVal arr As Variant, ByVal args As Variant) As Variant
    FnWrap = Application.Run(FnName, arr, args)
End Function

'=========================================================== UserForm ===========================================================
'Если поле ввода пустое
Public Function isControlEmpty(Control As Object) As Boolean
    isControlEmpty = (Len(Trim(Control.Value)) = 0)
End Function

Public Function ControlCreate(Parent As Variant, bType As String, Name As String, Width As Long, Height As Long, _
                              Optional IdxVert As Long = 0, Optional IdxHoriz As Long = 0, _
                              Optional TopOffset As Long = 0, Optional LeftOffset As Long = 0, Optional CaptText As String) As MSForms.Control
    If InStr(bType, "Frame;TextBox;CheckBox;ComboBox;CommandButton;Page") <> 0 Then
        MsgBox "Неверный тип формы: " & bType
    End If
    Dim ctrl As MSForms.Control
    With Parent.Controls
        Set ctrl = .Add("Forms." & bType & ".1", Join(Array(bType, Name, IdxVert, IdxHoriz), "_"), True)
        CtrlDefaultParams ctrl, Width, Height, TopOffset + ((Height + 10) * IdxVert), LeftOffset + ((Width + 10) * IdxHoriz), CaptText, IdxVert
    End With
    Set ControlCreate = ctrl
End Function

Public Sub CtrlDefaultParams(ctrl As Variant, ByVal Width As Long, ByVal Height As Long, Optional ByVal Top As Long, Optional ByVal Left As Long, Optional ByVal Caption As String = vbNullString, Optional ByVal Tag As String = vbNullString)
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
        .KeepScrollBarsVisible = fmScrollBarsNone
        .ListWidth = 400                         'combobox ширина выпадающего списка
        If TypeOf ctrl Is UserForm Then
            .StartUpPosition = 1
        Else
            .ZOrder = fmZOrderFront
        End If
    End With
    On Error GoTo 0
End Sub

Public Function AddTextBox(Wrap As Variant, Name As String, Index As Long, Optional Width As Long = 60, Optional Height As Long = 20, Optional Left As Long = 80, Optional helper As String, Optional isDropDown As Boolean = False) As MSForms.TextBox
    Dim placeHolder As MSForms.TextBox
    Dim Box As MSForms.TextBox
    With Wrap
        Set placeHolder = .Controls.Add("Forms.TextBox.1", "Placeholder_" & Name)
        CtrlDefaultParams placeHolder, Width, Height, Index * 30 + 10, Left, , Index
        placeHolder.text = helper
        placeHolder.MultiLine = True
        placeHolder.TextAlign = fmTextAlignLeft
        placeHolder.Enabled = False
        
        Set Box = .Controls.Add("Forms.TextBox.1", Name)
        CtrlDefaultParams Box, Width, Height, Index * 30 + 10, Left, , Index
        If isDropDown Then
            Box.DropButtonStyle = fmDropButtonStyleReduce
            Box.ShowDropButtonWhen = fmShowDropButtonWhenAlways
        End If
        Box.BackStyle = fmBackStyleTransparent
    End With
    Set AddTextBox = Box
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

    Index = ThisWorkbook.VBProject.VBComponents(codename).Properties("Index").Value
    Set SheetByCodename = Worksheets(Index)
End Function

'=========================================================== RANGE ===========================================================

Public Function RangeByName(Name As String, Optional DefaultRng As Range) As Range
    On Error GoTo Default
    Set RangeByName = ThisWorkbook.Names(Name).RefersToRange
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
        If rng.Rows.count > RngMinus.Rows.count Then cutRow = RngMinus.Rows.count
        If rng.Columns.count > RngMinus.Columns.count Then cutCol = RngMinus.Columns.count
    End If
    With rng
        Set RangeExclude = .Resize(.Rows.count - cutRow, .Rows.Columns.count - cutCol).offset(cutRow, cutCol)
    End With
End Function

'Проходит по первому ряду и возвращает весь range заголовка(включая объединения)
Public Function RangeGetMerged(rng As Range) As Range
    Dim col As Long
    Dim Row As Long
    With rng
        With .Rows(1)
            For col = 1 To .Columns.count
                If Not .Columns(col).MergeCells Then GoTo Continue
                col = col + .Columns(col).MergeArea.Columns.count - 1
                If .Columns(col).MergeArea.Rows.count > Row Then Row = .Columns(col).MergeArea.Rows.count
Continue:
            Next
        End With
        Set RangeGetMerged = .Range(.Cells(1, 1), .Cells(Row, col - 1))
    End With
End Function

'Получить Исходные данные
Public Function RangeMainData() As Range
    With ShtMainData.Cells(1).CurrentRegion
        Set RangeMainData = RangeExclude(.Cells, RangeGetMerged(.Cells))
    End With
End Function

'Использовать аккуратно, предпочтительно в режиме разработчика
Public Sub RangeJumpAndSelect(rng As Range)
    Application.Goto rng.Cells(1, 1), Scroll:=True
    rng.Select
End Sub

'=========================================================== FILE ===========================================================
'Если AskUser=True то спрашивает где сохранить файл(при отмене укажет путь в текущую папку)
Public Function FileNameCreate(Optional ByVal Filename As String, Optional AskUser As Boolean = False, Optional Title As String = "Выберите путь для сохранения") As String
    Dim dir As FileDialog
    Dim sItem As String
    If AskUser = True Then
        Set dir = Application.FileDialog(msoFileDialogFolderPicker)
        With dir
            .Title = Title
            .AllowMultiSelect = False
            .InitialFileName = FilePathCurrent(Filename)
            If .Show <> -1 Then GoTo NoPath
            FileNameCreate = .SelectedItems(1)

            Exit Function
        End With
    End If
NoPath:
    Set dir = Nothing
    FileNameCreate = FilePathCurrent(Filename)
End Function

Public Function FilePathCurrent(Optional Filename As String) As String
    Dim path As String
    path = ThisWorkbook.path
    If Len(Filename) > 0 Then path = path & "\" & Filename
    FilePathCurrent = Application.Clean(path)
End Function

'=========================================================== Chunks ===========================================================
Public Function ChunksParse(rng As Range, ColNum As Long, Optional ColSize As Long) As Dictionary
    Set ChunksParse = New Dictionary
    Dim FilterRng As Range
    Dim ChunkRng As Range
    Dim Chunk As clsChunk
    Dim lastRow As Long, i As Long

    With rng
        lastRow = .Rows.count + .Row - 1
        'note: counts rm columns from 1
        Set FilterRng = .Columns(ColNum)
        If ColSize > 1 Then Set FilterRng = FilterRng.Resize(ColumnSize:=ColSize)
        With FilterRng.SpecialCells(xlCellTypeConstants, 7)

            For i = .Areas.count To 1 Step -1
                Set ChunkRng = .Areas(i).Resize(lastRow - .Areas(i).Row + 1, rng.Columns.count)
                Set Chunk = New clsChunk
                Chunk.Create ChunkRng
                If ChunksParse.Exists(Chunk.ID) Then GoTo ChunksDuplicated
                ChunksParse.Add Chunk.ID, Chunk
                lastRow = .Areas(i).Row - 1
            Next
        End With
    End With
    Exit Function
ChunksDuplicated:
    Dim oldChunk As clsChunk
    Set oldChunk = ChunksParse.Item(Chunk.ID)
    MsgBox Join(Array("Есть несколько РМ с одним ID", ChunkRng.Address, oldChunk.Range.Address), vbNewLine), Title:="Дублирование РМ"
End Function

'
Public Function ChunksByColumn(Optional ByVal ColNum As Long = 1) As Dictionary
    Dim dataRng As Range
    Set dataRng = RangeMainData
    Set ChunksByColumn = ChunksParse(dataRng, ColNum)
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

'Поворачивает(В некоторых версиях Application.Transpose и .Index не работают)
Public Function ArrayTranspose(arr As Variant) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim j As Long
    
    outputArr = ArrayTryFlat(arr)
    On Error GoTo done
    If UBound(outputArr, 2) <> LBound(outputArr, 2) Then ReDim outputArr(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
    
    For i = LBound(outputArr, 1) To UBound(outputArr, 1)
        For j = LBound(outputArr, 2) To UBound(outputArr, 2)
            outputArr(i, j) = arr(j, i)
        Next
    Next
done:
    ArrayTranspose = outputArr
End Function

'сжимает 2d список
Public Function ArrayTryFlat(arr As Variant) As Variant
    Dim outputArr() As Variant
    Dim i As Long
    On Error GoTo done
    If UBound(arr, 2) - LBound(arr, 2) = 0 Then
        ReDim outputArr(LBound(arr, 1) To UBound(arr, 1))
        For i = LBound(arr, 1) To UBound(arr, 1)
            outputArr(i) = arr(i, 1)
        Next
    ElseIf UBound(arr, 1) - LBound(arr, 1) = 0 Then
        ReDim outputArr(LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(arr, 2) To UBound(arr, 2)
            outputArr(i) = arr(1, i)
        Next
    Else
done:
        ArrayTryFlat = arr
        Exit Function
    End If
    ArrayTryFlat = outputArr
End Function

Public Function Array2DSlice(ByRef arr As Variant, dimention As EDimention, Idx As Long) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim diff As Long
    diff = 1 - LBound(arr, dimention)
    ReDim outputArr(1 To UBound(arr, dimention) + diff)
    For i = 1 To UBound(outputArr)
        If dimention = ERow Then
            outputArr(i) = arr(Idx, i)
        Else
            outputArr(i) = arr(i, Idx)
        End If
    Next
    Array2DSlice = outputArr
End Function

Public Function ArrayUpdateColumns(arr As Variant, ByVal FnDict As Dictionary) As Variant
    Dim outputArr As Variant
    Dim i As Long
    Dim colIdx As Variant
    Dim offset As Long
    Dim done As New Dictionary
    If FnDict.count = 0 Then
        ArrayUpdateColumns = arr
        Exit Function
    End If
    ReDim outputArr(LBound(arr) To UBound(arr))
    For i = LBound(outputArr) To UBound(outputArr)
        If FnDict.Exists(i) And Not done.Exists(i) Then
            outputArr(i + offset) = Application.Run(FnDict(i)("Name"), arr, FnDict(i)("ColNum"), FnDict(i)("Args"))
            If FnDict(i)("NewCol") Then
                offset = offset + 1
                ReDim Preserve outputArr(LBound(arr) To UBound(arr) + offset)
                done.Add i, True
                i = i - 1
            End If
        Else
            On Error GoTo BeforeExit
            outputArr(i + offset) = arr(i)
        End If

    Next
BeforeExit:
    done.RemoveAll
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
    Dim Row As Long
    Dim Title As Variant
    
    For col = 1 To UBound(arr, 2)
        If Not IsEmpty(arr(1, col)) Then Title = arr(1, col)
        For Row = 1 To UBound(arr, 1)
            If IsEmpty(arr(Row, col)) Then GoTo Continue2
            If Not outputDict.Exists(Title) Then
                Set dict = New Dictionary
            Else
                Set dict = outputDict(Title)
            End If
            If Not dict.Exists(arr(Row, col)) Then dict.Add arr(Row, col), col
Continue2:
        Next
        Set outputDict(Title) = dict
    Next
    
    Set Array2DToDictByCol = outputDict
End Function

Public Function ArrayToDict(ByVal Keys As Variant, ByVal ValArr As Variant, Optional SubFunc As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Variant
    Dim i As Long
    If IsNumeric(Keys) Then
        For i = LBound(ValArr) To UBound(ValArr)
            key = Trim(ValArr(i, Keys))
            val = ArrayCutByIndex(ValArr, i)
            If Len(SubFunc) > 0 Then val = Application.Run(SubFunc, val)
            dict(key) = ArrayPush(dict(key), val)
        Next
    ElseIf IsArray(Keys) Then
        If UBound(Keys) - LBound(Keys) <> UBound(ValArr) - LBound(ValArr) Then Err.Raise 13, Description:="ArrayToDict: keys length not equal to values lenght"
        
        For i = LBound(Keys) To UBound(Keys)
            key = Trim(Keys(i))
            
            val = ArrayCutByIndex(ValArr, i)
            If Len(SubFunc) > 0 Then val = Application.Run(SubFunc, val)
            dict(key) = ArrayPush(dict(key), val)
        Next
    End If
    Set ArrayToDict = dict
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
        Array("updateValue", "Фиксированное значение", Array("Значение=Variant", "Менять не пустые=Boolean")), _
        Array("addPrefix", "Добавить префикс", Array("Префикс=String")), _
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
    subDict.Add "ColNum", ColNum
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
' Dim ColNum As Long                            '0
' Dim [ArgName] as [ArgType]                    '[index]
' ArgName = Args(index)
' --- do something
'End Function
''

'Инкрементирует автоматически(статическая функция запоминает значения внутри runtime)
Static Function autoIncrement(ByRef rowArr As Variant, ColNum As Long, args As Variant) As Long
    Dim n As Long
    Dim start As Long                            '0
    If IsArray(args) Then start = args(0)
    If start > n Then n = start
    n = n + 1
    autoIncrement = n
End Function

'Назначаем значение по умолчанию если ячейка пуста
Public Function updateValue(ByRef rowArr As Variant, ColNum As Long, args As Variant) As Variant
    Dim Value As Variant                         '0
    Dim NotEmptyUpdate As Boolean                '1
    ColNum = args(0)
    Value = args(1)
    NotEmptyUpdate = args(2)
    If NotEmptyUpdate = True Or Len(rowArr(ColNum)) = 0 Then
        updateValue = Value
    Else
        updateValue = rowArr(ColNum)
    End If
End Function

'Добавляем префикс с возможным условием
Public Function addPrefix(ByRef rowArr As Variant, ColNum As Long, args As Variant) As String
    Dim Prefix As String                         '0
    On Error GoTo done
    Prefix = args(0)
    If VarType(rowArr(ColNum)) <> vbString Then
        rowArr(ColNum) = CStr(rowArr(ColNum))
    End If
    addPrefix = Prefix & rowArr(ColNum)
done:
End Function

'Проверяем условие для значения и назначаем положительные и отрицательные значения
Public Function checkCondition(ByRef rowArr As Variant, ColNum As Long, args As Variant) As Variant
    Const keyword As String = "target"
    Dim output As Variant
    Dim condition As String                      '0
    Dim FnNameOnTrue As String                   '1
    Dim FnArgs As Variant                        '2
    condition = args(0)
    If UBound(args) > 1 Then
        FnNameOnTrue = args(1)
        FnArgs = args(2)
    End If
    
    condition = "=IF(" & condition & ", True, False)"
    condition = WorksheetFunction.Substitute(condition, keyword, rowArr(ColNum))
    output = Evaluate(condition)
    
    If Len(FnNameOnTrue) > 0 And output = True Then checkCondition = Application.Run(FnNameOnTrue, rowArr, ColNum, FnArgs)
    If Len(FnNameOnTrue) = 0 Then checkCondition = output
End Function

'Заменяет запятую (в выпадающем списке)
Function StringEscComma(ByVal str As String) As String
    StringEscComma = Replace(str, ",", Chr(130))
End Function














