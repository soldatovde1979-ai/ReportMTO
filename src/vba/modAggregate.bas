Attribute VB_Name = "modAggregate"
' modAggregate - CORE, generic. Не знает про поля направления - принимает имена столбцов параметрами.
' Нужен там, где классический PivotTable (без Data Model) не справляется: Distinct Count,
' % от группы, сортировка/топ-N, попарные сравнения между строками.
'
' v3.1 (правки по ревью 24.08.2026):
'   P0-1  - все массивы принимаются как Variant (VBA не допускает Optional-массивы и требует
'           точного совпадения типа у параметра-массива; Array(...) - всегда Variant).
'   P2-1  - карта "имя столбца -> индекс" строится один раз в BeginSnapshot, а не линейным
'           обходом ListColumns через COM внутри цикла по строкам.
'   P2-2  - данные читаются в память ОДИН раз на прогон (BeginSnapshot), а не в каждом GroupCount.
'   P1-13 - пустая tbDATA (DataBodyRange = Nothing) не роняет модуль: RowCount = 0.
'
' КОНТРАКТ ИСПОЛЬЗОВАНИЯ (изменён в v3.1):
'   modAggregate.BeginSnapshot lo      ' один раз перед серией агрегаций
'   ... GroupCount / GroupCountDistinct / GroupAverage / DistinctValues ...
'   modAggregate.EndSnapshot           ' освободить память (необязательно)
' Функции агрегации больше НЕ принимают ListObject - они работают по активному снимку.
Option Explicit

Private mLo As ListObject
Private mData As Variant
Private mCols As Object          ' имя столбца -> индекс (1-based), TextCompare
Private mRows As Long
Private mReady As Boolean

' =====================================================================================
' Снимок данных
' =====================================================================================
Public Sub BeginSnapshot(lo As ListObject)
    Set mLo = lo

    Set mCols = CreateObject("Scripting.Dictionary")
    mCols.CompareMode = 1 ' vbTextCompare

    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        mCols(lo.ListColumns(i).Name) = i
    Next i

    If lo.DataBodyRange Is Nothing Then
        mData = Empty
        mRows = 0
    Else
        mData = lo.DataBodyRange.Value2
        mRows = UBound(mData, 1)
    End If

    mReady = True
End Sub

Public Sub EndSnapshot()
    mData = Empty
    Set mCols = Nothing
    Set mLo = Nothing
    mRows = 0
    mReady = False
End Sub

Public Function IsReady() As Boolean
    IsReady = mReady
End Function

Public Function RowCount() As Long
    RowCount = mRows
End Function

Public Function ColIndex(columnName As String) As Long
    If Not mReady Then Err.Raise vbObjectError + 3, , "modAggregate: снимок не создан (BeginSnapshot)"
    If Not mCols.Exists(columnName) Then
        Err.Raise vbObjectError + 2, , "Столбец '" & columnName & "' не найден в " & mLo.Name
    End If
    ColIndex = mCols(columnName)
End Function

Public Function HasColumn(columnName As String) As Boolean
    HasColumn = mReady And Not (mCols Is Nothing)
    If HasColumn Then HasColumn = mCols.Exists(columnName)
End Function

' Значение ячейки снимка как нормализованный текст (см. NormText).
Public Function CellText(r As Long, columnName As String) As String
    CellText = NormText(mData(r, ColIndex(columnName)))
End Function

Public Function CellRaw(r As Long, columnName As String) As Variant
    CellRaw = mData(r, ColIndex(columnName))
End Function

' Нормализация значения ячейки к тексту.
' Boolean приводится к "True"/"False" независимо от локали Excel - иначе фильтр
' "in_bounds=True" ломается на русской локали, где ячейка отображается как ИСТИНА.
Private Function NormText(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        NormText = ""
    ElseIf VarType(v) = vbBoolean Then
        NormText = IIf(v, "True", "False")
    ElseIf IsError(v) Then
        NormText = ""
    Else
        NormText = Trim$(CStr(v))
    End If
End Function

' Приводит распространённые написания булева значения к "True"/"False",
' чтобы фильтр не зависел от того, как значение попало на лист.
Private Function NormBoolLiteral(s As String) As String
    Select Case UCase$(s)
        Case "TRUE", "ИСТИНА", "1", "ДА": NormBoolLiteral = "True"
        Case "FALSE", "ЛОЖЬ", "0", "НЕТ": NormBoolLiteral = "False"
        Case Else: NormBoolLiteral = s
    End Select
End Function

' =====================================================================================
' Фильтры
' Формат элемента массива filters:
'   "col=знач"        - равно
'   "col<>знач"       - не равно
'   "col<>"           - непусто
'   "col@=a;b;c"      - значение входит в список (разделитель ";")
' Несколько элементов соединяются через И. filters можно не передавать.
' =====================================================================================
Private Function RowMatchesFilters(r As Long, filters As Variant) As Boolean
    RowMatchesFilters = True

    If IsMissing(filters) Then Exit Function
    If IsEmpty(filters) Then Exit Function
    If Not IsArray(filters) Then Exit Function
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(filters): ub = UBound(filters)
    If Err.Number <> 0 Then Err.Clear: Exit Function ' нераспределённый массив
    On Error GoTo 0
    If ub < lb Then Exit Function

    Dim i As Long
    For i = lb To ub
        Dim f As String: f = CStr(filters(i))
        If Len(f) > 0 Then
            Dim opPos As Long, op As String, col As String, val As String
            opPos = InStr(f, "@=")
            If opPos > 0 Then
                op = "@=": col = Left$(f, opPos - 1): val = Mid$(f, opPos + 2)
            Else
                opPos = InStr(f, "<>")
                If opPos > 0 Then
                    op = "<>": col = Left$(f, opPos - 1): val = Mid$(f, opPos + 2)
                Else
                    opPos = InStr(f, "=")
                    If opPos = 0 Then Err.Raise vbObjectError + 4, , "Неверный фильтр: " & f
                    op = "=": col = Left$(f, opPos - 1): val = Mid$(f, opPos + 1)
                End If
            End If

            Dim cellVal As String
            cellVal = NormText(mData(r, ColIndex(col)))
            val = NormBoolLiteral(val)
            cellVal = NormBoolLiteral(cellVal)

            Select Case op
                Case "="
                    If StrComp(cellVal, val, vbTextCompare) <> 0 Then RowMatchesFilters = False: Exit Function
                Case "<>"
                    If val = "" Then
                        If cellVal = "" Then RowMatchesFilters = False: Exit Function
                    Else
                        If StrComp(cellVal, val, vbTextCompare) = 0 Then RowMatchesFilters = False: Exit Function
                    End If
                Case "@="
                    Dim parts() As String, j As Long, hit As Boolean
                    parts = Split(val, ";")
                    hit = False
                    For j = LBound(parts) To UBound(parts)
                        If StrComp(cellVal, Trim$(parts(j)), vbTextCompare) = 0 Then hit = True: Exit For
                    Next j
                    If Not hit Then RowMatchesFilters = False: Exit Function
            End Select
        End If
    Next i
End Function

Private Function GroupKey(r As Long, groupIdx() As Long) As String
    Dim i As Long, parts As String
    For i = LBound(groupIdx) To UBound(groupIdx)
        parts = parts & NormText(mData(r, groupIdx(i))) & "|"
    Next i
    GroupKey = parts
End Function

' Предрасчёт индексов колонок группировки (один раз на вызов, не на строку).
Private Function ColIndexes(cols As Variant) As Long()
    Dim idx() As Long
    Dim n As Long: n = UBound(cols) - LBound(cols) + 1
    ReDim idx(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        idx(i) = ColIndex(CStr(cols(LBound(cols) + i)))
    Next i
    ColIndexes = idx
End Function

Private Sub RequireSnapshot()
    If Not mReady Then Err.Raise vbObjectError + 3, , "modAggregate: снимок не создан (BeginSnapshot)"
End Sub

' =====================================================================================
' Агрегации. Все возвращают Scripting.Dictionary: "знач1|знач2|" -> число.
' =====================================================================================
Public Function GroupCount(groupCols As Variant, Optional filters As Variant) As Object
    RequireSnapshot
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set GroupCount = d
    If mRows = 0 Then Exit Function

    Dim gi() As Long: gi = ColIndexes(groupCols)
    Dim r As Long, key As String
    For r = 1 To mRows
        If RowMatchesFilters(r, filters) Then
            key = GroupKey(r, gi)
            If d.Exists(key) Then d(key) = d(key) + 1 Else d(key) = 1
        End If
    Next r
End Function

Public Function GroupCountDistinct(groupCols As Variant, distinctCol As String, Optional filters As Variant) As Object
    RequireSnapshot
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set GroupCountDistinct = d
    If mRows = 0 Then Exit Function

    Dim gi() As Long: gi = ColIndexes(groupCols)
    Dim dci As Long: dci = ColIndex(distinctCol)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Dim r As Long, key As String, seenKey As String
    For r = 1 To mRows
        If RowMatchesFilters(r, filters) Then
            key = GroupKey(r, gi)
            seenKey = key & "||" & NormText(mData(r, dci))
            If Not seen.Exists(seenKey) Then
                seen(seenKey) = True
                If d.Exists(key) Then d(key) = d(key) + 1 Else d(key) = 1
            End If
        End If
    Next r
End Function

Public Function GroupAverage(groupCols As Variant, valueCol As String, Optional filters As Variant) As Object
    RequireSnapshot
    Dim avgs As Object: Set avgs = CreateObject("Scripting.Dictionary")
    Set GroupAverage = avgs
    If mRows = 0 Then Exit Function

    Dim gi() As Long: gi = ColIndexes(groupCols)
    Dim vci As Long: vci = ColIndex(valueCol)
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")

    Dim r As Long, key As String, v As Variant
    For r = 1 To mRows
        If RowMatchesFilters(r, filters) Then
            v = mData(r, vci)
            If Not IsEmpty(v) Then
                If Not IsError(v) Then
                    If IsNumeric(v) Then
                        key = GroupKey(r, gi)
                        If sums.Exists(key) Then
                            sums(key) = sums(key) + CDbl(v)
                            counts(key) = counts(key) + 1
                        Else
                            sums(key) = CDbl(v)
                            counts(key) = 1
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Dim k As Variant
    For Each k In sums.Keys
        avgs(k) = sums(k) / counts(k)
    Next k
End Function

' Уникальные значения столбца (с учётом фильтров): значение -> количество строк.
Public Function DistinctValues(columnName As String, Optional filters As Variant) As Object
    Set DistinctValues = GroupCount(Array(columnName), filters)
End Function

' =====================================================================================
' Сортировки
' =====================================================================================
' Сортировка Dictionary (ключ -> Double/Long) по значению; возвращает массив ключей.
' P1-13: при пустом словаре возвращается пустой Array(), а не ошибка 9.
Public Function SortDictionaryKeysByValue(d As Object, descending As Boolean) As Variant
    If d Is Nothing Then SortDictionaryKeysByValue = Array(): Exit Function
    Dim n As Long: n = d.Count
    If n = 0 Then SortDictionaryKeysByValue = Array(): Exit Function

    Dim keys() As Variant, vals() As Double
    ReDim keys(0 To n - 1)
    ReDim vals(0 To n - 1)

    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In d.Keys
        keys(i) = k
        vals(i) = CDbl(d(k))
        i = i + 1
    Next k

    ' Сортировка пузырьком - таблицы небольшие (десятки-сотни групп), O(n^2) приемлемо.
    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            Dim swapNeeded As Boolean
            If descending Then
                swapNeeded = (vals(b) < vals(b + 1))
            Else
                swapNeeded = (vals(b) > vals(b + 1))
            End If
            If swapNeeded Then
                Dim tv As Double: tv = vals(b): vals(b) = vals(b + 1): vals(b + 1) = tv
                Dim tk As Variant: tk = keys(b): keys(b) = keys(b + 1): keys(b + 1) = tk
            End If
        Next b
    Next a

    SortDictionaryKeysByValue = keys
End Function

' Сортировка ключей словаря по самому ключу: numeric:=True - как числа, иначе как текст.
' Нужна для осей матриц (недели по возрастанию, посты по алфавиту) - порядок ключей
' Scripting.Dictionary равен порядку появления в данных и на роль оси не годится.
Public Function SortKeys(d As Object, numeric As Boolean) As Variant
    If d Is Nothing Then SortKeys = Array(): Exit Function
    Dim n As Long: n = d.Count
    If n = 0 Then SortKeys = Array(): Exit Function

    Dim keys() As Variant
    ReDim keys(0 To n - 1)
    Dim i As Long: i = 0
    Dim k As Variant
    For Each k In d.Keys
        keys(i) = k: i = i + 1
    Next k

    Dim a As Long, b As Long
    For a = 0 To n - 2
        For b = 0 To n - 2 - a
            Dim swapNeeded As Boolean
            If numeric Then
                swapNeeded = (SafeDbl(keys(b)) > SafeDbl(keys(b + 1)))
            Else
                swapNeeded = (StrComp(CStr(keys(b)), CStr(keys(b + 1)), vbTextCompare) > 0)
            End If
            If swapNeeded Then
                Dim tk As Variant: tk = keys(b): keys(b) = keys(b + 1): keys(b + 1) = tk
            End If
        Next b
    Next a

    SortKeys = keys
End Function

' Числовое значение ключа. Ключи составных группировок имеют вид "202643|ДГМ|",
' поэтому берётся первая часть до разделителя - иначе IsNumeric даёт False и числовая
' сортировка осей (недели!) молча вырождается в порядок появления в данных.
Private Function SafeDbl(v As Variant) As Double
    Dim s As String
    s = Split(CStr(v), "|")(0)
    SafeDbl = Val(s)
End Function
