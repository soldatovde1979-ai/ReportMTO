Attribute VB_Name = "modAggregate"
' modAggregate — CORE, generic. Не знает про поля МТО — принимает имена столбцов параметрами.
' Нужен там, где классический PivotTable (без Data Model) не справляется: Distinct Count,
' % от группы, сортировка/топ-N, попарные сравнения между строками. Работает на 2D-массиве,
' считанном из ListObject целиком — быстрее, чем построчный доступ к Range на больших таблицах.
Option Explicit

' Считывает ListObject в 2D-массив (1-based), первая строка — НЕ заголовки (заголовки отдельно).
Public Function ReadTable(lo As ListObject) As Variant
    ReadTable = lo.DataBodyRange.Value2
End Function

Public Function HeaderIndex(lo As ListObject, columnName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If lo.ListColumns(i).Name = columnName Then
            HeaderIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, , "Столбец '" & columnName & "' не найден в " & lo.Name
End Function

' filters: массив строк вида "col=знач" или "col<>знач" (несколько — соединяются через И).
' "col<>" (пустое значение после оператора) означает "непусто".
Private Function RowMatchesFilters(data As Variant, r As Long, lo As ListObject, filters As Variant) As Boolean
    Dim i As Long
    RowMatchesFilters = True
    If Not IsArray(filters) Then Exit Function

    For i = LBound(filters) To UBound(filters)
        Dim f As String: f = filters(i)
        If f = "" Then GoTo NextFilter

        Dim opPos As Long, op As String, col As String, val As String
        opPos = InStr(f, "<>")
        If opPos > 0 Then
            op = "<>": col = Left$(f, opPos - 1): val = Mid$(f, opPos + 2)
        Else
            opPos = InStr(f, "=")
            op = "=": col = Left$(f, opPos - 1): val = Mid$(f, opPos + 1)
        End If

        Dim colIdx As Long: colIdx = HeaderIndex(lo, col)
        Dim cellVal As String: cellVal = CStr(data(r, colIdx))

        If op = "=" Then
            If cellVal <> val Then RowMatchesFilters = False: Exit Function
        Else ' "<>"
            If val = "" Then
                If cellVal = "" Then RowMatchesFilters = False: Exit Function
            Else
                If cellVal = val Then RowMatchesFilters = False: Exit Function
            End If
        End If
NextFilter:
    Next i
End Function

Private Function GroupKey(data As Variant, r As Long, lo As ListObject, groupCols As Variant) As String
    Dim i As Long, parts As String
    For i = LBound(groupCols) To UBound(groupCols)
        parts = parts & CStr(data(r, HeaderIndex(lo, groupCols(i)))) & "|"
    Next i
    GroupKey = parts
End Function

' Возвращает Dictionary: ключ группы ("знач1|знач2|") -> Long (количество строк, прошедших фильтр).
Public Function GroupCount(lo As ListObject, groupCols As Variant, Optional filters As Variant) As Object
    Dim data As Variant: data = ReadTable(lo)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If RowMatchesFilters(data, r, lo, filters) Then
            Dim key As String: key = GroupKey(data, r, lo, groupCols)
            If d.Exists(key) Then d(key) = d(key) + 1 Else d(key) = 1
        End If
    Next r
    Set GroupCount = d
End Function

' Distinct Count: считает уникальные значения distinctCol внутри каждой группы.
Public Function GroupCountDistinct(lo As ListObject, groupCols As Variant, distinctCol As String, Optional filters As Variant) As Object
    Dim data As Variant: data = ReadTable(lo)
    Dim distinctColIdx As Long: distinctColIdx = HeaderIndex(lo, distinctCol)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary") ' "groupKey||value" -> True
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If RowMatchesFilters(data, r, lo, filters) Then
            Dim key As String: key = GroupKey(data, r, lo, groupCols)
            Dim val As String: val = CStr(data(r, distinctColIdx))
            Dim seenKey As String: seenKey = key & "||" & val
            If Not seen.Exists(seenKey) Then
                seen(seenKey) = True
                If d.Exists(key) Then d(key) = d(key) + 1 Else d(key) = 1
            End If
        End If
    Next r
    Set GroupCountDistinct = d
End Function

' Среднее значение valueCol внутри группы (null/пустые ячейки игнорируются).
Public Function GroupAverage(lo As ListObject, groupCols As Variant, valueCol As String, Optional filters As Variant) As Object
    Dim data As Variant: data = ReadTable(lo)
    Dim valueColIdx As Long: valueColIdx = HeaderIndex(lo, valueCol)
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If RowMatchesFilters(data, r, lo, filters) Then
            If IsNumeric(data(r, valueColIdx)) And Not IsEmpty(data(r, valueColIdx)) Then
                Dim key As String: key = GroupKey(data, r, lo, groupCols)
                If sums.Exists(key) Then
                    sums(key) = sums(key) + CDbl(data(r, valueColIdx))
                    counts(key) = counts(key) + 1
                Else
                    sums(key) = CDbl(data(r, valueColIdx))
                    counts(key) = 1
                End If
            End If
        End If
    Next r

    Dim avgs As Object: Set avgs = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In sums.Keys
        avgs(k) = sums(k) / counts(k)
    Next k
    Set GroupAverage = avgs
End Function

' Сортировка Dictionary (ключ->Double/Long) по значению; возвращает массив ключей.
' descending:=True — по убыванию.
Public Function SortDictionaryKeysByValue(d As Object, descending As Boolean) As Variant
    Dim n As Long: n = d.Count
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

    ' Простая сортировка пузырьком — таблицы небольшие (десятки-сотни сотрудников), O(n^2) приемлемо.
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
