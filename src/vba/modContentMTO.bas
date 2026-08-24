Attribute VB_Name = "modContentMTO"
' modContentMTO — CONTENT SPEC (МТО). Реализует 4 функции по контракту modMain.bas (Core):
'   BuildPivots, BuildPrompt, ParseAIResponse, BuildPlaceholders.
' Блоки 1, 4, 5, 9 — через PivotTable (modPivotBuilder), т.к. это простой Count/группировка.
' Блоки 2, 3, 6, 7, 8 — через generic in-memory агрегатор modAggregate (Core), т.к. PivotTable
' без Data Model не умеет Distinct Count, % от группы, сортировку/топ-N и попарные разницы дат.
Option Explicit

Private Const COLOR_BAD As String = "#ef4444"
Private Const COLOR_WARN As String = "#f59e0b"
Private Const COLOR_GOOD As String = "#10b981"

' =====================================================================================
' 1. BuildPivots — Блоки 1, 4, 5, 9 (Pivot); Блоки 2, 3, 6, 7, 8 не строятся здесь —
'    они считаются "на лету" в BuildPlaceholders через modAggregate (не требуют PivotCache).
' =====================================================================================
Public Sub BuildPivots()
    Dim srcRange As Range
    Set srcRange = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").Range

    Dim cache As PivotCache
    Set cache = modPivotBuilder.CreatePivotCache(srcRange)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Pivots")

    ' --- Блок 1: Свод по неделям и постам (Отчёт №4) ---
    modPivotBuilder.BuildPivotTable cache, ws.Range("A1"), "ptBlock1", _
        rowFields:=Array("direction", "arm"), colFields:=Array("week_status"), _
        filterFields:=Array("in_bounds"), _
        dataFieldSpecs:=OneSpec("Key", xlCount, "Кол-во")

    ' --- Блок 4: Динамика % планшет по дирекциям ---
    modPivotBuilder.BuildPivotTable cache, ws.Range("A100"), "ptBlock4", _
        rowFields:=Array("direction", "arm"), colFields:=Array("week_status"), _
        filterFields:=Array("in_bounds"), _
        dataFieldSpecs:=OneSpec("Key", xlCount, "Кол-во")

    ' --- Блок 5: Динамика % планшет по постам ---
    modPivotBuilder.BuildPivotTable cache, ws.Range("A150"), "ptBlock5", _
        rowFields:=Array("postN", "arm"), colFields:=Array("week_status"), _
        filterFields:=Array("in_bounds"), _
        dataFieldSpecs:=OneSpec("Key", xlCount, "Кол-во")

    ' --- Блок 9: Виды дефектов/ремонта ---
    modPivotBuilder.BuildPivotTable cache, ws.Range("A400"), "ptBlock9", _
        rowFields:=Array("zn_type", "defekt_type"), colFields:=Array(), _
        filterFields:=Array(), _
        dataFieldSpecs:=OneSpec("Key", xlCount, "Кол-во")
End Sub

Private Function OneSpec(fieldName As String, fn As XlConsolidationFunction, caption As String) As Collection
    Dim c As New Collection
    c.Add Array(fieldName, fn, caption)
    Set OneSpec = c
End Function

Private Function GetOrCreateSheet(name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = name
        GetOrCreateSheet.Visible = xlSheetVeryHidden
    End If
End Function

Private Function TbData() As ListObject
    Set TbData = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA")
End Function

' =====================================================================================
' 2. BuildPrompt — тело запроса к DeepSeek (data.md §3.1)
' =====================================================================================
Public Function BuildPrompt() As String
    Const SYSTEM_PROMPT As String = _
        "Ты ведущий аналитик данных. Проанализируй предоставленные агрегированные метрики " & _
        "использования планшетов в МТО (доля планшетов по дирекциям, ремзонам и синхронность). " & _
        "Сформируй краткие бизнес-выводы (до 4 предложений на каждый) для 3-х слайдов. Ищи аномалии. " & _
        "Не используй данные, которых нет во входном JSON. Ответ строго в формате JSON: " & _
        "{""slide3_conclusions"": ""..."", ""slide4_conclusions"": ""..."", ""slide5_conclusions"": ""...""}, " & _
        "без markdown-разметки вокруг JSON."

    Dim userMessage As String
    userMessage = "Блок 4 (% планшет по дирекциям по неделям): " & PivotToCompactJSON("ptBlock4") & vbCrLf & _
                  "Блок 5 (% планшет по постам по неделям): " & PivotToCompactJSON("ptBlock5") & vbCrLf & _
                  "Блок 7 (синхронность дирекций по неделям): " & Block7ToCompactText() & vbCrLf & _
                  "Блок 8 (синхронность по неделям и постам): " & Block8ToCompactText() & vbCrLf & _
                  "Блок 9 (виды дефектов): " & PivotToCompactJSON("ptBlock9")
    ' Блок 6 (ФИО) сюда намеренно не включается — контракт Core §8 / политика ПД (data.md §2.1).

    Dim model As String
    model = modMain.GetVariable("AI/MODEL")

    BuildPrompt = "{""model"":""" & JsonEscape(model) & """,""messages"":[" & _
        "{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
        "{""role"":""user"",""content"":""" & JsonEscape(userMessage) & """}]}"
End Function

Private Function PivotToCompactJSON(ptName As String) As String
    On Error GoTo ErrHandler
    Dim pt As PivotTable
    Set pt = ThisWorkbook.Sheets("Pivots").PivotTables(ptName)

    Dim result As String
    Dim r As Range, cell As Range
    Set r = pt.TableRange2
    For Each cell In r.Rows(1).Cells
        result = result & cell.Value & "; "
    Next cell

    PivotToCompactJSON = result
    Exit Function
ErrHandler:
    PivotToCompactJSON = "[недоступно]"
End Function

Private Function JsonEscape(s As String) As String
    JsonEscape = Replace(Replace(Replace(s, "\", "\\"), """", "\"""), vbCrLf, "\n")
End Function

' =====================================================================================
' 3. ParseAIResponse — разбор ответа DeepSeek (data.md §3.2)
' =====================================================================================
Public Function ParseAIResponse(responseText As String, ByRef slide3 As String, ByRef slide4 As String, ByRef slide5 As String) As Boolean
    Const FALLBACK As String = "Внешний ИИ недоступен, показатели см. в таблицах выше."

    slide3 = FALLBACK: slide4 = FALLBACK: slide5 = FALLBACK
    If responseText = "" Then
        ParseAIResponse = False
        Exit Function
    End If

    Dim ok As Boolean
    ok = True

    Dim v As String
    v = ExtractJsonStringValue(responseText, "slide3_conclusions")
    If v <> "" Then slide3 = v Else ok = False

    v = ExtractJsonStringValue(responseText, "slide4_conclusions")
    If v <> "" Then slide4 = v Else ok = False

    v = ExtractJsonStringValue(responseText, "slide5_conclusions")
    If v <> "" Then slide5 = v Else ok = False

    ParseAIResponse = ok
End Function

Private Function ExtractJsonStringValue(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """"
    Dim posKey As Long
    posKey = InStr(json, pattern)
    If posKey = 0 Then Exit Function

    Dim posColon As Long
    posColon = InStr(posKey, json, ":")
    If posColon = 0 Then Exit Function

    Dim posQuoteStart As Long
    posQuoteStart = InStr(posColon, json, """")
    If posQuoteStart = 0 Then Exit Function

    Dim posQuoteEnd As Long
    posQuoteEnd = posQuoteStart + 1
    Do While posQuoteEnd <= Len(json)
        If Mid$(json, posQuoteEnd, 1) = """" And Mid$(json, posQuoteEnd - 1, 1) <> "\" Then Exit Do
        posQuoteEnd = posQuoteEnd + 1
    Loop
    If posQuoteEnd > Len(json) Then Exit Function

    ExtractJsonStringValue = Mid$(json, posQuoteStart + 1, posQuoteEnd - posQuoteStart - 1)
End Function

' =====================================================================================
' 4. BuildPlaceholders — словарь {{ИМЯ}} -> значение для tmp_index.html (data.md §3.3)
' =====================================================================================
Public Function BuildPlaceholders() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d("BLOCK_3_KPI") = BuildBlock3KPI()
    d("BLOCK_1_TABLE") = PivotToHTMLTableWithPercent("ptBlock1", groupCols:=Array("direction"))
    d("BLOCK_4_GAUGE") = PivotToHTMLTableWithPercent("ptBlock4", groupCols:=Array("direction"))
    d("BLOCK_2_TABLE") = BuildBlock2Table()
    d("BLOCK_5_TABLE") = PivotToHTMLTableWithPercent("ptBlock5", groupCols:=Array("postN"))
    d("BLOCK_7_TABLE") = Block7ToCompactText() ' текстовое представление; HTML-таблица — доп. вёрстка при необходимости
    d("BLOCK_8_TABLE") = Block8ToCompactText()
    d("BLOCK_9_TABLE") = PivotToHTMLTable("ptBlock9")

    Dim topN As Long
    topN = CLng(modMain.GetVariable("REPORT/TOPN"))
    d("BLOCK_6_TOP") = Block6RankedTable(topN, descending:=True)
    d("BLOCK_6_BOTTOM") = Block6RankedTable(topN, descending:=False)

    d("AI_INSIGHT_SLIDE_3") = "Внешний ИИ недоступен, показатели см. в таблицах выше."
    d("AI_INSIGHT_SLIDE_4") = "Внешний ИИ недоступен, показатели см. в таблицах выше."
    d("AI_INSIGHT_SLIDE_5") = "Внешний ИИ недоступен, показатели см. в таблицах выше."

    Set BuildPlaceholders = d
End Function

' --- Блок 3: KPI. "Текущая неделя" = максимальный week_status, реально присутствующий в данных
' (а не календарная неделя) — так KPI остаётся согласован с данными даже при задержке выгрузки. ---
Private Function BuildBlock3KPI() As String
    Dim weeksWindow As Long
    weeksWindow = CLng(modMain.GetVariable("REPORT/WEEKS_WINDOW"))

    Dim byWeekDir As Object
    Set byWeekDir = modAggregate.GroupCount(TbData(), Array("week_status", "direction"), _
        filters:=Array("in_bounds=True", "arm<>"))
    Dim byWeekDirTablet As Object
    Set byWeekDirTablet = modAggregate.GroupCount(TbData(), Array("week_status", "direction"), _
        filters:=Array("in_bounds=True", "arm=ПЛАНШЕТ"))

    Dim currentWeek As Long, k As Variant, w As Long
    currentWeek = 0
    For Each k In byWeekDir.Keys
        w = CLng(Split(k, "|")(0))
        If w > currentWeek Then currentWeek = w
    Next k
    Dim previousWeek As Long: previousWeek = currentWeek - 1

    Dim totalZN As Long
    Dim distinctByWeek As Object
    Set distinctByWeek = modAggregate.GroupCountDistinct(TbData(), Array("week_status"), "number", _
        filters:=Array("in_bounds=True"))
    For Each k In distinctByWeek.Keys
        If CLng(k) >= currentWeek - weeksWindow + 1 And CLng(k) <= currentWeek Then
            totalZN = totalZN + distinctByWeek(k)
        End If
    Next k

    Dim pctDGM_cur As Double, pctDENT_cur As Double, pctDGM_prev As Double, pctDENT_prev As Double
    pctDGM_cur = SafePercent(byWeekDirTablet, byWeekDir, currentWeek & "|ДГМ")
    pctDENT_cur = SafePercent(byWeekDirTablet, byWeekDir, currentWeek & "|ДЭНТ")
    pctDGM_prev = SafePercent(byWeekDirTablet, byWeekDir, previousWeek & "|ДГМ")
    pctDENT_prev = SafePercent(byWeekDirTablet, byWeekDir, previousWeek & "|ДЭНТ")

    Dim html As String
    html = "<div class='kpi-grid'>" & _
        "<div class='kpi'>Всего ЗН за " & weeksWindow & " нед.: <b>" & totalZN & "</b></div>" & _
        "<div class='kpi'>ДГМ, % планшет (нед. " & currentWeek & "): <b>" & FormatPct(pctDGM_cur) & _
        "</b> (Δ " & FormatDeltaPP(pctDGM_cur - pctDGM_prev) & ")</div>" & _
        "<div class='kpi'>ДЭНТ, % планшет (нед. " & currentWeek & "): <b>" & FormatPct(pctDENT_cur) & _
        "</b> (Δ " & FormatDeltaPP(pctDENT_cur - pctDENT_prev) & ")</div>" & _
        "</div>"
    BuildBlock3KPI = html
End Function

Private Function SafePercent(numDict As Object, denomDict As Object, key As String) As Double
    If denomDict.Exists(key) Then
        If denomDict(key) > 0 Then
            Dim num As Double
            If numDict.Exists(key) Then num = numDict(key) Else num = 0
            SafePercent = num / denomDict(key)
            Exit Function
        End If
    End If
    SafePercent = 0
End Function

Private Function FormatPct(pct As Double) As String
    FormatPct = Format(pct * 100, "0.0") & "%"
End Function

Private Function FormatDeltaPP(deltaPct As Double) As String
    Dim pp As Double: pp = deltaPct * 100
    If pp >= 0 Then
        FormatDeltaPP = "+" & Format(pp, "0.0") & " п.п."
    Else
        FormatDeltaPP = Format(pp, "0.0") & " п.п."
    End If
End Function

' --- Блок 2: количество ремонтов по постам — честный Distinct Count по number (не прокси). ---
Private Function BuildBlock2Table() As String
    Dim counts As Object
    Set counts = modAggregate.GroupCountDistinct(TbData(), Array("direction", "postN"), "number", _
        filters:=Array("in_bounds=True"))

    Dim html As String
    html = "<table class='block-table'><tr><th>Дирекция</th><th>Пост</th><th>Кол-во ремонтов</th></tr>"

    Dim k As Variant
    For Each k In counts.Keys
        Dim parts() As String: parts = Split(k, "|")
        html = html & "<tr><td>" & parts(0) & "</td><td>" & parts(1) & "</td><td>" & counts(k) & "</td></tr>"
    Next k
    html = html & "</table>"
    BuildBlock2Table = html
End Function

' --- Блок 6: рейтинг инженеров, честная сортировка по % планшет + среднее deltaHours. ---
Private Function Block6RankedTable(topN As Long, descending As Boolean) As String
    Dim totalCounts As Object
    Set totalCounts = modAggregate.GroupCount(TbData(), Array("employee"), _
        filters:=Array("in_bounds=True", "arm<>"))
    Dim tabletCounts As Object
    Set tabletCounts = modAggregate.GroupCount(TbData(), Array("employee"), _
        filters:=Array("in_bounds=True", "arm=ПЛАНШЕТ"))
    Dim avgDelta As Object
    Set avgDelta = modAggregate.GroupAverage(TbData(), Array("employee"), "deltaHours", _
        filters:=Array("ready_for=Готов к выбытию"))

    Dim pctDict As Object: Set pctDict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In totalCounts.Keys
        pctDict(k) = SafePercent(tabletCounts, totalCounts, CStr(k))
    Next k

    Dim sortedKeys As Variant
    sortedKeys = modAggregate.SortDictionaryKeysByValue(pctDict, descending)

    Dim html As String
    html = "<table class='block-table'><tr><th>Инженер</th><th>% планшет</th><th>Кол-во нарядов</th><th>Средняя длительность, ч</th></tr>"

    Dim i As Long, shown As Long
    shown = 0
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        If shown >= topN Then Exit For
        Dim emp As String: emp = CStr(sortedKeys(i))
        Dim deltaVal As String
        If avgDelta.Exists(emp) Then deltaVal = Format(avgDelta(emp), "0.0") Else deltaVal = "н/д"
        html = html & "<tr><td>" & emp & "</td><td>" & FormatPct(pctDict(emp)) & "</td><td>" & _
            totalCounts(emp) & "</td><td>" & deltaVal & "</td></tr>"
        shown = shown + 1
    Next i
    html = html & "</table>"
    Block6RankedTable = html
End Function

' --- Блоки 7/8: попарная разница status_date «Готов к выбытию» между ДГМ и ДЭНТ по одному number.
' postN относится к заказ-наряду целиком (не к дирекции) — у одного number он всегда одинаков
' в обеих строках, поэтому берём его из любой из них (здесь — из строки ДГМ), без допущений. ---
Private Function BuildSyncPairs() As Object
    ' Возвращает Dictionary: number -> Array(weekДГМ, postNДГМ, deltaHours)
    Dim lo As ListObject: Set lo = TbData()
    Dim data As Variant: data = lo.DataBodyRange.Value2

    Dim colNumber As Long: colNumber = modAggregate.HeaderIndex(lo, "number")
    Dim colDirection As Long: colDirection = modAggregate.HeaderIndex(lo, "direction")
    Dim colReadyFor As Long: colReadyFor = modAggregate.HeaderIndex(lo, "ready_for")
    Dim colStatusDate As Long: colStatusDate = modAggregate.HeaderIndex(lo, "status_date")
    Dim colWeek As Long: colWeek = modAggregate.HeaderIndex(lo, "week_status")
    Dim colPostN As Long: colPostN = modAggregate.HeaderIndex(lo, "postN")

    Dim dgmDates As Object: Set dgmDates = CreateObject("Scripting.Dictionary")
    Dim dgmWeeks As Object: Set dgmWeeks = CreateObject("Scripting.Dictionary")
    Dim dgmPosts As Object: Set dgmPosts = CreateObject("Scripting.Dictionary")
    Dim dentDates As Object: Set dentDates = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If CStr(data(r, colReadyFor)) = "Готов к выбытию" Then
            Dim num As String: num = CStr(data(r, colNumber))
            If CStr(data(r, colDirection)) = "ДГМ" Then
                dgmDates(num) = data(r, colStatusDate)
                dgmWeeks(num) = data(r, colWeek)
                dgmPosts(num) = CStr(data(r, colPostN))
            ElseIf CStr(data(r, colDirection)) = "ДЭНТ" Then
                dentDates(num) = data(r, colStatusDate)
            End If
        End If
    Next r

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dgmDates.Keys
        If dentDates.Exists(k) Then
            Dim deltaH As Double
            deltaH = Abs(CDbl(dentDates(k)) - CDbl(dgmDates(k))) * 24
            result(k) = Array(dgmWeeks(k), dgmPosts(k), deltaH)
        End If
    Next k
    Set BuildSyncPairs = result
End Function

Private Function Block7ToCompactText() As String
    Dim pairs As Object: Set pairs = BuildSyncPairs()
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In pairs.Keys
        Dim week As String: week = CStr(pairs(k)(0))
        Dim d As Double: d = pairs(k)(2)
        If sums.Exists(week) Then
            sums(week) = sums(week) + d: counts(week) = counts(week) + 1
        Else
            sums(week) = d: counts(week) = 1
        End If
    Next k

    Dim result As String
    For Each k In sums.Keys
        result = result & "Неделя " & k & ": средняя разница " & Format(sums(k) / counts(k), "0.0") & " ч; "
    Next k
    If result = "" Then result = "[нет пар для сопоставления]"
    Block7ToCompactText = result
End Function

Private Function Block8ToCompactText() As String
    Dim pairs As Object: Set pairs = BuildSyncPairs()
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In pairs.Keys
        Dim key As String: key = CStr(pairs(k)(0)) & "|" & CStr(pairs(k)(1))
        Dim d As Double: d = pairs(k)(2)
        If sums.Exists(key) Then
            sums(key) = sums(key) + d: counts(key) = counts(key) + 1
        Else
            sums(key) = d: counts(key) = 1
        End If
    Next k

    Dim result As String
    For Each k In sums.Keys
        Dim parts() As String: parts = Split(CStr(k), "|")
        result = result & "Неделя " & parts(0) & ", пост " & parts(1) & ": средняя разница " & _
            Format(sums(k) / counts(k), "0.0") & " ч; "
    Next k
    If result = "" Then result = "[нет пар для сопоставления]"
    Block8ToCompactText = result
End Function

' --- Универсальный рендер PivotTable в HTML (Блок 9, без %) ---
Private Function PivotToHTMLTable(ptName As String) As String
    On Error GoTo ErrHandler
    Dim pt As PivotTable
    Set pt = ThisWorkbook.Sheets("Pivots").PivotTables(ptName)

    Dim html As String
    html = "<table class='block-table'>"
    Dim r As Range, row As Range, cell As Range
    Set r = pt.TableRange2
    For Each row In r.Rows
        html = html & "<tr>"
        For Each cell In row.Cells
            html = html & "<td>" & cell.Value & "</td>"
        Next cell
        html = html & "</tr>"
    Next row
    html = html & "</table>"

    PivotToHTMLTable = html
    Exit Function
ErrHandler:
    PivotToHTMLTable = "<!-- " & ptName & " недоступен -->"
End Function

' --- Рендер PivotTable в HTML + добавленная строка "% планшет" на группу (Блоки 1, 4, 5).
' % считается напрямую по tbDATA через modAggregate (не из уже построенного Pivot, там только Count),
' раскраска ячеек — через modColor.PercentToColor. ---
Private Function PivotToHTMLTableWithPercent(ptName As String, groupCols As Variant) As String
    Dim baseTable As String
    baseTable = PivotToHTMLTable(ptName)

    Dim gcArr() As String
    ReDim gcArr(0 To UBound(groupCols))
    Dim i As Long
    For i = 0 To UBound(groupCols): gcArr(i) = CStr(groupCols(i)): Next i

    Dim total As Object: Set total = modAggregate.GroupCount(TbData(), gcArr, filters:=Array("in_bounds=True", "arm<>"))
    Dim tablet As Object: Set tablet = modAggregate.GroupCount(TbData(), gcArr, filters:=Array("in_bounds=True", "arm=ПЛАНШЕТ"))

    Dim pctRow As String
    pctRow = "<tr class='pct-row'><td colspan='" & (UBound(groupCols) + 1) & "'>% планшет</td>"
    Dim k As Variant
    For Each k In total.Keys
        Dim pct As Double: pct = SafePercent(tablet, total, CStr(k))
        Dim color As String: color = modColor.PercentToColor(pct, COLOR_BAD, COLOR_WARN, COLOR_GOOD)
        pctRow = pctRow & "<td style='background:" & color & ";color:#fff;'>" & FormatPct(pct) & "</td>"
    Next k
    pctRow = pctRow & "</tr>"

    PivotToHTMLTableWithPercent = baseTable & "<table class='block-table pct-table'>" & pctRow & "</table>"
End Function
