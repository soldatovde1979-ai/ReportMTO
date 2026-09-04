Attribute VB_Name = "modContentMTO"
' modContentMTO - CONTENT SPEC (МТО). Реализует 4 функции по контракту modMain.bas (Core):
'   BuildPivots, BuildPrompt, ParseAIResponse, BuildPlaceholders(s3, s4, s5).
'
' v3.1 - переработан по итогам ревью 24.08.2026. Ключевые изменения:
'   P1-1/2/3 - Блоки 1, 4, 5, 9 БОЛЬШЕ НЕ СТРОЯТСЯ ЧЕРЕЗ PivotTable. Причина: Pivot без Data
'              Model не фильтрует поле в области страницы (оставался «(Все)»), строка «% планшет»
'              физически не могла совпасть со столбцами сводной, а Блоки 4/5 были копией Блока 1
'              (Count вместо %). Всё считается через modAggregate + generic-рендер матрицы.
'              modPivotBuilder.bas остаётся в Core как есть, направлением МТО просто не используется.
'   P1-4     - группировка/сортировка недель по yearWeek (year_status*100 + week_status),
'              отображение - номер недели; отчёт больше не ломается на границе года.
'   P0-5     - в промпт уходит полная матрица агрегатов в виде настоящего JSON, по белому списку
'              полей (раньше уходила первая строка TableRange2, т.е. строка фильтра «(Все)»).
'   P0-4     - двухшаговый разбор ответа: сначала content внешнего JSON, затем JsonUnescape,
'              затем три ключа внутри развёрнутого текста.
'   P0-2     - BuildPlaceholders принимает выводы ИИ параметрами.
'   P1-8     - все значения проходят через modHTMLEngine.HtmlEscape.
'   P1-12    - антитоп Блока 6 отсекает сотрудников с числом записей < REPORT/MIN_RECORDS.
'   P1-14    - avgDelta в Блоке 6 фильтруется согласованно с соседними агрегатами.
'   P1-15    - полный JsonEscape (управляющие символы, \n, \r, \t).
'   P2-3     - BuildSyncPairs кэшируется на прогон (был 4 полных прохода).
'
' (!) 24.08.2026: поле arm принимает не два, а ТРИ значения - "ПК", "ПЛАНШЕТ" и "НЕ ПОДПИСАНО"
' (статус смены не подписан). Такие строки временно исключаются из всех блоков по решению
' владельца процесса. Реализовано белым списком (arm@=ПК;ПЛАНШЕТ), а не отсечением пустых
' значений ("arm<>"): "НЕ ПОДПИСАНО" - непустое значение и через прежний фильтр проходило,
' завышая знаменатель "% планшет" и счётчики Блоков 2/9.
' Если заказчик решит показывать неподписанные - менять только три функции FBase/FArm/FTablet.
Option Explicit

Private Const COLOR_BAD As String = "#ef4444"
Private Const COLOR_WARN As String = "#f59e0b"
Private Const COLOR_GOOD As String = "#10b981"

Private Const AI_FALLBACK As String = "Внешний ИИ недоступен, показатели см. в таблицах выше."

Private mPairs As Object          ' кэш BuildSyncPairs на один прогон (P2-3)
Private mPairsReady As Boolean
Private mMultiYear As Integer     ' -1 не определено, 0 один год, 1 несколько лет

' =====================================================================================
' 0. Инфраструктура
' =====================================================================================
Private Function TbData() As ListObject
    Set TbData = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA")
End Function

' Контракт Core §16. Начиная с v3.1 процедура не строит PivotTable - она готовит
' снимок данных для modAggregate и сбрасывает кэши блоков. Имя сохранено, т.к.
' зафиксировано контрактом Core (см. §16 и A-2 ревью).
Public Sub BuildPivots()
    mPairsReady = False
    Set mPairs = Nothing
    mMultiYear = -1

    modAggregate.EndSnapshot
    modAggregate.BeginSnapshot TbData()
End Sub

Private Sub EnsureSnapshot()
    If Not modAggregate.IsReady Then
        modAggregate.BeginSnapshot TbData()
        mMultiYear = -1
        mPairsReady = False
    End If
End Sub

' --- Наборы фильтров (Content Spec §5) ---
' Базовый фильтр направления: in_bounds = ИСТИНА И arm из {ПК, ПЛАНШЕТ}.
' (!) 24.08.2026: поле arm принимает не два, а ТРИ значения - "ПК", "ПЛАНШЕТ" и "НЕ ПОДПИСАНО"
' (статус смены не подписан). Такие строки временно исключаются из всех блоков по решению
' владельца процесса. Реализовано белым списком (arm@=ПК;ПЛАНШЕТ), а не отсечением пустых
' значений ("arm<>"): "НЕ ПОДПИСАНО" - непустое значение и через прежний фильтр проходило,
' завышая знаменатель "% планшет" и счётчики Блоков 2/9.
' Если заказчик решит показывать неподписанные - менять только эти три функции.
Private Function FBase() As Variant
    FBase = Array("in_bounds=True", "arm@=ПК;ПЛАНШЕТ")
End Function

Private Function FArm() As Variant
    FArm = FBase()
End Function

Private Function FTablet() As Variant
    FTablet = Array("in_bounds=True", "arm=ПЛАНШЕТ")
End Function

' =====================================================================================
' 1. Работа с осями матриц (недели/строки)
' =====================================================================================
' Определяет, присутствует ли в данных больше одного года - от этого зависит подпись недели.
Private Function IsMultiYear() As Boolean
    If mMultiYear <> -1 Then IsMultiYear = (mMultiYear = 1): Exit Function

    EnsureSnapshot
    Dim years As Object
    Set years = modAggregate.DistinctValues("year_status", FBase())
    mMultiYear = IIf(years.Count > 1, 1, 0)
    IsMultiYear = (mMultiYear = 1)
End Function

' yearWeek (202643) -> подпись столбца ("43" либо "43/2026").
Private Function WeekLabel(yw As Variant) As String
    Dim n As Long
    If Not IsNumeric(yw) Then WeekLabel = CStr(yw): Exit Function
    n = CLng(yw)
    If IsMultiYear() Then
        WeekLabel = CStr(n Mod 100) & "/" & CStr(n \ 100)
    Else
        WeekLabel = CStr(n Mod 100)
    End If
End Function

' Извлекает часть составного ключа modAggregate ("a|b|c|") по индексу (0-based).
Private Function KeyPart(k As Variant, idx As Long) As String
    Dim parts() As String
    parts = Split(CStr(k), "|")
    If idx <= UBound(parts) Then KeyPart = parts(idx) Else KeyPart = ""
End Function

' Уникальные значения части составного ключа, отсортированные.
Private Function AxisFromKeys(d As Object, partIdx As Long, numeric As Boolean) As Variant
    Dim ax As Object: Set ax = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In d.Keys
        ax(KeyPart(k, partIdx)) = 1
    Next k
    AxisFromKeys = modAggregate.SortKeys(ax, numeric)
End Function

Private Function DictVal(d As Object, key As String) As Double
    If d.Exists(key) Then DictVal = CDbl(d(key)) Else DictVal = 0
End Function

Private Function SafePercent(numDict As Object, denomDict As Object, key As String) As Double
    If denomDict.Exists(key) Then
        If denomDict(key) > 0 Then
            SafePercent = DictVal(numDict, key) / CDbl(denomDict(key))
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

Private Function Esc(s As Variant) As String
    Esc = modHTMLEngine.HtmlEscape(CStr(s))
End Function

Private Function PctCell(pct As Double) As String
    Dim color As String
    color = modColor.PercentToColor(pct, COLOR_BAD, COLOR_WARN, COLOR_GOOD)
    PctCell = "<td class='pct' style='background:" & color & ";color:#fff;'>" & FormatPct(pct) & "</td>"
End Function

' =====================================================================================
' 2. BuildPrompt - тело запроса к DeepSeek (data.md §3.1)
'    В промпт уходят ТОЛЬКО агрегаты по белому списку полей (Content Spec §8):
'    direction, postN, неделя, счётчики, доли, zn_type, defekt_type, часы расхождения.
'    employee и defect_desc не сериализуются ни в каком виде.
' =====================================================================================
Public Function BuildPrompt() As String
    EnsureSnapshot

    Const SYSTEM_PROMPT As String = _
        "Ты ведущий аналитик данных. Проанализируй предоставленные агрегированные метрики " & _
        "использования планшетов в МТО (доля планшетов по дирекциям, ремзонам и синхронность). " & _
        "Сформируй краткие бизнес-выводы (до 4 предложений на каждый) для 3-х слайдов. Ищи аномалии. " & _
        "Не используй данные, которых нет во входном JSON. Ответ строго в формате JSON: " & _
        "{""slide3_conclusions"": ""..."", ""slide4_conclusions"": ""..."", ""slide5_conclusions"": ""...""}, " & _
        "без markdown-разметки вокруг JSON."

    Dim userMessage As String
    userMessage = "{""block4_percent_by_direction"":" & PctMatrixToJson("direction") & _
                  ",""block5_percent_by_post"":" & PctMatrixToJson("postN") & _
                  ",""block7_sync_by_week"":" & SyncToJson(False) & _
                  ",""block8_sync_by_week_post"":" & SyncToJson(True) & _
                  ",""block9_defect_types"":" & Block9ToJson() & "}"
    ' Блок 6 (ФИО) сюда намеренно не включается - контракт Core §8 / политика ПД (data.md §2.1).

    Dim model As String
    model = modMain.GetVariable("AI/MODEL")

    ' response_format/temperature - контентные настройки запроса (P0-4): просим провайдера
    ' гарантировать JSON-объект и снижаем вариативность формулировок.
    BuildPrompt = "{""model"":""" & JsonEscape(model) & """," & _
        """temperature"":0.2," & _
        """response_format"":{""type"":""json_object""}," & _
        """messages"":[" & _
        "{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
        "{""role"":""user"",""content"":""" & JsonEscape(userMessage) & """}]}"
End Function

' Матрица «строка x неделя» с долей планшета -> JSON-массив записей.
Private Function PctMatrixToJson(rowCol As String) As String
    EnsureSnapshot

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FArm())
    Set tablet = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FTablet())

    Dim s As String: s = "["
    Dim first As Boolean: first = True
    Dim k As Variant
    For Each k In total.Keys
        If Not first Then s = s & ","
        first = False
        s = s & "{""row"":""" & JsonEscape(KeyPart(k, 0)) & """," & _
                """week"":""" & JsonEscape(WeekLabel(KeyPart(k, 1))) & """," & _
                """total"":" & CStr(DictVal(total, CStr(k))) & "," & _
                """tablet"":" & CStr(DictVal(tablet, CStr(k))) & "," & _
                """pct"":" & Format(SafePercent(tablet, total, CStr(k)), "0.000") & "}"
    Next k
    PctMatrixToJson = s & "]"
End Function

Private Function SyncToJson(byPost As Boolean) As String
    Dim agg As Object
    Set agg = SyncAggregate(byPost)

    Dim s As String: s = "["
    Dim first As Boolean: first = True
    Dim k As Variant
    For Each k In agg.Keys
        If Not first Then s = s & ","
        first = False
        Dim v As Variant: v = agg(k) ' Array(avgHours, pairsCount)
        s = s & "{""week"":""" & JsonEscape(WeekLabel(KeyPart(k, 0))) & """"
        If byPost Then s = s & ",""post"":""" & JsonEscape(KeyPart(k, 1)) & """"
        s = s & ",""avg_hours"":" & Format(v(0), "0.00") & ",""pairs"":" & CStr(v(1)) & "}"
    Next k
    SyncToJson = s & "]"
End Function

Private Function Block9ToJson() As String
    EnsureSnapshot
    Dim counts As Object
    Set counts = modAggregate.GroupCount(Array("zn_type", "defekt_type"), FBase())

    Dim s As String: s = "["
    Dim first As Boolean: first = True
    Dim k As Variant
    For Each k In counts.Keys
        If Not first Then s = s & ","
        first = False
        s = s & "{""zn_type"":""" & JsonEscape(KeyPart(k, 0)) & """," & _
                """defekt_type"":""" & JsonEscape(KeyPart(k, 1)) & """," & _
                """count"":" & CStr(counts(k)) & "}"
    Next k
    Block9ToJson = s & "]"
End Function

' Полное JSON-экранирование (P1-15): обратный слэш, кавычка, все управляющие символы < 0x20.
Private Function JsonEscape(s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8:  out = out & "\b"
            Case 9:  out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right$("000" & Hex$(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsonEscape = out
End Function

' =====================================================================================
' 3. ParseAIResponse - двухшаговый разбор ответа Chat Completions (P0-4)
'    Уровень 1: {"choices":[{"message":{"content":"<строка с экранированным JSON>"}}]}
'    Уровень 2: развёрнутый content -> три ключа slideN_conclusions.
' =====================================================================================
Public Function ParseAIResponse(responseText As String, ByRef slide3 As String, ByRef slide4 As String, ByRef slide5 As String) As Boolean
    slide3 = AI_FALLBACK: slide4 = AI_FALLBACK: slide5 = AI_FALLBACK

    If Trim$(responseText) = "" Then
        ParseAIResponse = False
        Exit Function
    End If

    Dim payload As String
    payload = ExtractJsonStringValue(responseText, "content")

    If payload = "" Then
        ' Провайдер вернул целевой JSON верхним уровнем (или иной формат) - пробуем как есть.
        payload = responseText
    Else
        payload = JsonUnescape(payload)
    End If

    payload = StripMarkdownFence(payload)

    Dim ok As Boolean: ok = True
    Dim v As String

    v = JsonUnescape(ExtractJsonStringValue(payload, "slide3_conclusions"))
    If v <> "" Then slide3 = v Else ok = False

    v = JsonUnescape(ExtractJsonStringValue(payload, "slide4_conclusions"))
    If v <> "" Then slide4 = v Else ok = False

    v = JsonUnescape(ExtractJsonStringValue(payload, "slide5_conclusions"))
    If v <> "" Then slide5 = v Else ok = False

    ParseAIResponse = ok
End Function

' Снимает обёртку ```json ... ``` , если модель всё-таки её добавила.
Private Function StripMarkdownFence(s As String) As String
    Dim t As String: t = Trim$(s)
    If Left$(t, 3) = "```" Then
        Dim p As Long
        p = InStr(t, vbLf)
        If p > 0 Then t = Mid$(t, p + 1)
        p = InStrRev(t, "```")
        If p > 0 Then t = Left$(t, p - 1)
    End If
    StripMarkdownFence = Trim$(t)
End Function

Private Function ExtractJsonStringValue(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """"
    Dim posKey As Long
    posKey = InStr(json, pattern)
    If posKey = 0 Then Exit Function

    Dim posColon As Long
    posColon = InStr(posKey + Len(pattern), json, ":")
    If posColon = 0 Then Exit Function

    Dim posQuoteStart As Long
    posQuoteStart = InStr(posColon, json, """")
    If posQuoteStart = 0 Then Exit Function

    Dim posQuoteEnd As Long, backslashes As Long, j As Long
    posQuoteEnd = posQuoteStart + 1
    Do While posQuoteEnd <= Len(json)
        If Mid$(json, posQuoteEnd, 1) = """" Then
            ' Кавычка закрывающая, если перед ней ЧЁТНОЕ число обратных слэшей.
            backslashes = 0
            j = posQuoteEnd - 1
            Do While j >= 1
                If Mid$(json, j, 1) = "\" Then
                    backslashes = backslashes + 1
                    j = j - 1
                Else
                    Exit Do
                End If
            Loop
            If backslashes Mod 2 = 0 Then Exit Do
        End If
        posQuoteEnd = posQuoteEnd + 1
    Loop
    If posQuoteEnd > Len(json) Then Exit Function

    ExtractJsonStringValue = Mid$(json, posQuoteStart + 1, posQuoteEnd - posQuoteStart - 1)
End Function

' Разворачивает JSON-экранирование строки: \" \\ \/ \n \r \t \b \f \uXXXX.
Private Function JsonUnescape(s As String) As String
    If InStr(s, "\") = 0 Then JsonUnescape = s: Exit Function

    Dim i As Long, out As String, ch As String, nx As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case """": out = out & """": i = i + 2
                Case "\":  out = out & "\":  i = i + 2
                Case "/":  out = out & "/":  i = i + 2
                Case "n":  out = out & vbLf: i = i + 2
                Case "r":  out = out & vbCr: i = i + 2
                Case "t":  out = out & vbTab: i = i + 2
                Case "b":  out = out & Chr$(8): i = i + 2
                Case "f":  out = out & Chr$(12): i = i + 2
                Case "u"
                    If i + 5 <= Len(s) Then
                        out = out & ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
                        i = i + 6
                    Else
                        out = out & ch: i = i + 1
                    End If
                Case Else
                    out = out & nx: i = i + 2
            End Select
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    JsonUnescape = out
End Function

' =====================================================================================
' 4. BuildPlaceholders - словарь {{ИМЯ}} -> значение для tmp_index.html (data.md §3.3)
'    Сигнатура расширена (P0-2 / A-2): выводы ИИ приходят параметрами.
' =====================================================================================
Public Function BuildPlaceholders(aiSlide3 As String, aiSlide4 As String, aiSlide5 As String) As Object
    EnsureSnapshot

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d("BLOCK_3_KPI") = BuildBlock3KPI()
    d("BLOCK_1_TABLE") = BuildBlock1Table()
    d("BLOCK_4_GAUGE") = BuildPctMatrixTable("direction", "Дирекция")
    d("BLOCK_2_TABLE") = BuildBlock2Table()
    d("BLOCK_5_TABLE") = BuildPctMatrixTable("postN", "Пост")
    d("BLOCK_7_TABLE") = BuildSyncTable(False)
    d("BLOCK_8_TABLE") = BuildSyncTable(True)
    d("BLOCK_9_TABLE") = BuildBlock9Table()

    Dim topN As Long
    topN = CLng(modMain.GetVariableDef("REPORT/TOPN", "10"))
    d("BLOCK_6_TOP") = Block6RankedTable(topN, descending:=True)
    d("BLOCK_6_BOTTOM") = Block6RankedTable(topN, descending:=False)

    d("AI_INSIGHT_SLIDE_3") = Esc(aiSlide3)
    d("AI_INSIGHT_SLIDE_4") = Esc(aiSlide4)
    d("AI_INSIGHT_SLIDE_5") = Esc(aiSlide5)

    Set BuildPlaceholders = d
End Function

' --- Блок 3: KPI. «Текущая неделя» = максимальный yearWeek, реально присутствующий в данных.
' Окно недель берётся как последние WEEKS_WINDOW значений yearWeek из данных, а не
' арифметикой currentWeek - N + 1 (P1-4: арифметика ломалась на границе года). ---
Private Function BuildBlock3KPI() As String
    EnsureSnapshot

    Dim weeksWindow As Long
    weeksWindow = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))

    Dim byWeekDir As Object, byWeekDirTablet As Object
    Set byWeekDir = modAggregate.GroupCount(Array("yearWeek", "direction"), FArm())
    Set byWeekDirTablet = modAggregate.GroupCount(Array("yearWeek", "direction"), FTablet())

    Dim distinctByWeek As Object
    Set distinctByWeek = modAggregate.GroupCountDistinct(Array("yearWeek"), "number", FBase())

    Dim weeks As Variant
    weeks = modAggregate.SortKeys(distinctByWeek, True)

    If UBound(weeks) < LBound(weeks) Then
        BuildBlock3KPI = "<div class='kpi-grid'><div class='kpi'>Нет данных за выбранный период</div></div>"
        Exit Function
    End If

    Dim curW As String, prevW As String
    curW = KeyPart(weeks(UBound(weeks)), 0)
    If UBound(weeks) - 1 >= LBound(weeks) Then prevW = KeyPart(weeks(UBound(weeks) - 1), 0) Else prevW = ""

    Dim totalZN As Long, i As Long, taken As Long
    For i = UBound(weeks) To LBound(weeks) Step -1
        If taken >= weeksWindow Then Exit For
        totalZN = totalZN + CLng(distinctByWeek(weeks(i)))
        taken = taken + 1
    Next i

    Dim pctDGM_cur As Double, pctDENT_cur As Double, pctDGM_prev As Double, pctDENT_prev As Double
    pctDGM_cur = SafePercent(byWeekDirTablet, byWeekDir, curW & "|ДГМ|")
    pctDENT_cur = SafePercent(byWeekDirTablet, byWeekDir, curW & "|ДЭНТ|")
    If prevW <> "" Then
        pctDGM_prev = SafePercent(byWeekDirTablet, byWeekDir, prevW & "|ДГМ|")
        pctDENT_prev = SafePercent(byWeekDirTablet, byWeekDir, prevW & "|ДЭНТ|")
    End If

    Dim label As String
    label = WeekLabel(curW)

    BuildBlock3KPI = "<div class='kpi-grid'>" & _
        "<div class='kpi'><span class='kpi-cap'>Всего ЗН за " & taken & " нед.</span>" & _
        "<span class='kpi-val'>" & totalZN & "</span></div>" & _
        "<div class='kpi'><span class='kpi-cap'>ДГМ, % планшет (нед. " & Esc(label) & ")</span>" & _
        "<span class='kpi-val'>" & FormatPct(pctDGM_cur) & "</span>" & _
        "<span class='kpi-delta'>&Delta; " & FormatDeltaPP(pctDGM_cur - pctDGM_prev) & "</span></div>" & _
        "<div class='kpi'><span class='kpi-cap'>ДЭНТ, % планшет (нед. " & Esc(label) & ")</span>" & _
        "<span class='kpi-val'>" & FormatPct(pctDENT_cur) & "</span>" & _
        "<span class='kpi-delta'>&Delta; " & FormatDeltaPP(pctDENT_cur - pctDENT_prev) & "</span></div>" & _
        "</div>"
End Function

' --- Блок 1: строки direction -> arm, столбцы - недели, значения Count(Key);
' под каждой дирекцией - строка «% планшет» по ТОЙ ЖЕ сетке недель (P1-2). ---
Private Function BuildBlock1Table() As String
    EnsureSnapshot

    Dim counts As Object, total As Object, tablet As Object
    Set counts = modAggregate.GroupCount(Array("direction", "arm", "yearWeek"), FArm())
    Set total = modAggregate.GroupCount(Array("direction", "yearWeek"), FArm())
    Set tablet = modAggregate.GroupCount(Array("direction", "yearWeek"), FTablet())

    If counts.Count = 0 Then BuildBlock1Table = EmptyTable(): Exit Function

    Dim weeks As Variant, dirs As Variant
    weeks = AxisFromKeys(total, 1, True)
    dirs = AxisFromKeys(total, 0, False)

    Dim html As String, i As Long, j As Long
    html = "<table class='block-table matrix'><thead><tr><th>Дирекция</th><th>АРМ</th>"
    For j = LBound(weeks) To UBound(weeks)
        html = html & "<th>" & Esc(WeekLabel(weeks(j))) & "</th>"
    Next j
    html = html & "<th>Итого</th></tr></thead><tbody>"

    Dim arms As Variant
    arms = Array("ПК", "ПЛАНШЕТ")

    For i = LBound(dirs) To UBound(dirs)
        Dim dirName As String: dirName = CStr(dirs(i))
        Dim a As Long
        For a = LBound(arms) To UBound(arms)
            Dim rowTotal As Double: rowTotal = 0
            html = html & "<tr>"
            If a = LBound(arms) Then
                html = html & "<th rowspan='3' class='row-head'>" & Esc(dirName) & "</th>"
            End If
            html = html & "<td>" & Esc(arms(a)) & "</td>"
            For j = LBound(weeks) To UBound(weeks)
                Dim c As Double
                c = DictVal(counts, dirName & "|" & CStr(arms(a)) & "|" & CStr(weeks(j)) & "|")
                rowTotal = rowTotal + c
                html = html & "<td>" & CStr(CLng(c)) & "</td>"
            Next j
            html = html & "<td class='total'>" & CStr(CLng(rowTotal)) & "</td></tr>"
        Next a

        ' Строка «% планшет» - по каждой неделе этой дирекции (Content Spec §5).
        html = html & "<tr class='pct-row'><td>% планшет</td>"
        Dim sumTot As Double, sumTab As Double
        ' (!) v6.1. Dim в VBA - объявление на этапе компиляции, а не инициализация: переменная
        ' живёт всю процедуру и НЕ обнуляется на новой итерации цикла по дирекциям. Без явного
        ' сброса ДЭНТ получал накопленные суммы ДГМ: в столбце «Итого» выходило 8/23 = 34,8%
        ' вместо собственных 0/4 = 0,0%. В Блоке 4 сброс был, в Блоке 1 - нет; отсюда два
        ' разных ответа на одних данных. Проверено на отчёте от 04.09.2026.
        sumTot = 0: sumTab = 0
        For j = LBound(weeks) To UBound(weeks)
            Dim key As String: key = dirName & "|" & CStr(weeks(j)) & "|"
            sumTot = sumTot + DictVal(total, key)
            sumTab = sumTab + DictVal(tablet, key)
            ' (!) v6.1. Нулевой знаменатель - это ОТСУТСТВИЕ событий, а не 0% использования.
            ' SafePercent при делении на ноль возвращает 0, PctCell красит ячейку в красный и
            ' пишет «0,0%» - читается как «планшетом не пользуются», хотя данных просто нет.
            ' Блок 4 в той же ситуации ставит «—»; приводим Блок 1 к тому же поведению.
            If DictVal(total, key) = 0 Then
                html = html & "<td class='pct empty'>—</td>"
            Else
                html = html & PctCell(SafePercent(tablet, total, key))
            End If
        Next j
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot)
        Else
            html = html & "<td class='pct empty'>—</td>"
        End If
        html = html & "</tr>"
    Next i

    BuildBlock1Table = html & "</tbody></table>"
End Function

' --- Блоки 4 и 5: строки direction/postN, столбцы - недели, значение «% планшет» (P1-3).
' Ранее оба блока были копией Блока 1 (Count), что расходилось с Content Spec §5. ---
Private Function BuildPctMatrixTable(rowCol As String, rowCaption As String) As String
    EnsureSnapshot

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FArm())
    Set tablet = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FTablet())

    If total.Count = 0 Then BuildPctMatrixTable = EmptyTable(): Exit Function

    Dim weeks As Variant, rows As Variant
    weeks = AxisFromKeys(total, 1, True)
    rows = AxisFromKeys(total, 0, False)

    Dim html As String, i As Long, j As Long
    html = "<table class='block-table matrix'><thead><tr><th>" & Esc(rowCaption) & "</th>"
    For j = LBound(weeks) To UBound(weeks)
        html = html & "<th>" & Esc(WeekLabel(weeks(j))) & "</th>"
    Next j
    html = html & "<th>Итого</th></tr></thead><tbody>"

    For i = LBound(rows) To UBound(rows)
        Dim rowName As String: rowName = CStr(rows(i))
        Dim sumTot As Double, sumTab As Double
        sumTot = 0: sumTab = 0
        html = html & "<tr><th class='row-head'>" & Esc(rowName) & "</th>"
        For j = LBound(weeks) To UBound(weeks)
            Dim key As String: key = rowName & "|" & CStr(weeks(j)) & "|"
            sumTot = sumTot + DictVal(total, key)
            sumTab = sumTab + DictVal(tablet, key)
            If DictVal(total, key) = 0 Then
                html = html & "<td class='pct empty'>—</td>"
            Else
                html = html & PctCell(SafePercent(tablet, total, key))
            End If
        Next j
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot)
        Else
            html = html & "<td class='pct empty'>—</td>"
        End If
        html = html & "</tr>"
    Next i

    BuildPctMatrixTable = html & "</tbody></table>"
End Function

' --- Блок 2: количество ремонтов по постам - честный Distinct Count по number. ---
Private Function BuildBlock2Table() As String
    EnsureSnapshot

    Dim counts As Object
    Set counts = modAggregate.GroupCountDistinct(Array("direction", "postN"), "number", FBase())
    If counts.Count = 0 Then BuildBlock2Table = EmptyTable(): Exit Function

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(counts, True)

    Dim html As String
    html = "<table class='block-table'><thead><tr><th>Дирекция</th><th>Пост</th>" & _
           "<th>Кол-во ремонтов</th></tr></thead><tbody>"

    Dim i As Long
    For i = LBound(sorted) To UBound(sorted)
        Dim k As String: k = CStr(sorted(i))
        html = html & "<tr><td>" & Esc(KeyPart(k, 0)) & "</td><td>" & Esc(KeyPart(k, 1)) & _
               "</td><td class='num'>" & counts(k) & "</td></tr>"
    Next i

    BuildBlock2Table = html & "</tbody></table>"
End Function

' --- Блок 9: виды ремонта/дефектов. ---
Private Function BuildBlock9Table() As String
    EnsureSnapshot

    Dim counts As Object
    Set counts = modAggregate.GroupCount(Array("zn_type", "defekt_type"), FBase())
    If counts.Count = 0 Then BuildBlock9Table = EmptyTable(): Exit Function

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(counts, True)

    Dim html As String
    html = "<table class='block-table'><thead><tr><th>Вид ремонта</th><th>Тип дефекта</th>" & _
           "<th>Кол-во</th></tr></thead><tbody>"

    Dim i As Long
    For i = LBound(sorted) To UBound(sorted)
        Dim k As String: k = CStr(sorted(i))
        Dim defekt As String: defekt = KeyPart(k, 1)
        If defekt = "" Then defekt = "не указан"
        html = html & "<tr><td>" & Esc(KeyPart(k, 0)) & "</td><td>" & Esc(defekt) & _
               "</td><td class='num'>" & counts(k) & "</td></tr>"
    Next i

    BuildBlock9Table = html & "</tbody></table>"
End Function

' --- Блок 6: рейтинг инженеров. Сортировка по % планшет, отсечение групп с малым числом
' записей (P1-12: инженер с одним нарядом на ПК давал 0,0% и занимал верх антитопа). ---
Private Function Block6RankedTable(topN As Long, descending As Boolean) As String
    EnsureSnapshot

    Dim minRecords As Long
    minRecords = CLng(modMain.GetVariableDef("REPORT/MIN_RECORDS", "5"))

    Dim totalCounts As Object, tabletCounts As Object, avgDelta As Object
    Set totalCounts = modAggregate.GroupCount(Array("employee"), FArm())
    Set tabletCounts = modAggregate.GroupCount(Array("employee"), FTablet())
    ' P1-14: тот же базовый фильтр, что у соседних агрегатов, плюс срез по закрытым нарядам.
    Set avgDelta = modAggregate.GroupAverage(Array("employee"), "deltaHours", _
        Array("in_bounds=True", "arm@=ПК;ПЛАНШЕТ", "ready_for=Готов к выбытию"))

    Dim pctDict As Object: Set pctDict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In totalCounts.Keys
        If CLng(totalCounts(k)) >= minRecords Then
            pctDict(k) = SafePercent(tabletCounts, totalCounts, CStr(k))
        End If
    Next k

    If pctDict.Count = 0 Then
        Block6RankedTable = "<p class='empty-note'>Недостаточно данных: нет сотрудников с числом " & _
            "записей не менее " & minRecords & " (Variable/REPORT/MIN_RECORDS).</p>"
        Exit Function
    End If

    Dim sortedKeys As Variant
    sortedKeys = modAggregate.SortDictionaryKeysByValue(pctDict, descending)

    ' (!) v6.1. Столбец считает СОБЫТИЯ ПОДПИСАНИЯ (строки tbDATA), а не заказ-наряды:
    ' totalCounts приходит из GroupCount по строкам. Заголовок «Кол-во нарядов» смешивал
    ' единицы счёта треков: у Иванова 8 событий, но 4 наряда (ЗН-001, 006, 012, 013).
    ' Трек А считается в событиях подписания - так и подписываем.
    Dim html As String
    html = "<table class='block-table'><thead><tr><th>Инженер</th><th>% планшет</th>" & _
           "<th>Кол-во подписаний</th><th>Средняя длительность, ч</th></tr></thead><tbody>"

    Dim i As Long, shown As Long
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        If shown >= topN Then Exit For
        Dim empKey As String: empKey = CStr(sortedKeys(i))
        Dim empName As String: empName = KeyPart(empKey, 0)
        Dim deltaVal As String
        If avgDelta.Exists(empKey) Then deltaVal = Format(avgDelta(empKey), "0.0") Else deltaVal = "н/д"

        html = html & "<tr><td>" & Esc(empName) & "</td>" & _
            PctCell(CDbl(pctDict(empKey))) & _
            "<td class='num'>" & totalCounts(empKey) & "</td>" & _
            "<td class='num'>" & deltaVal & "</td></tr>"
        shown = shown + 1
    Next i

    Block6RankedTable = html & "</tbody></table>"
End Function

' =====================================================================================
' 5. Блоки 7/8 - попарная разница status_date «Готов к выбытию» между ДГМ и ДЭНТ
'    по одному number. postN относится к заказ-наряду целиком, берётся из строки ДГМ.
' =====================================================================================
' Возвращает Dictionary: number -> Array(yearWeekДГМ, postNДГМ, deltaHours). Кэшируется (P2-3).
Private Function BuildSyncPairs() As Object
    If mPairsReady Then Set BuildSyncPairs = mPairs: Exit Function

    EnsureSnapshot

    Dim dgmDates As Object: Set dgmDates = CreateObject("Scripting.Dictionary")
    Dim dgmWeeks As Object: Set dgmWeeks = CreateObject("Scripting.Dictionary")
    Dim dgmPosts As Object: Set dgmPosts = CreateObject("Scripting.Dictionary")
    Dim dentDates As Object: Set dentDates = CreateObject("Scripting.Dictionary")

    Dim r As Long, n As Long
    n = modAggregate.RowCount()

    For r = 1 To n
        ' Базовый фильтр применяется и здесь: строки с arm = "НЕ ПОДПИСАНО" и вне in_bounds
        ' не участвуют в сопоставлении дирекций (иначе Блоки 7/8 считались бы по другому
        ' набору строк, чем все остальные блоки).
        ' v5 (C-1): in_bounds сравнивался строго с "True". modAggregate.NormText приводит к
        ' "True" только НАСТОЯЩЕЕ булево значение ячейки; если Power Query выгрузил столбец
        ' текстом ("ИСТИНА"), фильтры блоков 1-9 это переживали (RowMatchesFilters прогоняет
        ' значение через NormBoolLiteral), а вот здесь сравнение не совпадало НИКОГДА - и
        ' Блоки 7/8 молча показывали "нет пар для сопоставления" при полных данных.
        If IsStatusReadyToLeave(modAggregate.CellText(r, "ready_for")) _
           And IsTrueText(modAggregate.CellText(r, "in_bounds")) _
           And (modAggregate.CellText(r, "arm") = "ПК" Or modAggregate.CellText(r, "arm") = "ПЛАНШЕТ") Then
            Dim num As String: num = modAggregate.CellText(r, "number")
            Dim dirVal As String: dirVal = modAggregate.CellText(r, "direction")
            Dim sd As Variant: sd = modAggregate.CellRaw(r, "status_date")
            If IsNumeric(sd) Then
                If dirVal = "ДГМ" Then
                    dgmDates(num) = CDbl(sd)
                    dgmWeeks(num) = modAggregate.CellText(r, "yearWeek")
                    dgmPosts(num) = modAggregate.CellText(r, "postN")
                ElseIf dirVal = "ДЭНТ" Then
                    dentDates(num) = CDbl(sd)
                End If
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

    Set mPairs = result
    mPairsReady = True
    Set BuildSyncPairs = result
End Function

' Агрегация пар: ключ "yearWeek" либо "yearWeek|postN" -> Array(среднее, кол-во пар).
Private Function SyncAggregate(byPost As Boolean) As Object
    Dim pairs As Object: Set pairs = BuildSyncPairs()
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In pairs.Keys
        Dim key As String
        If byPost Then
            key = CStr(pairs(k)(0)) & "|" & CStr(pairs(k)(1))
        Else
            key = CStr(pairs(k)(0))
        End If
        Dim d As Double: d = pairs(k)(2)
        If sums.Exists(key) Then
            sums(key) = sums(key) + d: counts(key) = counts(key) + 1
        Else
            sums(key) = d: counts(key) = 1
        End If
    Next k

    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    Dim srt As Variant
    srt = modAggregate.SortKeys(sums, Not byPost)   ' по неделям - численно, по "неделя x пост" - текстом
    Dim i As Long
    For i = LBound(srt) To UBound(srt)
        Dim kk As String: kk = CStr(srt(i))
        res(kk) = Array(sums(kk) / counts(kk), counts(kk))
    Next i

    Set SyncAggregate = res
End Function

' HTML-таблица Блока 7 (по неделям) или Блока 8 (по неделям x постам).
Private Function BuildSyncTable(byPost As Boolean) As String
    Dim agg As Object: Set agg = SyncAggregate(byPost)
    If agg.Count = 0 Then
        BuildSyncTable = "<p class='empty-note'>Нет пар «ДГМ &harr; ДЭНТ» по одному заказ-наряду " & _
            "для сопоставления.</p>"
        Exit Function
    End If

    Dim threshold As Double
    threshold = CDbl(modMain.GetVariableDef("REPORT/SYNC_THRESHOLD_MIN", "0")) / 60#

    Dim html As String
    html = "<table class='block-table'><thead><tr><th>Неделя</th>"
    If byPost Then html = html & "<th>Пост</th>"
    html = html & "<th>Средняя разница, ч</th><th>Пар нарядов</th></tr></thead><tbody>"

    Dim k As Variant
    For Each k In agg.Keys
        Dim v As Variant: v = agg(k)
        Dim cls As String
        cls = ""
        If threshold > 0 Then
            If CDbl(v(0)) > threshold Then cls = " class='warn'"
        End If
        html = html & "<tr><td>" & Esc(WeekLabel(KeyPart(k, 0))) & "</td>"
        If byPost Then html = html & "<td>" & Esc(KeyPart(k, 1)) & "</td>"
        html = html & "<td" & cls & ">" & Format(v(0), "0.0") & "</td>" & _
               "<td class='num'>" & CStr(v(1)) & "</td></tr>"
    Next k

    BuildSyncTable = html & "</tbody></table>"
End Function

' v5 (C-1): распространённые написания булева значения. Дублирует приватную
' NormBoolLiteral из modAggregate сознательно - чтобы не править Core ради специфики МТО.
' v6 (P2-C): единое правило сравнения статуса. Раньше на одно и то же поле ready_for
' действовали три разных правила: fnComputeGroupMetrics нормализует "ё"->"е" и регистр,
' modAggregate.RowMatchesFilters сравнивает через StrComp(vbTextCompare) без учёта регистра,
' а BuildSyncPairs - строгим "=". Слова "выбытию" опечатка через "ё" не касается, поэтому
' расхождение не стреляло; но при смене регистра в выгрузке Блоки 1-9 строку увидели бы,
' а Блоки 7/8 - нет, и таблица пар молча опустела бы. Это тот же класс дефекта, что C-1.
Private Function NormStatus(s As String) As String
    NormStatus = LCase$(Replace(Trim$(s), ChrW$(&H451), ChrW$(&H435)))  ' "ё" -> "е"
End Function

Private Function IsStatusReadyToLeave(s As String) As Boolean
    IsStatusReadyToLeave = (Left$(NormStatus(s), Len("готов к выбытию")) = "готов к выбытию")
End Function

Private Function IsTrueText(s As String) As Boolean
    Select Case UCase$(Trim$(s))
        Case "TRUE", "ИСТИНА", "1", "ДА": IsTrueText = True
        Case Else: IsTrueText = False
    End Select
End Function

Private Function EmptyTable() As String
    EmptyTable = "<p class='empty-note'>Нет данных, удовлетворяющих фильтру " & _
        "(in_bounds = ИСТИНА, arm &isin; {ПК, ПЛАНШЕТ}; строки «НЕ ПОДПИСАНО» исключены).</p>"
End Function
