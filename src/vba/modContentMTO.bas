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
'              затем ключи слайдов внутри развёрнутого текста.
'   P0-2     - BuildPlaceholders принимает выводы ИИ параметрами.
'   P1-8     - все значения проходят через modHTMLEngine.HtmlEscape.
'   P1-12    - антитоп Блока 6 отсекает сотрудников с числом записей < REPORT/MIN_RECORDS.
'   P1-14    - avgDelta в Блоке 6 фильтруется согласованно с соседними агрегатами.
'   P1-15    - полный JsonEscape (управляющие символы, \n, \r, \t).
'   P2-3     - BuildSyncPairs кэшируется на прогон (был 4 полных прохода).
'
' v7.0 - постановка v1.1 (task-for-coder.md, «Презентация и блоки данных»): 7 слайдов,
'        Дашборд 1, ключ REPORT/WEEK, окно Блока 1 по присутствующим yearWeek, Блок 2 с оценкой,
'        понедельный Блок 6 + рейтинг за отчётную неделю, топ-10 аномалий Блока 7, подблок 9а,
'        выводы ИИ slide1..slide7 (кэш mInsights, сигнатуры контракта не менялись), псевдонимизация
'        employee маркерами [EMP_N], JSON-дамп расшифровки за 2 недели с data-drill по ячейкам,
'        офлайн-график по Блоку 2, подписи расчёта, NormStatus во всех сравнениях статусов.
'
' Версия 7.0.1 от 07.09.2026 - отладка прогона SelfTest:
'   - Block6WeeklyCore: переменная «tab» совпадала с ключевым словом VBA Tab; из-за ленивой
'     компиляции процедур модальный Compile error: Syntax error возникал при первом вызове
'     (BuildPlaceholders, метки bp8/bp9). Переименована в tabObj (6 мест использования).
'   - удалена временная диагностика DbgBp / bp_log.txt.
'
' Версия 7.1 от 08.09.2026 - рестайлинг отчёта под дизайн-систему remzona-reports.html:
'   - BuildDashboard: период вынесен в заголовок панели (mock-bar ttl) вместо KPI-карточки
'     «Период»; переключатель периодов - чипы .chip/.on вместо кнопок .dash-btn; сетка KPI -
'     классы .kpis/.kpi вместо .kpi-grid/.kpi (шаблон tmp_index.html v2.0).
'   - DashKpiGrid: удалена первая KPI-карточка «Период»; подписи и значения - .lab/.val.
'   - новая приватная DashPeriodCaption - единый источник заголовков периодов дашборда.
'
' Версия 7.2 от 08.09.2026 - закрыты остатки tz_Reports2.md:
'   - BuildBlock6Weekly получил параметр byDept: BLOCK_6_DENT/BLOCK_6_DGM - только «По сотрудникам»,
'     новые плейсхолдеры BLOCK_6_DENT_DEPT/BLOCK_6_DGM_DEPT - только «По подразделениям» (ТЗ раздел 4).
'   - пустой REPORT/SLIDE_ZONES -> BLOCK_1_*_ZONES выводят пояснение вместо пустоты (ТЗ 3.2).
'   - «Создали ЗН» Дашборда считается без FBase - по всей истории, включая «НЕ ПОДПИСАНО» (ТЗ 3.1).
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

Private mInsights As Object       ' кэш выводов ИИ: "slide1".."slide7" (task-for-coder §5.1)
Private mInsightsReady As Boolean

Private mEmpList As Variant       ' единый алфавитный список ФИО снимка для маркеров [EMP_N]
Private mEmpReady As Boolean

Private mWeeks As Variant         ' отсортированные по возрастанию yearWeek из данных (FBase)
Private mWeeksReady As Boolean
Private mReportWeek As Long       ' отчётная неделя (REPORT/WEEK); -2 = не вычислена, 0 = нет данных
Private mLatestWeek As Long       ' последняя неделя данных; -2 = не вычислена, 0 = нет данных

Private mDashboard As Object      ' кэш показателей Дашборда 1 (ytd/prev/last)
Private mDashboardReady As Boolean

Private mBlock2 As Object         ' кэш агрегатов Блока 2 (таблица + график)
Private mBlock2Ready As Boolean

Private mDumpJson As String       ' кэш JSON-дампа расшифровки (один раз на прогон, §3)
Private mDumpReady As Boolean

' =====================================================================================
' 0. Инфраструктура
' =====================================================================================
Private Function TbData() As ListObject
    Set TbData = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA")
End Function

' Контракт Core §16. Начиная с v3.1 процедура не строит PivotTable - она готовит
' снимок данных для modAggregate и сбрасывает кэши блоков. Имя сохранено, т.к.
' зафиксировано контрактом Core.
Public Sub BuildPivots()
    ResetContentCaches

    modAggregate.EndSnapshot
    modAggregate.BeginSnapshot TbData()
End Sub

' Сброс всех кэшей одного прогона. mInsights сбрасывается здесь же - иначе при
' недоступном ИИ (STUB) отчёт содержал бы протухшие выводы предыдущего прогона
' (task-for-coder §5.1 / §6).
Private Sub ResetContentCaches()
    mPairsReady = False
    Set mPairs = Nothing
    mMultiYear = -1

    mInsightsReady = False
    Set mInsights = Nothing

    mEmpReady = False
    If IsArray(mEmpList) Then Erase mEmpList

    mWeeksReady = False
    If IsArray(mWeeks) Then Erase mWeeks
    mReportWeek = -2
    mLatestWeek = -2

    mDashboardReady = False
    Set mDashboard = Nothing

    mBlock2Ready = False
    Set mBlock2 = Nothing

    mDumpReady = False
    mDumpJson = ""
End Sub

Private Sub EnsureSnapshot()
    If Not modAggregate.IsReady Then
        ResetContentCaches
        modAggregate.BeginSnapshot TbData()
    End If
End Sub

' Защитный контракт: проверяет обязательные столбцы tbDATA один раз до сборки отчёта.
' Возвращает False и показывает внятное сообщение, если какого-то столбца нет, -
' вместо падения на 13-й секунде с Err -2147221502 «Столбец не найден».
Public Function ValidateRequiredColumns() As Boolean
    EnsureSnapshot
    Dim required() As String
    required = Split("yearWeek,postN,Key,in_bounds,arm,number,ready_for,status_date,date,direction,zn_type,defekt_type", ",")
    Dim missing As String
    missing = ""
    Dim i As Long
    For i = LBound(required) To UBound(required)
        If Not modAggregate.HasColumn(required(i)) Then
            missing = missing & required(i) & ", "
        End If
    Next i
    If missing <> "" Then
        missing = Left$(missing, Len(missing) - 2)
        modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "ValidateRequiredColumns", _
            "В tbDATA отсутствуют обязательные столбцы: " & missing
        MsgBox "В таблице tbDATA отсутствуют обязательные столбцы:" & vbCrLf & missing & vbCrLf & _
               "Обновите M-код (install.ps1) и перезагрузите данные.", vbCritical
        ValidateRequiredColumns = False
        Exit Function
    End If
    ValidateRequiredColumns = True
End Function

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

' Добавляет ещё один фильтр к набору (массивы-фильтры передаются в modAggregate).
Private Function AppendFilter(base As Variant, extra As String) As Variant
    Dim n As Long
    n = UBound(base) - LBound(base) + 1
    Dim out() As Variant
    ReDim out(0 To n)
    Dim i As Long
    For i = 0 To n - 1
        out(i) = base(LBound(base) + i)
    Next i
    out(n) = extra
    AppendFilter = out
End Function

' =====================================================================================
' 1. Работа с осями матриц (недели/строки)
' =====================================================================================
' Определяет, присутствует ли в данных больше одного года - от этого зависит подпись недели.
Private Function IsMultiYear() As Boolean
    If mMultiYear <> -1 Then IsMultiYear = (mMultiYear = 1): Exit Function

    EnsureSnapshot
    ' Год берём из yearWeek (year*100+week), а не из year_status: выгрузка 2026
    ' не содержит *_status, и опора на year_status роняла отчёт (Err -2147221502).
    Dim years As Object
    Set years = CreateObject("Scripting.Dictionary")
    Dim weeks As Variant
    weeks = WeeksList()
    Dim i As Long
    For i = LBound(weeks) To UBound(weeks)
        years(CStr(CLng(KeyPart(weeks(i), 0)) \ 100)) = 1
    Next i
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

Private Function Esc(s As Variant) As String
    Esc = modHTMLEngine.HtmlEscape(CStr(s))
End Function

' Ячейка «% планшет» с раскраской по палитре. drill - JSON-условие клик-расшифровки (п.10).
Private Function PctCell(pct As Double, Optional drill As String) As String
    Dim color As String
    color = modColor.PercentToColor(pct, COLOR_BAD, COLOR_WARN, COLOR_GOOD)
    PctCell = "<td class='pct'" & DrillAttr(drill) & " style='background:" & color & ";color:#fff;'>" & FormatPct(pct) & "</td>"
End Function

' Числовая ячейка с клик-расшифровкой (п.10): визуально - обычный текст, data-drill для JS.
Private Function NumCell(v As Variant, Optional drill As String) As String
    NumCell = "<td class='num'" & DrillAttr(drill) & ">" & CStr(v) & "</td>"
End Function

' Атрибут data-drill: JSON-объект условий, экранированный для HTML-атрибута.
Private Function DrillAttr(parts As String) As String
    If Trim$(parts) = "" Then DrillAttr = "": Exit Function
    DrillAttr = " data-drill='" & Esc("{" & parts & "}") & "'"
End Function

' Пара «поле»:«значение» для JSON-условия расшифровки (значение - через JsonEscape).
Private Function J(field As String, value As String) As String
    J = """" & field & """:""" & JsonEscape(value) & """"
End Function

Private Function EmptyTable() As String
    EmptyTable = "<p class='empty-note'>Нет данных, удовлетворяющих фильтру " & _
        "(in_bounds = ИСТИНА, arm из {ПК, ПЛАНШЕТ}; строки «НЕ ПОДПИСАНО» исключены).</p>"
End Function

' Подпись «как считается» под блоком (п.12).
Private Function Recipe(text As String) As String
    Recipe = "<p class='recipe'>" & text & "</p>"
End Function

' =====================================================================================
' 2. Отчётная неделя и недельные окна (только по присутствующим yearWeek, без
'    арифметики номеров недель - она ломается на границе года, дефект P1-4)
' =====================================================================================
' Отсортированные по возрастанию yearWeek, реально присутствующие в данных (базовый фильтр).
Private Function WeeksList() As Variant
    If mWeeksReady Then WeeksList = mWeeks: Exit Function
    EnsureSnapshot
    Dim d As Object
    Set d = modAggregate.DistinctValues("yearWeek", FBase())
    mWeeks = modAggregate.SortKeys(d, True)
    mWeeksReady = True
    WeeksList = mWeeks
End Function

' Отчётная неделя (ключ REPORT/WEEK, task-for-coder §1):
'   авто (пусто/0): второй с конца yearWeek, присутствующий в данных;
'   явное N (1..50): последний год, в котором неделя N встречается в данных.
' Возвращает 0, если данных нет.
Private Function ReportWeekValue() As Long
    If mReportWeek <> -2 Then ReportWeekValue = mReportWeek: Exit Function

    Dim weeks As Variant
    weeks = WeeksList()
    Dim wk As Long
    wk = 0
    If UBound(weeks) < LBound(weeks) Then
        mReportWeek = 0
        ReportWeekValue = 0
        Exit Function
    End If

    Dim raw As String
    raw = Trim$(modMain.GetVariableDef("REPORT/WEEK", "0"))
    If raw = "" Or raw = "0" Then
        ' Авто: второй с конца присутствующий; при единственной неделе - она сама.
        If UBound(weeks) - 1 >= LBound(weeks) Then
            wk = CLng(KeyPart(weeks(UBound(weeks) - 1), 0))
        Else
            wk = CLng(KeyPart(weeks(UBound(weeks)), 0))
        End If
    Else
        Dim n As Long
        n = CLng(Val(raw))
        Dim i As Long
        For i = UBound(weeks) To LBound(weeks) Step -1
            If CLng(KeyPart(weeks(i), 0)) Mod 100 = n Then
                wk = CLng(KeyPart(weeks(i), 0))
                Exit For
            End If
        Next i
        If wk = 0 Then
            ' Явной недели в данных нет - откат на авто (второй с конца).
            If UBound(weeks) - 1 >= LBound(weeks) Then
                wk = CLng(KeyPart(weeks(UBound(weeks) - 1), 0))
            Else
                wk = CLng(KeyPart(weeks(UBound(weeks)), 0))
            End If
        End If
    End If

    mReportWeek = wk
    ReportWeekValue = wk
End Function

' «Последняя неделя» = максимальный yearWeek, присутствующий в данных (§3 допущений).
Private Function LatestWeekValue() As Long
    If mLatestWeek <> -2 Then LatestWeekValue = mLatestWeek: Exit Function

    Dim weeks As Variant
    weeks = WeeksList()
    If UBound(weeks) < LBound(weeks) Then
        mLatestWeek = 0
        LatestWeekValue = 0
        Exit Function
    End If
    mLatestWeek = CLng(KeyPart(weeks(UBound(weeks)), 0))
    LatestWeekValue = mLatestWeek
End Function

' Неделя, предшествующая заданной в данных (последняя присутствующая < w); "" если нет.
Private Function PrevWeekBefore(w As Long) As String
    Dim weeks As Variant
    weeks = WeeksList()
    Dim i As Long
    For i = UBound(weeks) To LBound(weeks) Step -1
        If CLng(KeyPart(weeks(i), 0)) < w Then
            PrevWeekBefore = CStr(KeyPart(weeks(i), 0))
            Exit Function
        End If
    Next i
    PrevWeekBefore = ""
End Function

' Последние limit присутствующих yearWeek <= отчётной недели, по УБЫВАНИЮ, массив 1..count.
Private Function RecentWeeksUpTo(limit As Long, ByRef count As Long) As Variant
    Dim weeks As Variant
    weeks = WeeksList()
    If UBound(weeks) < LBound(weeks) Then
        count = 0
        RecentWeeksUpTo = Array()
        Exit Function
    End If

    Dim rw As Long
    rw = ReportWeekValue()
    Dim i As Long

    ' Проход 1: подсчёт подходящих недель (не более limit).
    count = 0
    For i = UBound(weeks) To LBound(weeks) Step -1
        If count >= limit Then Exit For
        If CLng(KeyPart(weeks(i), 0)) <= rw Then count = count + 1
    Next i

    ' Проход 2: заполнение.
    Dim out() As Variant
    ReDim out(1 To count)
    Dim pos As Long
    pos = 0
    For i = UBound(weeks) To LBound(weeks) Step -1
        If pos >= count Then Exit For
        If CLng(KeyPart(weeks(i), 0)) <= rw Then
            pos = pos + 1
            out(pos) = weeks(i)
        End If
    Next i
    RecentWeeksUpTo = out
End Function

' Нормализация REPORT/SLIDE_ZONES: "+" -> ";" + Trim каждой части (§5.2). Без неё
' Split(";") получит один элемент "СТК+ПРК", фильтр postN@= не совпадёт ни с чем, и
' вторая матрица Блока 1 молча станет пустой.
Private Function NormalizeZones(raw As String) As String
    Dim parts() As String
    parts = Split(raw, "+")
    Dim out As String
    out = ""
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String
        p = Trim$(parts(i))
        If p <> "" Then out = out & p & ";"
    Next i
    NormalizeZones = out
End Function

' =====================================================================================
' 3. Дашборд 1 (п.4): периоды «с начала года» / «предыдущая неделя» / «последняя неделя»
'    Показатели: создали (Distinct number по date), закрыли (Distinct number со статусом
'    «Готов к выбытию» по status_date), % планшет, без поста ремзоны (Distinct number
'    с пустым postN), медиана (по кэшу mPairs, вывод «чч:мм», §5.5).
' =====================================================================================
' Один проход по снимку: заполняет mDashboard (period -> метрики).
Private Sub EnsureDashboard()
    If mDashboardReady Then Exit Sub
    EnsureSnapshot

    Dim rw As Long
    rw = ReportWeekValue()
    Dim lw As Long
    lw = LatestWeekValue()
    Dim thisYear As Long
    thisYear = Year(Date)

    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")
    Dim per As Variant
    For Each per In Array("ytd", "prev", "last")
        Dim m As Object
        Set m = CreateObject("Scripting.Dictionary")
        Set m("created") = CreateObject("Scripting.Dictionary")
        Set m("closed") = CreateObject("Scripting.Dictionary")
        Set m("noPost") = CreateObject("Scripting.Dictionary")
        m("total") = CDbl(0)
        m("tablet") = CDbl(0)
        Set stats(per) = m
    Next per

    Dim n As Long
    n = modAggregate.RowCount()
    Dim r As Long
    Dim r2 As Long, num2 As String, dte2 As Variant, wdte2 As Long
    For r = 1 To n
        If IsTrueText(modAggregate.CellText(r, "in_bounds")) Then
            Dim armT As String
            armT = modAggregate.CellText(r, "arm")
            If armT = "ПК" Or armT = "ПЛАНШЕТ" Then
                Dim num As String
                num = modAggregate.CellText(r, "number")
                Dim sd As Variant
                sd = modAggregate.CellRaw(r, "status_date")
                Dim wsd As Long
                wsd = 0
                If IsNumeric(sd) Then wsd = WeekKeyOf(CDbl(sd))
                Dim ySt As Long
                ' Год события из yearWeek (year*100+week): выгрузка 2026 не содержит
                ' year_status, опора на него роняла отчёт (Err -2147221502).
                ySt = YearOfWeekKey(modAggregate.CellText(r, "yearWeek"))
                Dim isTab As Boolean
                isTab = (armT = "ПЛАНШЕТ")
                Dim isLeave As Boolean
                isLeave = IsStatusReadyToLeave(modAggregate.CellText(r, "ready_for"))
                Dim postN As String
                postN = modAggregate.CellText(r, "postN")
                Dim mm As Object

                ' «С начала года»: события текущего года (год из yearWeek).
                ' «Создали ЗН» считается отдельным проходом ниже - без FBase (ТЗ 3.1).
                If ySt = thisYear Then
                    Set mm = stats("ytd")
                    mm("total") = CDbl(mm("total")) + 1
                    If isTab Then mm("tablet") = CDbl(mm("tablet")) + 1
                    If isLeave Then mm("closed")(num) = True
                    If postN = "" Then mm("noPost")(num) = True
                End If

                ' Предыдущая неделя (отчётная, REPORT/WEEK): по status_date.
                If rw > 0 And wsd = rw Then
                    Set mm = stats("prev")
                    mm("total") = CDbl(mm("total")) + 1
                    If isTab Then mm("tablet") = CDbl(mm("tablet")) + 1
                    If isLeave Then mm("closed")(num) = True
                    If postN = "" Then mm("noPost")(num) = True
                End If

                ' Последняя неделя данных.
                If lw > 0 And wsd = lw Then
                    Set mm = stats("last")
                    mm("total") = CDbl(mm("total")) + 1
                    If isTab Then mm("tablet") = CDbl(mm("tablet")) + 1
                    If isLeave Then mm("closed")(num) = True
                    If postN = "" Then mm("noPost")(num) = True
                End If
            End If
        End If
    Next r

    ' «Создали ЗН» - Distinct number по дате создания (date), БЕЗ базового фильтра FBase:
    ' создание ЗН не зависит от статуса подписи - строки «НЕ ПОДПИСАНО» участвуют (ТЗ 3.1).
    For r2 = 1 To n
        num2 = modAggregate.CellText(r2, "number")
        If num2 <> "" Then
            dte2 = modAggregate.CellRaw(r2, "date")
            wdte2 = 0
            If IsNumeric(dte2) Then wdte2 = WeekKeyOf(CDbl(dte2))
            If wdte2 > 0 And wdte2 \ 100 = thisYear Then stats("ytd")("created")(num2) = True
            If rw > 0 And wdte2 = rw Then stats("prev")("created")(num2) = True
            If lw > 0 And wdte2 = lw Then stats("last")("created")(num2) = True
        End If
    Next r2

    Set mDashboard = stats
    mDashboardReady = True
End Sub

' Медиана deltaHours по парам из кэша mPairs (§5.5): ЗН со статусом «Готов к выбытию»
' в обеих дирекциях. Возвращает -1, если пар нет.
Private Function MedianFromPairs() As Double
    Dim pairs As Object
    Set pairs = BuildSyncPairs()
    If pairs.Count = 0 Then MedianFromPairs = -1: Exit Function

    Dim vals() As Double
    ReDim vals(0 To pairs.Count - 1)
    Dim i As Long
    i = 0
    Dim k As Variant
    For Each k In pairs.Keys
        vals(i) = CDbl(pairs(k)(2))
        i = i + 1
    Next k

    ' Сортировка по возрастанию (пузырёк - число пар невелико).
    Dim a As Long, b As Long, t As Double
    For a = 0 To pairs.Count - 2
        For b = 0 To pairs.Count - 2 - a
            If vals(b) > vals(b + 1) Then
                t = vals(b): vals(b) = vals(b + 1): vals(b + 1) = t
            End If
        Next b
    Next a

    If pairs.Count Mod 2 = 1 Then
        MedianFromPairs = vals(pairs.Count \ 2)
    Else
        MedianFromPairs = (vals(pairs.Count \ 2 - 1) + vals(pairs.Count \ 2)) / 2
    End If
End Function

' Часы -> «чч:мм» с округлением до минут (решение заказчика).
Private Function FormatHHMM(hours As Double) As String
    Dim totalMin As Long
    totalMin = CLng(Round(hours * 60))
    FormatHHMM = CStr(totalMin \ 60) & ":" & Right$("0" & CStr(totalMin Mod 60), 2)
End Function

' ISO-год*100+неделя для даты (для отнесения date/status_date к неделе в Дашборде).
Private Function WeekKeyOf(d As Double) As Long
    Dim dt As Date
    dt = CDate(d)
    Dim wk As Long
    wk = DatePart("ww", dt, vbMonday, vbFirstFourDays)
    Dim yr As Long
    yr = Year(dt)
    If Month(dt) = 12 And wk = 1 Then yr = yr + 1
    If Month(dt) = 1 And wk >= 52 Then yr = yr - 1
    WeekKeyOf = yr * 100 + wk
End Function

Private Function WeekKeyOfSafe(v As Variant) As Long
    If IsNumeric(v) Then WeekKeyOfSafe = WeekKeyOf(CDbl(v)) Else WeekKeyOfSafe = 0
End Function

' Год из yearWeek (year*100+week). Пустой/нечисловой yearWeek -> 0 (строка «НЕ ПОДПИСАНО»).
Private Function YearOfWeekKey(yw As Variant) As Long
    If Not IsNumeric(yw) Then YearOfWeekKey = 0: Exit Function
    Dim n As Long
    n = CLng(yw)
    If n <= 0 Then YearOfWeekKey = 0: Exit Function
    YearOfWeekKey = n \ 100
End Function

Private Function YearOfSafe(v As Variant) As Long
    If IsNumeric(v) Then YearOfSafe = Year(CDate(CDbl(v))) Else YearOfSafe = 0
End Function

' HTML Дашборда с переключателем периодов. primary: "ytd" | "prev" | "last" (активная панель).
Private Function BuildDashboard(primary As String) As String
    EnsureDashboard

    Dim periods As Variant
    periods = Array("ytd", "prev", "last")
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    labels("ytd") = "С начала года"
    labels("prev") = "Пред. неделя"
    labels("last") = "Последняя неделя"

    Dim html As String
    ' Панель-макет в стиле remzona-reports: mock-bar (заголовок периода + чипы) и KPI-сетка.
    html = "<div class='dashboard mock'>"
    html = html & "<div class='mock-bar'>"
    html = html & "<span class='ttl dash-ttl'>Дашборд · " & Esc(DashPeriodCaption(primary)) & "</span>"
    html = html & "<div class='chips'>"
    Dim p As Variant
    For Each p In periods
        Dim cls As String
        cls = "chip"
        If CStr(p) = primary Then cls = cls & " on"
        html = html & "<button class='" & cls & "' data-dash-period='" & CStr(p) & "'>" & Esc(CStr(labels(p))) & "</button>"
    Next p
    html = html & "</div></div>"
    For Each p In periods
        Dim hidden As String
        hidden = ""
        If CStr(p) <> primary Then hidden = " hidden"
        html = html & "<div class='kpis' data-dash-panel='" & CStr(p) & "'" & _
            " data-dash-title='Дашборд · " & Esc(DashPeriodCaption(CStr(p))) & "'" & hidden & ">" & _
            DashKpiGrid(CStr(p)) & "</div>"
    Next p
    html = html & "</div>"

    BuildDashboard = html & Recipe("Создали - уникальные ЗН по дате создания (date) за период; " & _
        "Закрыли - уникальные ЗН со статусом «Готов к выбытию» по дате статуса (status_date); " & _
        "% планшет - события ПЛАНШЕТ / (ПК + ПЛАНШЕТ); Без поста ремзоны - уникальные ЗН с пустым postN; " & _
        "Медиана - по ЗН со статусом «Готов к выбытию» в обеих дирекциях (вывод «чч:мм»).")
End Function

' Заголовок периода для mock-bar дашборда (v7.1, единый источник: заголовок панели и
' data-dash-title чипов-переключателей).
Private Function DashPeriodCaption(per As String) As String
    If per = "ytd" Then
        DashPeriodCaption = "С начала года " & CStr(Year(Date))
    ElseIf per = "prev" Then
        DashPeriodCaption = "Предыдущая неделя (нед. " & WeekLabel(ReportWeekValue()) & ")"
    Else
        DashPeriodCaption = "Последняя неделя (нед. " & WeekLabel(LatestWeekValue()) & ")"
    End If
End Function

' Карточки KPI одного периода.
Private Function DashKpiGrid(per As String) As String
    Dim mm As Object
    Set mm = mDashboard(per)

    Dim thisYear As Long
    thisYear = Year(Date)
    Dim rw As Long, lw As Long
    rw = ReportWeekValue()
    lw = LatestWeekValue()

    Dim createdDrill As String, closedDrill As String, pctDrill As String, noPostDrill As String
    If per = "ytd" Then
        createdDrill = J("date_year", CStr(thisYear))
        closedDrill = J("status_year", CStr(thisYear)) & "," & J("norm_status", "готов к выбытию")
        pctDrill = J("status_year", CStr(thisYear))
        noPostDrill = J("status_year", CStr(thisYear)) & "," & J("postN", "")
    ElseIf per = "prev" Then
        createdDrill = J("date_week", CStr(rw))
        closedDrill = J("status_week", CStr(rw)) & "," & J("norm_status", "готов к выбытию")
        pctDrill = J("status_week", CStr(rw))
        noPostDrill = J("status_week", CStr(rw)) & "," & J("postN", "")
    Else
        createdDrill = J("date_week", CStr(lw))
        closedDrill = J("status_week", CStr(lw)) & "," & J("norm_status", "готов к выбытию")
        pctDrill = J("status_week", CStr(lw))
        noPostDrill = J("status_week", CStr(lw)) & "," & J("postN", "")
    End If

    Dim total As Double, tablet As Double
    total = CDbl(mm("total"))
    tablet = CDbl(mm("tablet"))

    Dim medianVal As String, medianDrill As String
    Dim med As Double
    med = MedianFromPairs()
    If med < 0 Then
        medianVal = "н/д"
        medianDrill = ""
    Else
        medianVal = FormatHHMM(med)
        medianDrill = J("norm_status", "готов к выбытию")
    End If

    Dim html As String
    html = "<div class='kpi'><div class='lab'>Создали заказ-нарядов</div>" & _
        "<div class='val'" & DrillAttr(createdDrill) & ">" & CStr(mm("created").Count) & "</div></div>"
    html = html & "<div class='kpi'><div class='lab'>Закрыли заказ-нарядов</div>" & _
        "<div class='val'" & DrillAttr(closedDrill) & ">" & CStr(mm("closed").Count) & "</div></div>"
    html = html & "<div class='kpi'><div class='lab'>% планшет</div>" & _
        "<div class='val'" & DrillAttr(pctDrill) & ">" & DashPct(total, tablet) & "</div></div>"
    html = html & "<div class='kpi'><div class='lab'>Без поста ремзоны</div>" & _
        "<div class='val'" & DrillAttr(noPostDrill) & ">" & CStr(mm("noPost").Count) & "</div></div>"
    html = html & "<div class='kpi'><div class='lab'>Медиана</div>" & _
        "<div class='val'" & DrillAttr(medianDrill) & ">" & medianVal & "</div></div>"
    DashKpiGrid = html
End Function

Private Function DashPct(total As Double, tablet As Double) As String
    If total <= 0 Then DashPct = "—" Else DashPct = FormatPct(tablet / total)
End Function

' =====================================================================================
' 4. Блок 1 (п.3): строки direction -> arm, столбцы - недели, значения Count.
'    Параметры: direction ("" - обе), zones (пусто - все). Окно - последние
'    REPORT/WEEKS_WINDOW присутствующих yearWeek <= отчётной недели (P5.3).
' =====================================================================================
Private Function BuildBlock1Table(direction As String, zones As String) As String
    EnsureSnapshot

    Dim f As Variant
    f = FArm()
    If direction <> "" Then f = AppendFilter(f, "direction=" & direction)
    If Trim$(zones) <> "" Then f = AppendFilter(f, "postN@=" & NormalizeZones(zones))
    Dim ft As Variant
    ft = FTablet()
    If direction <> "" Then ft = AppendFilter(ft, "direction=" & direction)
    If Trim$(zones) <> "" Then ft = AppendFilter(ft, "postN@=" & NormalizeZones(zones))

    Dim counts As Object, total As Object, tablet As Object
    Set counts = modAggregate.GroupCount(Array("direction", "arm", "yearWeek"), f)
    Set total = modAggregate.GroupCount(Array("direction", "yearWeek"), f)
    Set tablet = modAggregate.GroupCount(Array("direction", "yearWeek"), ft)

    Dim title As String
    title = "Дирекция: "
    If direction = "" Then title = title & "ДГМ + ДЭНТ" Else title = title & direction
    If Trim$(zones) = "" Then
        title = title & " · ремзоны: все"
    Else
        title = title & " · ремзоны: " & Esc(Replace(Trim$(zones), "+", ", "))
    End If

    If counts.Count = 0 Then
        BuildBlock1Table = "<h3>" & title & "</h3>" & EmptyTable()
        Exit Function
    End If

    ' Окно недель: только присутствующие в данных, не превосходящие отчётную неделю.
    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))
    Dim allWeeks As Variant
    allWeeks = AxisFromKeys(total, 1, True)
    Dim rw As Long
    rw = ReportWeekValue()
    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = UBound(allWeeks) To LBound(allWeeks) Step -1
        If cnt >= windowSize Then Exit For
        If CLng(KeyPart(allWeeks(i), 0)) <= rw Then cnt = cnt + 1
    Next i

    If cnt = 0 Then
        BuildBlock1Table = "<h3>" & title & "</h3>" & EmptyTable()
        Exit Function
    End If

    Dim sel() As Variant
    ReDim sel(1 To cnt)
    Dim pos As Long
    pos = 0
    For i = UBound(allWeeks) To LBound(allWeeks) Step -1
        If pos >= cnt Then Exit For
        If CLng(KeyPart(allWeeks(i), 0)) <= rw Then
            pos = pos + 1
            sel(pos) = allWeeks(i)
        End If
    Next i

    Dim dirs As Variant
    dirs = AxisFromKeys(total, 0, False)
    Dim arms As Variant
    arms = Array("ПК", "ПЛАНШЕТ")

    Dim html As String, wj As Long
    html = "<h3>" & title & "</h3>"
    html = html & "<table class='block-table matrix'><thead><tr><th>Дирекция</th><th>АРМ</th>"
    For wj = cnt To 1 Step -1
        html = html & "<th>" & Esc(WeekLabel(KeyPart(sel(wj), 0))) & "</th>"
    Next wj
    html = html & "<th>Итого</th></tr></thead><tbody>"

    For i = LBound(dirs) To UBound(dirs)
        Dim dirName As String
        dirName = CStr(dirs(i))
        Dim a As Long
        For a = LBound(arms) To UBound(arms)
            Dim rowTotal As Double
            rowTotal = 0
            html = html & "<tr>"
            If a = LBound(arms) Then
                html = html & "<th rowspan='3' class='row-head'>" & Esc(dirName) & "</th>"
            End If
            html = html & "<td>" & Esc(arms(a)) & "</td>"
            For wj = cnt To 1 Step -1
                Dim wkStr As String
                wkStr = CStr(KeyPart(sel(wj), 0))
                Dim c As Double
                c = DictVal(counts, dirName & "|" & CStr(arms(a)) & "|" & wkStr & "|")
                rowTotal = rowTotal + c
                html = html & NumCell(CLng(c), J("direction", dirName) & "," & J("arm", CStr(arms(a))) & "," & J("status_week", wkStr))
            Next wj
            html = html & "<td class='total'" & DrillAttr(J("direction", dirName) & "," & J("arm", CStr(arms(a)))) & ">" & CStr(CLng(rowTotal)) & "</td></tr>"
        Next a

        ' Строка «% планшет» - по каждой неделе этой дирекции.
        html = html & "<tr class='pct-row'><td>% планшет</td>"
        Dim sumTot As Double, sumTab As Double
        ' (!) v6.1. Dim в VBA - объявление на этапе компиляции: переменная живёт всю
        ' процедуру и НЕ обнуляется на новой итерации цикла. Без явного сброса ДЭНТ
        ' получал накопленные суммы ДГМ (34,8% вместо 0,0%). Сброс обязателен.
        sumTot = 0: sumTab = 0
        For wj = cnt To 1 Step -1
            wkStr = CStr(KeyPart(sel(wj), 0))
            Dim key As String
            key = dirName & "|" & wkStr & "|"
            sumTot = sumTot + DictVal(total, key)
            sumTab = sumTab + DictVal(tablet, key)
            If DictVal(total, key) = 0 Then
                html = html & "<td class='pct empty'>—</td>"
            Else
                html = html & PctCell(SafePercent(tablet, total, key), J("direction", dirName) & "," & J("status_week", wkStr))
            End If
        Next wj
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot, J("direction", dirName))
        Else
            html = html & "<td class='pct empty'>—</td>"
        End If
        html = html & "</tr>"
    Next i

    html = html & "</tbody></table>"
    html = html & Recipe("Окно - последние " & CStr(windowSize) & " недель данных (yearWeek), " & _
        "не превосходящие отчётную неделю REPORT/WEEK; % планшет = ПЛАНШЕТ / (ПК + ПЛАНШЕТ) " & _
        "по каждой неделе и дирекции; пустые ячейки - нет событий. " & _
        "Зоны набора REPORT/SLIDE_ZONES перед фильтром нормализуются «+» -> «;».")
    BuildBlock1Table = html
End Function

' =====================================================================================
' 5. Блок 2 (п.5): Пост | Нарядов | Событий | % планшета | ДГМ | ДЭНТ | Оценка,
'    за отчётную неделю (REPORT/WEEK). Оценка: мало данных -> норма -> провал -> иначе без.
'    Валидация порогов: при ProvalForPlanshet >= NormaForPlanshet оценки «норма»/«провал»
'    не выводятся, в лог пишется предупреждение (§5.7).
' =====================================================================================
Private Sub EnsureBlock2Data()
    If mBlock2Ready Then Exit Sub
    EnsureSnapshot

    Dim rw As Long
    rw = ReportWeekValue()
    Dim fw As Variant
    fw = AppendFilter(FArm(), "yearWeek=" & CStr(rw))
    Dim fwt As Variant
    fwt = AppendFilter(FTablet(), "yearWeek=" & CStr(rw))

    Dim records As Object, events As Object, totDir As Object, tabDir As Object
    Set records = modAggregate.GroupCountDistinct(Array("postN"), "number", fw)
    Set events = modAggregate.GroupCount(Array("postN"), fw)
    Set totDir = modAggregate.GroupCount(Array("postN", "direction"), fw)
    Set tabDir = modAggregate.GroupCount(Array("postN", "direction"), fwt)
    Dim totPost As Object, tabPost As Object
    Set totPost = modAggregate.GroupCount(Array("postN"), fw)
    Set tabPost = modAggregate.GroupCount(Array("postN"), fwt)

    Dim posts As Object
    Set posts = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In records.Keys
        posts(KeyPart(k, 0)) = True
    Next k
    For Each k In events.Keys
        posts(KeyPart(k, 0)) = True
    Next k
    For Each k In totDir.Keys
        posts(KeyPart(k, 0)) = True
    Next k

    Set mBlock2 = CreateObject("Scripting.Dictionary")
    Set mBlock2("records") = records
    Set mBlock2("events") = events
    Set mBlock2("totDir") = totDir
    Set mBlock2("tabDir") = tabDir
    Set mBlock2("totPost") = totPost
    Set mBlock2("tabPost") = tabPost
    Set mBlock2("posts") = posts
    mBlock2Ready = True
End Sub

' Чтение порогов Блока 2 с валидацией (§5.7). Возвращает признак корректной конфигурации.
Private Function Block2Thresholds(ByRef minPost As Long, ByRef norma As Double, ByRef proval As Double) As Boolean
    minPost = CLng(modMain.GetVariableDef("REPORT/MIN_POST_RECORDS", "10"))
    norma = CDbl(Val(Replace(modMain.GetVariableDef("NormaForPlanshet", "90"), ",", ".")))
    proval = CDbl(Val(Replace(modMain.GetVariableDef("ProvalForPlanshet", "50"), ",", ".")))
    If proval >= norma Then
        modLog.WriteLogEntry Now, "Предупреждение", "Блок 2", "Variable", _
            "ProvalForPlanshet (" & CStr(proval) & ") >= NormaForPlanshet (" & CStr(norma) & _
            ") - некорректная конфигурация порогов: оценки «норма»/«провал» не выводятся."
        Block2Thresholds = False
        Exit Function
    End If
    Block2Thresholds = True
End Function

Private Function Block2Grade(records As Double, pct As Double, minPost As Long, _
                             norma As Double, proval As Double, thresholdsValid As Boolean) As String
    If records < minPost Then
        Block2Grade = "<span class='grade-muted'>мало данных</span>"
        Exit Function
    End If
    If Not thresholdsValid Then
        Block2Grade = "—"
        Exit Function
    End If
    Dim p As Double
    p = pct * 100
    If p >= norma Then
        Block2Grade = "<span class='grade-good'>норма</span>"
    ElseIf p < proval Then
        Block2Grade = "<span class='grade-bad'>провал</span>"
    Else
        Block2Grade = "—"
    End If
End Function

Private Function BuildBlock2Table() As String
    EnsureBlock2Data

    Dim records As Object
    Set records = mBlock2("records")
    Dim rw As Long
    rw = ReportWeekValue()

    If records.Count = 0 Then
        BuildBlock2Table = "<h3>Количество заказ-нарядов по постам (нед. " & Esc(WeekLabel(rw)) & ")</h3>" & EmptyTable()
        Exit Function
    End If

    Dim minPost As Long, norma As Double, proval As Double
    Dim thresholdsValid As Boolean
    thresholdsValid = Block2Thresholds(minPost, norma, proval)

    Dim events As Object, totDir As Object, tabDir As Object, totPost As Object, tabPost As Object
    Set events = mBlock2("events")
    Set totDir = mBlock2("totDir")
    Set tabDir = mBlock2("tabDir")
    Set totPost = mBlock2("totPost")
    Set tabPost = mBlock2("tabPost")

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(records, True)

    Dim html As String, i As Long
    html = "<h3>Количество заказ-нарядов по постам (нед. " & Esc(WeekLabel(rw)) & ")</h3>"
    html = html & "<table class='block-table'><thead><tr><th>Пост</th><th>Нарядов</th><th>Событий</th>" & _
        "<th>% планшета</th><th>ДГМ</th><th>ДЭНТ</th><th>Оценка</th></tr></thead><tbody>"

    For i = LBound(sorted) To UBound(sorted)
        Dim postName As String
        postName = KeyPart(CStr(sorted(i)), 0)
        Dim postTitle As String
        postTitle = postName
        If postTitle = "" Then postTitle = "(пост не указан)"

        Dim recs As Double, evs As Double
        recs = DictVal(records, postName & "|")
        evs = DictVal(events, postName & "|")
        Dim pctPost As Double
        pctPost = SafePercent(tabPost, totPost, postName & "|")
        Dim pctDGM As Double, pctDENT As Double
        pctDGM = SafePercent(tabDir, totDir, postName & "|ДГМ|")
        pctDENT = SafePercent(tabDir, totDir, postName & "|ДЭНТ|")

        Dim drillPost As String, drillDGM As String, drillDENT As String
        drillPost = J("postN", postName) & "," & J("status_week", CStr(rw))
        drillDGM = drillPost & "," & J("direction", "ДГМ")
        drillDENT = drillPost & "," & J("direction", "ДЭНТ")

        html = html & "<tr><td>" & Esc(postTitle) & "</td>" & _
            NumCell(CLng(recs), drillPost) & _
            NumCell(CLng(evs), drillPost) & _
            PctCell(pctPost, drillPost) & _
            PctCell(pctDGM, drillDGM) & _
            PctCell(pctDENT, drillDENT) & _
            "<td>" & Block2Grade(recs, pctPost, minPost, norma, proval, thresholdsValid) & "</td></tr>"
    Next i

    html = html & "</tbody></table>"
    html = html & Recipe("За отчётную неделю (REPORT/WEEK, по yearWeek статуса). Нарядов - уникальные number; " & _
        "Событий - строки подписаний; % планшета = ПЛАНШЕТ / (ПК + ПЛАНШЕТ); ДГМ/ДЭНТ - та же доля по дирекции. " & _
        "Оценка: норма при % >= NormaForPlanshet, провал при % < ProvalForPlanshet, " & _
        "«мало данных» при нарядов < REPORT/MIN_POST_RECORDS, остальное - без оценки (серая зона).")
    BuildBlock2Table = html
End Function

' График по Блоку 2 (п.11): столбики «% планшет» по постам, инлайн-SVG, без внешних библиотек.
Private Function BuildBlock2Chart() As String
    EnsureBlock2Data

    Dim records As Object
    Set records = mBlock2("records")
    If records.Count = 0 Then BuildBlock2Chart = "": Exit Function

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(records, True)
    Dim nCharts As Long
    nCharts = UBound(sorted) - LBound(sorted) + 1

    Dim totPost As Object, tabPost As Object
    Set totPost = mBlock2("totPost")
    Set tabPost = mBlock2("tabPost")

    Dim slot As Long
    slot = 76
    Dim w As Long
    w = 50 + nCharts * slot + 30
    Dim h As Long
    h = 300
    Dim baseY As Long
    baseY = 255
    Dim maxBar As Long
    maxBar = 220

    Dim svg As String
    svg = "<div class='chart-wrap'><svg viewBox='0 0 " & w & " " & h & "' xmlns='http://www.w3.org/2000/svg' style='max-width:100%;height:auto;'>"
    svg = svg & "<line x1='40' y1='" & baseY & "' x2='" & (w - 10) & "' y2='" & baseY & "' stroke='#cfe4f5' stroke-width='1'/>"
    svg = svg & "<text x='36' y='" & (baseY - maxBar + 4) & "' font-size='10' fill='#5a7793' text-anchor='end'>100%</text>"
    svg = svg & "<text x='36' y='" & (baseY + 4) & "' font-size='10' fill='#5a7793' text-anchor='end'>0%</text>"

    Dim i As Long
    For i = LBound(sorted) To UBound(sorted)
        Dim idx As Long
        idx = i - LBound(sorted)
        Dim postN As String
        postN = KeyPart(CStr(sorted(i)), 0)
        Dim pct As Double
        pct = SafePercent(tabPost, totPost, postN & "|")
        Dim barH As Long
        barH = CLng(pct * maxBar)
        If barH < 1 And pct > 0 Then barH = 1
        Dim color As String
        color = modColor.PercentToColor(pct, COLOR_BAD, COLOR_WARN, COLOR_GOOD)
        Dim x As Long
        x = 50 + idx * slot
        Dim y As Long
        y = baseY - barH

        Dim postLabel As String
        postLabel = postN
        If postLabel = "" Then postLabel = "не указан"
        If Len(postLabel) > 12 Then postLabel = Left$(postLabel, 11) & "…"

        svg = svg & "<rect x='" & x & "' y='" & y & "' width='44' height='" & barH & "' rx='3' fill='" & color & "'/>"
        If barH > 0 Then
            svg = svg & "<text x='" & (x + 22) & "' y='" & (y - 5) & "' font-size='11' fill='#10283e' text-anchor='middle'>" & FormatPct(pct) & "</text>"
        End If
        svg = svg & "<text x='" & (x + 22) & "' y='" & (baseY + 16) & "' font-size='10' fill='#5a7793' text-anchor='middle'>" & Esc(postLabel) & "</text>"
    Next i

    svg = svg & "</svg></div>"
    BuildBlock2Chart = "<h3>% планшет по постам - за отчётную неделю</h3>" & svg & _
        Recipe("Высота столбика - % планшет поста за отчётную неделю (те же данные, что в таблице Блока 2).")
End Function

' =====================================================================================
' 6. Блоки 4 и 5: строки direction/postN, столбцы - недели, значение «% планшет» (P1-3).
' =====================================================================================
Private Function BuildPctMatrixTable(rowCol As String, rowCaption As String) As String
    EnsureSnapshot

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FArm())
    Set tablet = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FTablet())

    Dim title As String
    If rowCol = "direction" Then title = "Динамика % планшет по дирекциям" Else title = "Динамика % планшет по постам"

    If total.Count = 0 Then
        BuildPctMatrixTable = "<h3>" & title & "</h3>" & EmptyTable()
        Exit Function
    End If

    Dim weeks As Variant, rows As Variant
    weeks = AxisFromKeys(total, 1, True)
    rows = AxisFromKeys(total, 0, False)

    Dim html As String, i As Long, wj As Long
    html = "<h3>" & title & "</h3>"
    html = html & "<table class='block-table matrix'><thead><tr><th>" & Esc(rowCaption) & "</th>"
    For wj = LBound(weeks) To UBound(weeks)
        html = html & "<th>" & Esc(WeekLabel(KeyPart(weeks(wj), 0))) & "</th>"
    Next wj
    html = html & "<th>Итого</th></tr></thead><tbody>"

    For i = LBound(rows) To UBound(rows)
        Dim rowName As String
        rowName = CStr(rows(i))
        Dim sumTot As Double, sumTab As Double
        sumTot = 0: sumTab = 0
        html = html & "<tr><th class='row-head'>" & Esc(rowName) & "</th>"
        For wj = LBound(weeks) To UBound(weeks)
            Dim key As String
            key = rowName & "|" & CStr(KeyPart(weeks(wj), 0)) & "|"
            sumTot = sumTot + DictVal(total, key)
            sumTab = sumTab + DictVal(tablet, key)
            If DictVal(total, key) = 0 Then
                html = html & "<td class='pct empty'>—</td>"
            Else
                html = html & PctCell(SafePercent(tablet, total, key), J(rowCol, rowName) & "," & J("status_week", CStr(KeyPart(weeks(wj), 0))))
            End If
        Next wj
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot, J(rowCol, rowName))
        Else
            html = html & "<td class='pct empty'>—</td>"
        End If
        html = html & "</tr>"
    Next i

    html = html & "</tbody></table>"
    html = html & Recipe("Доля событий с arm = ПЛАНШЕТ среди arm из {ПК, ПЛАНШЕТ} по " & _
        "строкам и неделям (yearWeek из данных); колонка «Итого» - за все недели снимка; «—» - нет событий.")
    BuildPctMatrixTable = html
End Function

' =====================================================================================
' 7. Блок 6 (п.6): понедельная раскладка ПН-3..ПН по сотрудникам и по подразделениям
'    (emp_dep), сортировка по «% планшет» ПН по убыванию; рейтинг - за отчётную неделю.
' =====================================================================================
Private Function BuildBlock6Weekly(direction As String, Optional byDept As Boolean = False) As String
    EnsureSnapshot

    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(4, cnt)
    If cnt = 0 Then
        BuildBlock6Weekly = "<h3>Блок 6 - " & Esc(direction) & "</h3>" & EmptyTable()
        Exit Function
    End If

    Dim html As String
    html = "<h3>Понедельная статистика - " & Esc(direction) & "</h3>"
    If byDept Then
        html = html & "<h4>По подразделениям</h4>" & Block6WeeklyCore(direction, "emp_dep", "Подразделение", "emp_dep", weeks, cnt)
    Else
        html = html & "<h4>По сотрудникам</h4>" & Block6WeeklyCore(direction, "employee", "Сотрудник", "employee", weeks, cnt)
    End If
    html = html & Recipe("Недели ПН (REPORT/WEEK), ПН-1, ПН-2, ПН-3 - только присутствующие в данных, " & _
        "не превосходящие отчётную неделю. % планшет = ПЛАНШЕТ / (ПК + ПЛАНШЕТ) за неделю; " & _
        "ВСЕГО ПОДПИСЕЙ - события (строки tbDATA) за неделю; из них на планшете - с arm = ПЛАНШЕТ; " & _
        "«Готов к приёмке»/«Готов к выбытию» - подписания с планшета с соответствующим статусом; " & _
        "Среднее время - средняя deltaHours по строкам «Готов к выбытию» (вывод «чч:мм»). " & _
        "Сортировка - по % планшет за ПН, по убыванию.")
    BuildBlock6Weekly = html
End Function

Private Function Block6WeeklyCore(direction As String, rowCol As String, rowCaption As String, _
                                  drillField As String, weeks As Variant, cnt As Long) As String
    ' 1. Объявляем ВСЕ словари (включая tabObj):
    Dim tot As Object, tabObj As Object, tabAcc As Object, tabLeave As Object
    Dim delSum As Object, delCnt As Object

    ' 2. Объявляем все остальные счётчики и текстовые переменные:
    Dim n As Long, r As Long, i As Long
    Dim armT As String, rowVal As String, yw As String, key As String, dv As Variant
    Dim html As String

    ' 3. Инициализируем словари:
    Set tot = CreateObject("Scripting.Dictionary")
    Set tabObj = CreateObject("Scripting.Dictionary")
    Set tabAcc = CreateObject("Scripting.Dictionary")
    Set tabLeave = CreateObject("Scripting.Dictionary")
    Set delSum = CreateObject("Scripting.Dictionary")
    Set delCnt = CreateObject("Scripting.Dictionary")

    n = modAggregate.RowCount()
    For r = 1 To n
        If IsTrueText(modAggregate.CellText(r, "in_bounds")) Then
            armT = modAggregate.CellText(r, "arm")
            If armT = "ПК" Or armT = "ПЛАНШЕТ" Then
                If modAggregate.CellText(r, "emp_dep") = direction Then
                    rowVal = modAggregate.CellText(r, rowCol)
                    If rowVal <> "" Then
                        yw = modAggregate.CellText(r, "yearWeek")
                        For i = 1 To cnt
                            If yw = CStr(KeyPart(weeks(i), 0)) Then
                                key = rowVal & "|" & yw & "|"
                                AddCnt tot, key, 1
                                If armT = "ПЛАНШЕТ" Then
                                    AddCnt tabObj, key, 1
                                    If IsStatusReadyToAccept(modAggregate.CellText(r, "ready_for")) Then AddCnt tabAcc, key, 1
                                    If IsStatusReadyToLeave(modAggregate.CellText(r, "ready_for")) Then AddCnt tabLeave, key, 1
                                End If
                                If IsStatusReadyToLeave(modAggregate.CellText(r, "ready_for")) Then
                                    dv = modAggregate.CellRaw(r, "deltaHours")
                                    If IsNumeric(dv) Then
                                        AddCnt delSum, key, CDbl(dv)
                                        AddCnt delCnt, key, 1
                                    End If
                                End If
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    Next r

    ' Проценты за ПН (weeks(1) - самая свежая из окна) для сортировки.
    Dim pctDict As Object
    Set pctDict = CreateObject("Scripting.Dictionary")
    Dim rowSet As Object
    Set rowSet = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In tot.Keys
        rowSet(KeyPart(k, 0)) = True
    Next k
    Dim wPN As String
    wPN = CStr(KeyPart(weeks(1), 0))
    Dim rv As Variant
    For Each rv In rowSet.Keys
        If DictVal(tot, CStr(rv) & "|" & wPN & "|") = 0 Then
            pctDict(rv) = -1
        Else
            pctDict(rv) = SafePercent(tabObj, tot, CStr(rv) & "|" & wPN & "|")
        End If
    Next rv

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(pctDict, True)

    html = "<table class='block-table weekly'><thead><tr><th rowspan='2'>" & Esc(rowCaption) & "</th>"
    For i = cnt To 1 Step -1
        Dim grp As String
        If i = 1 Then
            grp = "ПН (" & Esc(WeekLabel(KeyPart(weeks(i), 0))) & ")"
        Else
            grp = "ПН-" & CStr(i - 1) & " (" & Esc(WeekLabel(KeyPart(weeks(i), 0))) & ")"
        End If
        html = html & "<th colspan='6'>" & grp & "</th>"
    Next i
    html = html & "</tr><tr>"
    For i = cnt To 1 Step -1
        html = html & "<th>% планшет</th><th>ВСЕГО ПОДПИСЕЙ</th><th>ИЗ НИХ НА ПЛАНШЕТЕ</th>" & _
            "<th>из них Готов к приёмке</th><th>из них Готов к выбытию</th><th>Среднее время</th>"
    Next i
    html = html & "</tr></thead><tbody>"

    For i = LBound(sorted) To UBound(sorted)
        Dim nameVal As String
        nameVal = CStr(sorted(i))
        html = html & "<tr><th class='row-head'>" & Esc(nameVal) & "</th>"
        Dim wj As Long
        For wj = cnt To 1 Step -1
            Dim wkStr As String
            wkStr = CStr(KeyPart(weeks(wj), 0))
            Dim key2 As String
            key2 = nameVal & "|" & wkStr & "|"
            Dim baseDrill As String
            baseDrill = J(drillField, nameVal) & "," & J("status_week", wkStr)

            If DictVal(tot, key2) = 0 Then
                html = html & "<td class='pct empty'>—</td>"
            Else
                html = html & PctCell(SafePercent(tabObj, tot, key2), baseDrill)
            End If
            html = html & NumCell(CLng(DictVal(tot, key2)), baseDrill)
            html = html & NumCell(CLng(DictVal(tabObj, key2)), baseDrill & "," & J("arm", "ПЛАНШЕТ"))
            html = html & NumCell(CLng(DictVal(tabAcc, key2)), baseDrill & "," & J("arm", "ПЛАНШЕТ") & "," & J("norm_status", "готов к приемке"))
            html = html & NumCell(CLng(DictVal(tabLeave, key2)), baseDrill & "," & J("arm", "ПЛАНШЕТ") & "," & J("norm_status", "готов к выбытию"))
            If DictVal(delCnt, key2) > 0 Then
                html = html & "<td class='num'" & DrillAttr(baseDrill & "," & J("norm_status", "готов к выбытию")) & ">" & _
                    FormatHHMM(DictVal(delSum, key2) / DictVal(delCnt, key2)) & "</td>"
            Else
                html = html & "<td class='num'>—</td>"
            End If
        Next wj
        html = html & "</tr>"
    Next i

    Block6WeeklyCore = html & "</tbody></table>"
End Function

Private Sub AddCnt(d As Object, key As String, v As Double)
    If d.Exists(key) Then d(key) = CDbl(d(key)) + v Else d(key) = v
End Sub

' Рейтинг топ/антитоп за ОТЧЁТНУЮ неделю (REPORT/WEEK), порог REPORT/MIN_RECORDS сохраняется.
Private Function BuildBlock6Rating() As String
    Dim topN As Long
    topN = CLng(modMain.GetVariableDef("REPORT/TOPN", "10"))

    Dim html As String
    html = "<h3 style='color:var(--good);'>Топ по использованию планшетов (за отчётную неделю)</h3>" & _
        Block6RankedTable(topN, descending:=True) & _
        "<h3 style='color:var(--bad);'>Требуют внимания (за отчётную неделю)</h3>" & _
        Block6RankedTable(topN, descending:=False)
    html = html & Recipe("Сотрудники с числом событий за отчётную неделю не менее REPORT/MIN_RECORDS; " & _
        "сортировка по % планшет за REPORT/WEEK; «Средняя длительность» - средняя deltaHours по строкам " & _
        "«Готов к выбытию» сотрудника за неделю.")
    BuildBlock6Rating = html
End Function

Private Function Block6RankedTable(topN As Long, descending As Boolean) As String
    EnsureSnapshot

    Dim rw As Long
    rw = ReportWeekValue()
    Dim minRecords As Long
    minRecords = CLng(modMain.GetVariableDef("REPORT/MIN_RECORDS", "5"))

    Dim totalCounts As Object, tabletCounts As Object, avgDelta As Object
    Set totalCounts = modAggregate.GroupCount(Array("employee"), AppendFilter(FArm(), "yearWeek=" & CStr(rw)))
    Set tabletCounts = modAggregate.GroupCount(Array("employee"), AppendFilter(FTablet(), "yearWeek=" & CStr(rw)))
    ' P1-14: тот же базовый фильтр, что у соседних агрегатов, плюс срез по закрытым нарядам.
    Set avgDelta = modAggregate.GroupAverage(Array("employee"), "deltaHours", _
        Array("in_bounds=True", "arm@=ПК;ПЛАНШЕТ", "ready_for=Готов к выбытию", "yearWeek=" & CStr(rw)))

    Dim pctDict As Object
    Set pctDict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In totalCounts.Keys
        If CLng(totalCounts(k)) >= minRecords Then
            pctDict(k) = SafePercent(tabletCounts, totalCounts, CStr(k))
        End If
    Next k

    If pctDict.Count = 0 Then
        Block6RankedTable = "<p class='empty-note'>Недостаточно данных за отчётную неделю: нет сотрудников " & _
            "с числом событий не менее " & minRecords & " (Variable/REPORT/MIN_RECORDS).</p>"
        Exit Function
    End If

    Dim sortedKeys As Variant
    sortedKeys = modAggregate.SortDictionaryKeysByValue(pctDict, descending)

    Dim html As String, i As Long, shown As Long
    html = "<table class='block-table'><thead><tr><th>Инженер</th><th>% планшет</th>" & _
        "<th>Кол-во подписаний</th><th>Средняя длительность, ч</th></tr></thead><tbody>"

    For i = LBound(sortedKeys) To UBound(sortedKeys)
        If shown >= topN Then Exit For
        Dim empKey As String
        empKey = CStr(sortedKeys(i))
        Dim empName As String
        empName = KeyPart(empKey, 0)
        Dim deltaVal As String
        If avgDelta.Exists(empKey) Then deltaVal = Format(avgDelta(empKey), "0.0") Else deltaVal = "н/д"

        Dim drillEmp As String
        drillEmp = J("employee", empName) & "," & J("status_week", CStr(rw))

        html = html & "<tr><td>" & Esc(empName) & "</td>" & _
            PctCell(CDbl(pctDict(empKey)), drillEmp) & _
            NumCell(totalCounts(empKey), drillEmp) & _
            "<td class='num'" & DrillAttr(drillEmp & "," & J("norm_status", "готов к выбытию")) & ">" & deltaVal & "</td></tr>"
        shown = shown + 1
    Next i

    Block6RankedTable = html & "</tbody></table>"
End Function

' =====================================================================================
' 8. Блоки 7/8 - попарная разница status_date «Готов к выбытию» между ДГМ и ДЭНТ
'    по одному number. postN относится к заказ-наряду целиком, берётся из строки ДГМ.
' =====================================================================================
' Возвращает Dictionary: number -> Array(yearWeekДГМ, postNДГМ, deltaHours, dateДГМ, dateДЭНТ).
' Кэшируется (P2-3). Даты обеих дирекций нужны для топ-10 аномалий (п.7).
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
            result(k) = Array(dgmWeeks(k), dgmPosts(k), deltaH, dgmDates(k), dentDates(k))
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

' Топ-10 аномалий Блока 7 (п.7): наряды с максимальным расхождением status_date «Готов к
' выбытию» между ДГМ и ДЭНТ. Вывод: номер ЗН, пост, даты-время обеих дирекций, разница «чч:мм».
Private Function BuildSyncAnomalies() As String
    Dim pairs As Object
    Set pairs = BuildSyncPairs()
    If pairs.Count = 0 Then BuildSyncAnomalies = "": Exit Function

    Dim dd As Object
    Set dd = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In pairs.Keys
        dd(k) = CDbl(pairs(k)(2))
    Next k

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(dd, True)
    Dim top As Long
    top = 10
    If UBound(sorted) - LBound(sorted) + 1 < top Then top = UBound(sorted) - LBound(sorted) + 1

    Dim html As String, i As Long
    html = "<h3>10 аномалий - наряды с максимальным расхождением</h3>"
    html = html & "<table class='block-table'><thead><tr><th>Заказ-наряд</th><th>Пост</th>" & _
        "<th>ДГМ (Готов к выбытию)</th><th>ДЭНТ (Готов к выбытию)</th><th>Разница</th></tr></thead><tbody>"

    For i = LBound(sorted) To LBound(sorted) + top - 1
        Dim num As String
        num = CStr(sorted(i))
        Dim v As Variant
        v = pairs(num)
        Dim postT As String
        postT = CStr(v(1))
        If postT = "" Then postT = "(пост не указан)"
        html = html & "<tr>" & _
            "<td>" & Esc(num) & "</td>" & _
            "<td>" & Esc(postT) & "</td>" & _
            "<td class='num'>" & Esc(Format(CDate(v(3)), "dd.mm.yyyy hh:nn")) & "</td>" & _
            "<td class='num'>" & Esc(Format(CDate(v(4)), "dd.mm.yyyy hh:nn")) & "</td>" & _
            "<td class='num'" & DrillAttr(J("number", num)) & ">" & FormatHHMM(CDbl(v(2))) & "</td></tr>"
    Next i

    BuildSyncAnomalies = html & "</tbody></table>"
End Function

' HTML-таблица Блока 7 (по неделям, + аномалии) или Блока 8 (по неделям x постам).
Private Function BuildSyncTable(byPost As Boolean) As String
    Dim agg As Object
    Set agg = SyncAggregate(byPost)
    If agg.Count = 0 Then
        BuildSyncTable = "<p class='empty-note'>Нет пар «ДГМ &harr; ДЭНТ» по одному заказ-наряду " & _
            "для сопоставления.</p>"
        Exit Function
    End If

    Dim threshold As Double
    threshold = CDbl(modMain.GetVariableDef("REPORT/SYNC_THRESHOLD_MIN", "0")) / 60#

    Dim html As String
    If byPost Then
        html = "<h3>По неделям и постам</h3>"
    Else
        html = "<h3>По неделям</h3>"
    End If
    html = html & "<table class='block-table'><thead><tr><th>Неделя</th>"
    If byPost Then html = html & "<th>Пост</th>"
    html = html & "<th>Средняя разница, ч</th><th>Пар нарядов</th></tr></thead><tbody>"

    Dim k As Variant
    For Each k In agg.Keys
        Dim v As Variant
        v = agg(k)
        Dim cls As String
        cls = ""
        If threshold > 0 Then
            If CDbl(v(0)) > threshold Then cls = " class='warn'"
        End If
        Dim drillBase As String
        drillBase = J("status_week", CStr(KeyPart(k, 0)))
        If byPost Then drillBase = drillBase & "," & J("postN", CStr(KeyPart(k, 1)))
        html = html & "<tr><td>" & Esc(WeekLabel(KeyPart(k, 0))) & "</td>"
        If byPost Then html = html & "<td>" & Esc(KeyPart(k, 1)) & "</td>"
        html = html & "<td" & cls & ">" & Format(v(0), "0.0") & "</td>" & _
               "<td class='num'" & DrillAttr(drillBase & "," & J("norm_status", "готов к выбытию")) & ">" & CStr(v(1)) & "</td></tr>"
    Next k

    html = html & "</tbody></table>"

    If Not byPost Then html = html & BuildSyncAnomalies()

    html = html & Recipe("Пары ЗН со статусом «Готов к выбытию» в обеих дирекциях (ДГМ и ДЭНТ); " & _
        "значение - среднее модуля разницы status_date между дирекциями по одному number (в часах) и число пар. " & _
        "Строки с разницей выше REPORT/SYNC_THRESHOLD_MIN подсвечены. postN берётся от заказ-наряда целиком (из строки ДГМ).")
    BuildSyncTable = html
End Function

' =====================================================================================
' 9. Блок 9 и подблок 9а
' =====================================================================================
Private Function BuildBlock9Table() As String
    EnsureSnapshot

    Dim counts As Object
    Set counts = modAggregate.GroupCount(Array("zn_type", "defekt_type"), FBase())

    If counts.Count = 0 Then
        BuildBlock9Table = "<h3>Разбивка по видам ремонта и дефектов</h3>" & EmptyTable()
        Exit Function
    End If

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(counts, True)

    Dim html As String, i As Long
    html = "<h3>Разбивка по видам ремонта и дефектов</h3>"
    html = html & "<table class='block-table'><thead><tr><th>Вид ремонта</th><th>Тип дефекта</th>" & _
        "<th>Кол-во</th></tr></thead><tbody>"

    For i = LBound(sorted) To UBound(sorted)
        Dim k As String
        k = CStr(sorted(i))
        Dim znT As String, defT As String
        znT = KeyPart(k, 0)
        defT = KeyPart(k, 1)
        Dim defTitle As String
        defTitle = defT
        If defTitle = "" Then defTitle = "не указан"
        Dim drill9 As String
        drill9 = J("zn_type", znT)
        If defT <> "" Then drill9 = drill9 & "," & J("defekt_type", defT)
        html = html & "<tr><td>" & Esc(znT) & "</td><td>" & Esc(defTitle) & _
            "</td>" & NumCell(counts(k), drill9) & "</tr>"
    Next i

    html = html & "</tbody></table>"
    html = html & Recipe("Строки - вид ремонта (zn_type) со срезом по типу дефекта (defekt_type, " & _
        "пустое - «не указан»); значение - число строк tbDATA под базовым фильтром.")
    BuildBlock9Table = html
End Function

' Подблок 9а (п.8): строки - zn_type (+ defekt_type), колонки ДГМ/ДЭНТ, значение «% планшет».
Private Function BuildBlock9aTable() As String
    EnsureSnapshot

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array("zn_type", "defekt_type", "direction"), FArm())
    Set tablet = modAggregate.GroupCount(Array("zn_type", "defekt_type", "direction"), FTablet())

    If total.Count = 0 Then
        BuildBlock9aTable = "<h3>% планшет по видам ремонта в разрезе дирекций</h3>" & EmptyTable()
        Exit Function
    End If

    ' Строки - пары (zn_type, defekt_type), сортировка по убыванию суммарных событий.
    Dim rowSet As Object, rowSum As Object
    Set rowSet = CreateObject("Scripting.Dictionary")
    Set rowSum = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In total.Keys
        Dim rk As String
        rk = KeyPart(k, 0) & "|" & KeyPart(k, 1)
        If Not rowSet.Exists(rk) Then
            rowSet(rk) = True
            rowSum(rk) = CDbl(total(k))
        Else
            rowSum(rk) = CDbl(rowSum(rk)) + CDbl(total(k))
        End If
    Next k

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(rowSum, True)

    Dim html As String, i As Long
    html = "<h3>% планшет по видам ремонта в разрезе дирекций</h3>"
    html = html & "<table class='block-table'><thead><tr><th>Вид ремонта</th><th>Тип дефекта</th>" & _
        "<th>ДГМ</th><th>ДЭНТ</th></tr></thead><tbody>"

    For i = LBound(sorted) To UBound(sorted)
        Dim rowKey As String
        rowKey = CStr(sorted(i))
        Dim znT As String, defT As String
        znT = KeyPart(rowKey, 0)
        defT = KeyPart(rowKey, 1)
        Dim defTitle As String
        defTitle = defT
        If defTitle = "" Then defTitle = "не указан"
        Dim drillBase As String
        drillBase = J("zn_type", znT)
        If defT <> "" Then drillBase = drillBase & "," & J("defekt_type", defT)

        html = html & "<tr><td>" & Esc(znT) & "</td><td>" & Esc(defTitle) & "</td>"
        If DictVal(total, rowKey & "|ДГМ|") = 0 Then
            html = html & "<td class='pct empty'>—</td>"
        Else
            html = html & PctCell(SafePercent(tablet, total, rowKey & "|ДГМ|"), drillBase & "," & J("direction", "ДГМ"))
        End If
        If DictVal(total, rowKey & "|ДЭНТ|") = 0 Then
            html = html & "<td class='pct empty'>—</td>"
        Else
            html = html & PctCell(SafePercent(tablet, total, rowKey & "|ДЭНТ|"), drillBase & "," & J("direction", "ДЭНТ"))
        End If
        html = html & "</tr>"
    Next i

    html = html & "</tbody></table>"
    html = html & Recipe("% планшет по виду ремонта (zn_type, срез defekt_type) отдельно по каждой дирекции: " & _
        "ПЛАНШЕТ / (ПК + ПЛАНШЕТ); «—» - нет событий.")
    BuildBlock9aTable = html
End Function

' =====================================================================================
' 10. Дамп расшифровки (п.10): строки tbDATA за предыдущую и текущую недели (отбор по
'     yearWeek) с ФИО. Сериализация - JsonEscape (+ «<» -> \u003c), НЕ HtmlEscape (§5.4) -
'     иначе кавычки-сущности сломают JSON.parse в браузере.
' =====================================================================================
Private Function BuildDataDump() As String
    If mDumpReady Then BuildDataDump = mDumpJson: Exit Function

    EnsureSnapshot
    Dim weeks As Variant
    weeks = WeeksList()
    If UBound(weeks) < LBound(weeks) Then
        mDumpJson = "[]"
        mDumpReady = True
        BuildDataDump = mDumpJson
        Exit Function
    End If

    Dim rw As Long
    rw = ReportWeekValue()
    Dim lw As Long
    lw = LatestWeekValue()
    Dim include As Object
    Set include = CreateObject("Scripting.Dictionary")
    include(CStr(lw)) = True
    include(CStr(rw)) = True

    Dim n As Long
    n = modAggregate.RowCount()
    Dim s As String
    s = "["
    Dim first As Boolean
    first = True
    Dim r As Long
    For r = 1 To n
        Dim yw As String
        yw = modAggregate.CellText(r, "yearWeek")
        If include.Exists(yw) Then
            If Not first Then s = s & ","
            first = False
            s = s & DumpRowJson(r)
        End If
    Next r
    s = s & "]"

    ' «<» -> \u003c: внутри <script> последовательность «</...» не должна закрывать тег.
    mDumpJson = Replace(s, "<", "\u003c")
    mDumpReady = True
    BuildDataDump = mDumpJson
End Function

Private Function DumpRowJson(r As Long) As String
    Dim sd As Variant, dte As Variant
    sd = modAggregate.CellRaw(r, "status_date")
    dte = modAggregate.CellRaw(r, "date")
    Dim sdS As String, dteS As String
    sdS = DtToIso(sd)
    dteS = DtToIso(dte)
    Dim rf As String
    rf = modAggregate.CellText(r, "ready_for")

    DumpRowJson = "{" & _
        J("number", modAggregate.CellText(r, "number")) & "," & _
        J("date", dteS) & "," & _
        J("ready_for", rf) & "," & _
        J("direction", modAggregate.CellText(r, "direction")) & "," & _
        J("status_date", sdS) & "," & _
        J("arm", modAggregate.CellText(r, "arm")) & "," & _
        J("post", modAggregate.CellText(r, "post")) & "," & _
        J("postN", modAggregate.CellText(r, "postN")) & "," & _
        J("employee", modAggregate.CellText(r, "employee")) & "," & _
        J("date_week", CStr(WeekKeyOfSafe(dte))) & "," & _
        J("status_week", modAggregate.CellText(r, "yearWeek")) & "," & _
        J("date_year", CStr(YearOfSafe(dte))) & "," & _
        J("status_year", CStr(YearOfWeekKey(modAggregate.CellText(r, "yearWeek")))) & "," & _
        J("norm_status", NormStatus(rf)) & "}"
End Function

' Дата из снимка (число или текст) -> ISO-строка для JSON-дампа.
Private Function DtToIso(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then DtToIso = "": Exit Function
    If IsNumeric(v) Then
        Dim d As Date
        d = CDate(v)
        If d < DateSerial(1900, 1, 1) Or d > DateSerial(9999, 1, 1) Then DtToIso = "": Exit Function
        DtToIso = Format(d, "yyyy-mm-dd\Thh:nn:ss")
        Exit Function
    End If
    DtToIso = CStr(v)
End Function

' =====================================================================================
' 11. BuildPrompt - тело запроса к DeepSeek (task-for-coder п.9, Content Spec §7)
'     В промпт уходят ТОЛЬКО агрегаты по белому списку полей:
'     direction, postN, неделя, счётчики, доли, zn_type, defekt_type, часы расхождения,
'     число пар; Блок 6 - только маркеры [EMP_N] без ФИО.
'     employee и defect_desc не сериализуются ни в каком виде.
' =====================================================================================
Public Function BuildPrompt() As String
    EnsureSnapshot

    ' Системный промпт под 7 слайдов (spec.md §8.2, дословно).
    Const SYSTEM_PROMPT As String = _
        "Ты ведущий аналитик данных. Проанализируй предоставленные агрегированные метрики " & _
        "использования планшетов в МТО (доля планшетов по дирекциям, ремзонам и синхронность). " & _
        "Сформируй краткие бизнес-выводы (до 4 предложений на каждый) для 7-ми слайдов. Ищи аномалии. " & _
        "Не используй данные, которых нет во входном JSON. ФИО сотрудников во входных данных заменены " & _
        "маркерами вида [EMP_12] - не изменяй и не склоняй маркеры, в выводах ссылайся на сотрудников " & _
        "только ими. Ответ строго в формате JSON с ключами slide1_conclusions … slide7_conclusions, " & _
        "без markdown-разметки вокруг JSON."

    ' Блоки собираются в отдельные переменные (порядок вычислений прежний) -
    ' при DEBUG=2 их длины идут в лог для поиска раздутых частей промпта.
    Dim block4 As String, block5 As String, block7 As String
    Dim block8 As String, block9 As String, block6 As String
    block4 = PctMatrixToJson("direction")
    block5 = PctMatrixToJson("postN")
    block7 = SyncToJson(False)
    block8 = SyncToJson(True)
    block9 = Block9ToJson()
    block6 = Block6PeopleToJson()

    Dim userMessage As String
    userMessage = "{""block4_percent_by_direction"":" & block4 & _
                  ",""block5_percent_by_post"":" & block5 & _
                  ",""block7_sync_by_week"":" & block7 & _
                  ",""block8_sync_by_week_post"":" & block8 & _
                  ",""block9_defect_types"":" & block9 & _
                  ",""block6_people"":" & block6 & "}"

    Dim model As String
    model = modMain.GetVariable("AI/MODEL")

    ' response_format/temperature - контентные настройки запроса: просим провайдера
    ' гарантировать JSON-объект и снижаем вариативность формулировок.
    BuildPrompt = "{""model"":""" & JsonEscape(model) & """," & _
        """temperature"":0.2," & _
        """response_format"":{""type"":""json_object""}," & _
        """messages"":[" & _
        "{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
        "{""role"":""user"",""content"":""" & JsonEscape(userMessage) & """}]}"

    modLog.WriteDebug 1, "Формирование отчёта", "BuildPrompt", _
        "Модель: " & model & "; userMessage=" & Len(userMessage) & _
        " символов; весь запрос=" & Len(BuildPrompt)
    modLog.WriteDebug 2, "Формирование отчёта", "BuildPrompt", _
        "Блоки: block4=" & Len(block4) & "; block5=" & Len(block5) & _
        "; block7=" & Len(block7) & "; block8=" & Len(block8) & _
        "; block9=" & Len(block9) & "; block6=" & Len(block6)
End Function

' Матрица «строка x неделя» с долей планшета -> JSON-массив записей.
Private Function PctMatrixToJson(rowCol As String) As String
    EnsureSnapshot

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FArm())
    Set tablet = modAggregate.GroupCount(Array(rowCol, "yearWeek"), FTablet())

    Dim s As String
    s = "["
    Dim first As Boolean
    first = True
    Dim k As Variant
    For Each k In total.Keys
        If Not first Then s = s & ","
        first = False
        s = s & "{""row"":""" & JsonEscape(KeyPart(k, 0)) & """," & _
                """week"":""" & JsonEscape(WeekLabel(KeyPart(k, 1))) & """," & _
                """total"":" & CStr(DictVal(total, CStr(k))) & "," & _
                """tablet"":" & CStr(DictVal(tablet, CStr(k))) & "," & _
                """pct"":" & FmtJson(SafePercent(tablet, total, CStr(k))) & "}"
    Next k
    PctMatrixToJson = s & "]"
End Function

Private Function SyncToJson(byPost As Boolean) As String
    Dim agg As Object
    Set agg = SyncAggregate(byPost)

    Dim s As String
    s = "["
    Dim first As Boolean
    first = True
    Dim k As Variant
    For Each k In agg.Keys
        If Not first Then s = s & ","
        first = False
        Dim v As Variant
        v = agg(k) ' Array(avgHours, pairsCount)
        s = s & "{""week"":""" & JsonEscape(WeekLabel(KeyPart(k, 0))) & """"
        If byPost Then s = s & ",""post"":""" & JsonEscape(KeyPart(k, 1)) & """"
        s = s & ",""avg_hours"":" & FmtJson(CDbl(v(0))) & ",""pairs"":" & CStr(v(1)) & "}"
    Next k
    SyncToJson = s & "]"
End Function

Private Function Block9ToJson() As String
    EnsureSnapshot
    Dim counts As Object
    Set counts = modAggregate.GroupCount(Array("zn_type", "defekt_type"), FBase())

    Dim s As String
    s = "["
    Dim first As Boolean
    first = True
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

' Блок 6 в промпте - только маркеры [EMP_N] и агрегаты по людям (маркер, % планшет ПН/ПН-1,
' число подписаний). ФИО в промпт не попадает никогда (Content Spec §8).
Private Function Block6PeopleToJson() As String
    EnsureSnapshot
    Dim emps As Variant
    emps = EmpList()
    If UBound(emps) < LBound(emps) Then Block6PeopleToJson = "[]": Exit Function

    Dim rw As Long
    rw = ReportWeekValue()
    Dim prevW As String
    prevW = PrevWeekBefore(rw)

    Dim totals As Object, tablets As Object
    Set totals = modAggregate.GroupCount(Array("employee", "yearWeek"), FArm())
    Set tablets = modAggregate.GroupCount(Array("employee", "yearWeek"), FTablet())

    Dim s As String
    s = "["
    Dim first As Boolean
    first = True
    Dim i As Long
    For i = LBound(emps) To UBound(emps)
        Dim marker As String
        marker = "[EMP_" & CStr(i - LBound(emps) + 1) & "]"
        Dim pctCur As Double, pctPrev As Double
        pctCur = SafePercent(tablets, totals, CStr(emps(i)) & "|" & CStr(rw) & "|")
        pctPrev = 0
        If prevW <> "" Then pctPrev = SafePercent(tablets, totals, CStr(emps(i)) & "|" & prevW & "|")
        Dim signs As Double
        signs = DictVal(totals, CStr(emps(i)) & "|" & CStr(rw) & "|")
        If Not first Then s = s & ","
        first = False
        s = s & "{""emp"":""" & marker & """," & _
                """pct_cur"":" & FmtJson(pctCur) & "," & _
                """pct_prev"":" & FmtJson(pctPrev) & "," & _
                """signs"":" & CStr(CLng(signs)) & "}"
    Next i
    Block6PeopleToJson = s & "]"
End Function

' Число для JSON с ТОЧКОЙ независимо от локали (Format в русской локали даёт запятую).
Private Function FmtJson(x As Double) As String
    FmtJson = Replace$(Format(x, "0.000"), ",", ".")
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

' Единый алфавитный список всех employee снимка для псевдонимизации (§5.6). Кэш на прогон:
' один человек получает один номер во всех частях промпта и ответа.
Private Function EmpList() As Variant
    If mEmpReady Then EmpList = mEmpList: Exit Function
    EnsureSnapshot
    Dim d As Object
    Set d = modAggregate.DistinctValues("employee", FBase())
    mEmpList = modAggregate.SortKeys(d, False)
    mEmpReady = True
    EmpList = mEmpList
End Function

' =====================================================================================
' 12. ParseAIResponse - двухшаговый разбор ответа Chat Completions (P0-4).
'     Все 7 ключей slide1..slide7_conclusions складываются в module-level кэш mInsights;
'     через ByRef slide3/4/5 наружу возвращаются прежние 3 значения (контракт, §5.1).
' =====================================================================================
Public Function ParseAIResponse(responseText As String, ByRef slide3 As String, ByRef slide4 As String, ByRef slide5 As String) As Boolean
    slide3 = AI_FALLBACK: slide4 = AI_FALLBACK: slide5 = AI_FALLBACK

    If Trim$(responseText) = "" Then
        modLog.WriteDebug 1, "Формирование отчёта", "ParseAIResponse", _
            "Ответ ИИ пуст - все выводы заменяются заглушками"
        Set mInsights = Nothing
        mInsightsReady = False
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

    Dim ins As Object
    Set ins = CreateObject("Scripting.Dictionary")
    Dim ok As Boolean
    ok = True
    Dim missing As String
    missing = ""
    Dim i As Long, v As String
    For i = 1 To 7
        v = JsonUnescape(ExtractJsonStringValue(payload, "slide" & CStr(i) & "_conclusions"))
        If v <> "" Then
            ins("slide" & CStr(i)) = v
        Else
            ins("slide" & CStr(i)) = AI_FALLBACK
            ok = False
            missing = missing & "slide" & CStr(i) & ", "
        End If
    Next i

    slide3 = ins("slide3")
    slide4 = ins("slide4")
    slide5 = ins("slide5")

    If ok Then
        modLog.WriteDebug 1, "Формирование отчёта", "ParseAIResponse", _
            "Распознаны все 7 ключей slide1..slide7_conclusions"
    Else
        modLog.WriteDebug 1, "Формирование отчёта", "ParseAIResponse", _
            "Не распознаны: " & Left$(missing, Len(missing) - 2) & _
            "; длина payload=" & Len(payload)
    End If
    modLog.WriteDebug 2, "Формирование отчёта", "ParseAIResponse", _
        "Payload после распаковки: " & Left$(payload, 4000)

    Set mInsights = ins
    mInsightsReady = True
    ParseAIResponse = ok
End Function

' Обратная замена маркеров [EMP_N] -> ФИО, ТОЛЬКО в тексте ИИ-выводов, по УБЫВАНИЮ номеров
' (иначе [EMP_1] зацепит [EMP_12]). Таблица соответствия наружу не уходит (§5.6).
Private Function ReplaceEmpMarkers(text As String) As String
    If Not mEmpReady Then ReplaceEmpMarkers = text: Exit Function

    Dim i As Long
    For i = UBound(mEmpList) To LBound(mEmpList) Step -1
        text = Replace(text, "[EMP_" & CStr(i - LBound(mEmpList) + 1) & "]", CStr(mEmpList(i)))
    Next i
    ReplaceEmpMarkers = text
End Function

' Снимает обёртку ```json ... ``` , если модель всё-таки её добавила.
Private Function StripMarkdownFence(s As String) As String
    Dim t As String
    t = Trim$(s)
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
' 13. BuildPlaceholders - словарь {{ИМЯ}} -> значение для tmp_index.html
'     Сигнатура контракта не меняется (3 параметра): 7 выводов читаются из кэша mInsights
'     при mInsightsReady, иначе - fallback на параметры (путь DebugGenerateOffline, §5.1).
' =====================================================================================
Public Function BuildPlaceholders(aiSlide3 As String, aiSlide4 As String, aiSlide5 As String) As Object
    EnsureSnapshot

    Dim t0 As Single
    t0 = Timer

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' --- Слайд 1: два Дашборда + Блок 2 + график ---
    d("DASH_1_YTD") = BuildDashboard("ytd")
    d("DASH_1_PREV") = BuildDashboard("prev")
    d("BLOCK_2_PREV") = BuildBlock2Table()
    d("BLOCK_2_CHART") = BuildBlock2Chart()

    ' --- Слайды 2/3: Блок 1 (все ремзоны / набор SLIDE_ZONES) + Блок 6 ---
    Dim zonesRaw As String
    zonesRaw = Trim$(modMain.GetVariableDef("REPORT/SLIDE_ZONES", "СТК+ПРК"))
    d("BLOCK_1_DENT_ALL") = BuildBlock1Table("ДЭНТ", "")
    d("BLOCK_1_DGM_ALL") = BuildBlock1Table("ДГМ", "")
    If zonesRaw = "" Then
        ' Пустой REPORT/SLIDE_ZONES: второй экземпляр Блока 1 заменяется пояснением (ТЗ 3.2).
        d("BLOCK_1_DENT_ZONES") = "<p class='empty-note'>Набор ремзон не задан (Variable/REPORT/SLIDE_ZONES).</p>"
        d("BLOCK_1_DGM_ZONES") = "<p class='empty-note'>Набор ремзон не задан (Variable/REPORT/SLIDE_ZONES).</p>"
    Else
        d("BLOCK_1_DENT_ZONES") = BuildBlock1Table("ДЭНТ", zonesRaw)
        d("BLOCK_1_DGM_ZONES") = BuildBlock1Table("ДГМ", zonesRaw)
    End If
    d("BLOCK_6_DENT") = BuildBlock6Weekly("ДЭНТ", False)
    d("BLOCK_6_DGM") = BuildBlock6Weekly("ДГМ", False)
    d("BLOCK_6_DENT_DEPT") = BuildBlock6Weekly("ДЭНТ", True)
    d("BLOCK_6_DGM_DEPT") = BuildBlock6Weekly("ДГМ", True)

    ' --- Слайд 4: синхронность ---
    d("BLOCK_7_TABLE") = BuildSyncTable(False)
    d("BLOCK_8_TABLE") = BuildSyncTable(True)

    ' --- Слайд 5: дефекты и рейтинг ---
    d("BLOCK_9_TABLE") = BuildBlock9Table()
    d("BLOCK_9A_TABLE") = BuildBlock9aTable()
    d("BLOCK_6_RATING") = BuildBlock6Rating()

    ' --- Слайды 6/7: динамика ---
    d("BLOCK_4_TABLE") = BuildPctMatrixTable("direction", "Дирекция")
    d("BLOCK_5_TABLE") = BuildPctMatrixTable("postN", "Пост")

    ' --- Дамп расшифровки (п.10): JsonEscape, без HtmlEscape (§5.4) ---
    d("DATA_DUMP") = "<script type='application/json' id='drill-dump'>" & BuildDataDump() & "</script>"

    ' --- Выводы ИИ: из кэша при mInsightsReady, иначе fallback (1/2/6/7 - заглушка) ---
    Dim ai As Object
    Set ai = CreateObject("Scripting.Dictionary")
    Dim i As Long
    If mInsightsReady Then
        For i = 1 To 7
            ai("slide" & CStr(i)) = CStr(mInsights("slide" & CStr(i)))
        Next i
    Else
        ai("slide1") = AI_FALLBACK
        ai("slide2") = AI_FALLBACK
        ai("slide3") = aiSlide3
        ai("slide4") = aiSlide4
        ai("slide5") = aiSlide5
        ai("slide6") = AI_FALLBACK
        ai("slide7") = AI_FALLBACK
    End If
    For i = 1 To 7
        ' Обратная замена [EMP_N] -> ФИО (по убыванию номеров), затем HtmlEscape.
        d("AI_INSIGHT_SLIDE_" & CStr(i)) = Esc(ReplaceEmpMarkers(CStr(ai("slide" & CStr(i)))))
    Next i

    modLog.WriteDebug 1, "Формирование отчёта", "BuildPlaceholders", _
        "Готово: " & d.Count & " плейсхолдеров за " & Round(Timer - t0, 2) & " c"

    If modLog.GetDebugLevel() >= 2 Then
        Dim dbgInfo As String
        dbgInfo = ""
        Dim k As Variant
        For Each k In d.Keys
            If dbgInfo <> "" Then dbgInfo = dbgInfo & "; "
            dbgInfo = dbgInfo & k & "=" & Len(CStr(d(k)))
        Next k
        modLog.WriteDebug 2, "Формирование отчёта", "BuildPlaceholders", _
            "Размеры плейсхолдеров: " & dbgInfo
    End If

    Set BuildPlaceholders = d
End Function

' =====================================================================================
' 14. Статусы (п.13): единое правило сравнения через NormStatus.
' =====================================================================================
' v5 (C-1): распространённые написания булева значения. Дублирует приватную
' NormBoolLiteral из modAggregate сознательно - чтобы не править Core ради специфики МТО.
' v6 (P2-C): единое правило сравнения статуса: «ё» -> «е» + нижний регистр.
Private Function NormStatus(s As String) As String
    NormStatus = LCase$(Replace(Trim$(s), ChrW$(&H451), ChrW$(&H435)))  ' "ё" -> "е"
End Function

Private Function IsStatusReadyToLeave(s As String) As Boolean
    IsStatusReadyToLeave = (Left$(NormStatus(s), Len("готов к выбытию")) = "готов к выбытию")
End Function

Private Function IsStatusReadyToAccept(s As String) As Boolean
    IsStatusReadyToAccept = (Left$(NormStatus(s), Len("готов к приемке")) = "готов к приемке")
End Function

Private Function IsTrueText(s As String) As Boolean
    Select Case UCase$(Trim$(s))
        Case "TRUE", "ИСТИНА", "1", "ДА": IsTrueText = True
        Case Else: IsTrueText = False
    End Select
End Function
