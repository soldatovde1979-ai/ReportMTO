Attribute VB_Name = "modContentMTO"
' modContentMTO - CONTENT SPEC (МТО). Реализует 4 функции по контракту modMain.bas (Core):
'   BuildPivots, BuildPrompt, ParseAIResponse, BuildPlaceholders(s3, s4, s5).
'
' Версия 8.0 от 10.09.2026: переход на шаблон v4.0 (8 слайдов, 59 плейсхолдеров).
'   - модуль стал оркестратором: шапка, подвал, выводы ИИ и сборка словаря;
'     содержимое слайдов 1 и 5-8 отдаёт modContentZone, слайдов 2-4 - modContentDisc;
'   - BuildPlaceholders переписан под 59 ключей эталона MTO_макет_отчета_v4.0.html;
'   - ParseAIResponse читает slide1..slide8_conclusions, промпт описывает 8 слайдов
'     и получает сводку части «Техника» (ZoneFactsJson). Сигнатуры контракта Core
'     не менялись: BuildPivots, BuildPrompt, ParseAIResponse(s3,s4,s5),
'     BuildPlaceholders(s3,s4,s5);
'   - AiList: вывод ИИ оборачивается в <ul><li>, как в эталоне;
'   - PctBorderColor: шкала рамки процента 0/50/75/100 (Core modColor не менялся,
'     используется только InterpolateHex);
'   - WeeksList: запасной путь, если колонка yearWeek в книге мертва (нули) -
'     ось времени берётся из даты создания наряда, факт пишется в лог.
'   ВНИМАНИЕ: построители слайдов 1-4 версий <= 7.4 (BuildKpiOverview,
'   BuildWeeksTable, BuildPostsChart, BuildPeopleWeekly, BuildSignStat,
'   BuildUnsigned*, BuildDashboard и их помощники) больше НИКЕМ не вызываются.
'   Не удалены намеренно - решение об удалении за автором проекта.
'
' Версия 7.4 от 10.09.2026:
'   P0-1 - символы вне ANSI-1251 вынесены в ChrW$: стрелки дельты, минус, знак
'          умножения, стрелка и знак принадлежности. При импорте .bas в VBE
'          (UTF-8 -> 1251) они молча превращались в '?', и в отчёте вместо
'          «(треугольник) 5 к пр. нед.» печаталось «? 5 к пр. нед.».
'          Проверка: tools/vba_lint_v1.0/vba_lint.py.
'   P0-2 - OverviewToJson больше не падает, когда в tbDATA нет столбца dateWeek
'          (данные загружены до M v7). Раньше Err -2147221502 обрывал ВЕСЬ запрос
'          к ИИ и все четыре вывода уходили в заглушки - см. журнал 09.09 8:52 и
'          10.09 0:19. Теперь opened/unsigned отдаются как null, промпт собирается.
'   P0-3 - помощники Esc, FmtInt, FmtPct, FormatHHMM, CalcNote, PctCell, EmptyNote,
'          WeekLabel, RecentWeeksUpTo сделаны Public: их переиспользует
'          modContentZone (часть «Техника», слайды 5-8) вместо дублирования.
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
' Версия 7.3 от 09.09.2026 - ТЗ v1.2 (переход на 4 слайда), задачи T3-T9:
'   T3 - фильтры: FBase() пуст (in_bounds и arm не фильтруются, P0-1), добавлены
'        FSigned/FUnsigned/FDir; псевдонимы сотрудников «Сотрудник N» (BuildEmployeeAliases/
'        AliasOf/DeAlias) вместо маркеров [EMP_N] в новых блоках;
'   T4 - PctCell переведена с заливки на РАМКУ (modColor.PercentToColor -> border-color),
'        добавлены SVG-примитивы SvgBarsV/SvgBarsH/SvgBarsLine/SvgSpark, KpiTile,
'        FmtInt/FmtPct («чч:мм» - существующая FormatHHMM);
'   T5-T7 - блоки слайдов 1-4; T9 - промпт, разбор и плейсхолдеры под 4 слайда.
'   Старые функции блоков 1-9 остаются в коде, но из BuildPlaceholders не вызываются
'   (постановка §8: снятое из презентации не удаляется из кода).
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

Private mAlias As Object          ' кэш псевдонимов ФИО -> «Сотрудник N» (ТЗ v1.2, T3.3)
Private mAliasReady As Boolean

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

Private mSignData As Object       ' кэш статистики подписания: dir -> Dictionary (ТЗ v1.2, T6)
Private mSignReady As Object      ' dir -> True

Private mUnsigned As Object       ' кэш прохода по снимку для слайда 4 (ТЗ v1.2, T7)
Private mUnsignedReady As Boolean

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
    modContentZone.ResetZone
    modContentDisc.ResetDisc

    mEmpReady = False
    If IsArray(mEmpList) Then Erase mEmpList

    mAliasReady = False
    Set mAlias = Nothing

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

    Set mSignData = Nothing
    Set mSignReady = Nothing

    mUnsignedReady = False
    Set mUnsigned = Nothing
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

' --- Наборы фильтров (ТЗ v1.2, T3.2) ---
' P0-1 (решение 09.09.2026): in_bounds не фильтруется - поле равно ЛОЖЬ во всех строках
' выгрузки. Базовый фильтр FBase стал ПУСТЫМ; строка проходит любой фильтр, если не задан.
' (!) arm принимает три значения: "ПК", "ПЛАНШЕТ" и "НЕ ПОДПИСАНО". Подписанные события
' отбираются белым списком arm@=ПК;ПЛАНШЕТ (отсечение пустых "arm<>" пропускало
' "НЕ ПОДПИСАНО" и завышало знаменатель «% планшет»).
' FArm оставлен как синоним FSigned - старые функции блоков 1-9 ссылаются на него и
' остаются компилируемыми (из BuildPlaceholders они больше не вызываются, постановка §8).
Private Function FBase() As Variant
    FBase = Empty
End Function

Private Function FSigned() As Variant
    FSigned = Array("arm@=ПК;ПЛАНШЕТ")
End Function

Private Function FArm() As Variant
    FArm = FSigned()
End Function

Private Function FTablet() As Variant
    FTablet = Array("arm=ПЛАНШЕТ")
End Function

Private Function FUnsigned() As Variant
    FUnsigned = Array("arm=НЕ ПОДПИСАНО")
End Function

Private Function FDir(ByVal dir As String) As Variant
    FDir = Array("arm@=ПК;ПЛАНШЕТ", "direction=" & dir)
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
Public Function WeekLabel(yw As Variant) As String
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

Public Function Esc(s As Variant) As String
    Esc = modHTMLEngine.HtmlEscape(CStr(s))
End Function

' Ячейка «% планшет» (ТЗ v1.2, T4.1): цвет шкалы уходит в РАМКУ числа, фон ячейки - фон темы.
' hasValue=False -> <td class="pct empty">—</td> без рамки. Клик-расшифровка на процентах
' снята вместе с data-drill (§2.7 постановки: расшифровка - вне этой итерации).
Public Function PctCell(pct As Double, Optional hasValue As Boolean = True) As String
    If Not hasValue Then
        PctCell = "<td class='pct empty'>—</td>"
        Exit Function
    End If
    Dim color As String
    color = modColor.PercentToColor(pct, COLOR_BAD, COLOR_WARN, COLOR_GOOD)
    PctCell = "<td class='pct'><span style='border-color:" & color & "'>" & FmtPct(pct) & "</span></td>"
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

    ' Запасной путь: старый Power Query писал в yearWeek ноль на всех строках,
    ' и ось времени умирала молча. Если в колонке нет ни одной осмысленной недели,
    ' берём недели, посчитанные из даты создания наряда, и пишем это в лог.
    If Not WeeksUsable(mWeeks) Then
        mWeeks = modContentZone.WeeksFromDate()
        modLog.WriteDebug 1, "Формирование отчёта", "WeeksList", _
            "Колонка yearWeek не содержит недель (устаревший Power Query). " & _
            "Ось времени взята из поля date: недель " & CStr(ArrLen(mWeeks)) & "."
    End If

    mWeeksReady = True
    WeeksList = mWeeks
End Function

' Список недель годен, если в нём есть хотя бы одно значение больше нуля.
Private Function WeeksUsable(ByVal weeks As Variant) As Boolean
    WeeksUsable = False
    If ArrLen(weeks) = 0 Then Exit Function
    Dim i As Long
    For i = LBound(weeks) To UBound(weeks)
        If Val(KeyPart(weeks(i), 0)) > 0 Then
            WeeksUsable = True
            Exit Function
        End If
    Next i
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
Public Function RecentWeeksUpTo(limit As Long, ByRef count As Long) As Variant
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
Public Function FormatHHMM(hours As Double) As String
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
                html = html & PctCell(SafePercent(tablet, total, key))
            End If
        Next wj
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot)
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
            PctCell(pctPost) & _
            PctCell(pctDGM) & _
            PctCell(pctDENT) & _
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
                html = html & PctCell(SafePercent(tablet, total, key))
            End If
        Next wj
        If sumTot > 0 Then
            html = html & PctCell(sumTab / sumTot)
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
                html = html & PctCell(SafePercent(tabObj, tot, key2))
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
            PctCell(CDbl(pctDict(empKey))) & _
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
            html = html & PctCell(SafePercent(tablet, total, rowKey & "|ДГМ|"))
        End If
        If DictVal(total, rowKey & "|ДЭНТ|") = 0 Then
            html = html & "<td class='pct empty'>—</td>"
        Else
            html = html & PctCell(SafePercent(tablet, total, rowKey & "|ДЭНТ|"))
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

    ' (!) Построчная конкатенация s = s & ... на десятках тысяч строк дампа давала
    ' O(N^2) копирований растущей строки: этап BuildPlaceholders занимал десятки минут
    ' на 136k строк tbDATA. Двухпроходный сбор в массив строк + Join (паттерн RecentWeeksUpTo).
    Dim n As Long
    n = modAggregate.RowCount()
    Dim cnt As Long
    cnt = 0
    Dim r As Long
    Dim yw As String
    For r = 1 To n
        yw = modAggregate.CellText(r, "yearWeek")
        If include.Exists(yw) Then cnt = cnt + 1
    Next r

    Dim parts() As String
    ReDim parts(0 To cnt - 1)
    Dim p As Long
    p = 0
    For r = 1 To n
        yw = modAggregate.CellText(r, "yearWeek")
        If include.Exists(yw) Then
            parts(p) = DumpRowJson(r)
            p = p + 1
        End If
    Next r

    ' «<» -> \u003c: внутри <script> последовательность «</...» не должна закрывать тег.
    mDumpJson = Replace("[" & Join(parts, ",") & "]", "<", "\u003c")
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

    ' Системный промпт под 8 слайдов: две части, ИИ только интерпретирует готовые числа.
    Dim SYSTEM_PROMPT As String
    SYSTEM_PROMPT = _
        "Ты ведущий аналитик данных. Отчёт состоит из двух частей: слайды 1-4 - дисциплина " & _
        "подписания на планшете (обзор недели, дирекции ДЭНТ и ДГМ, неподписанные наряды), " & _
        "слайды 5-8 - операционка ремзоны (парк и заезды, что ломается и что возвращается, " & _
        "фазы наряда и хвост незакрытого, материалы и качество учёта). Сформируй краткие " & _
        "бизнес-выводы (2-3 пункта, каждый одним предложением) для каждого из 8 слайдов. "
    SYSTEM_PROMPT = SYSTEM_PROMPT & _
        "Ничего не вычисляй сам и не делай прогнозов: все числа уже посчитаны, твоя работа - " & _
        "их интерпретация. Не оценивай людей. Не используй данных, которых нет во входном JSON. " & _
        "ФИО сотрудников во входных данных заменены псевдонимами вида «Сотрудник 7» - не изменяй " & _
        "и не склоняй псевдонимы, ссылайся на сотрудников только ими. Ответ строго в формате JSON " & _
        "с ключами slide1_conclusions ... slide8_conclusions, без markdown-разметки вокруг JSON."

    ' Блоки собираются отдельно - при DEBUG=2 их длины идут в лог.
    Dim slide1 As String, slide2 As String, slide3 As String, slide4 As String
    slide1 = OverviewToJson()
    slide2 = DirToJson("ДЭНТ")
    slide3 = DirToJson("ДГМ")
    slide4 = UnsignedToJson()

    Dim userMessage As String
    Dim zoneFacts As String
    zoneFacts = modContentZone.ZoneFactsJson()
    userMessage = "{""slide1_overview"":" & slide1 & _
                  ",""slide2_dent"":" & slide2 & _
                  ",""slide3_dgm"":" & slide3 & _
                  ",""slide4_unsigned"":" & slide4 & _
                  ",""part_b_zone"":" & zoneFacts & "}"

    Dim model As String
    model = modMain.GetVariable("AI/MODEL")

    ' response_format/temperature - контентные настройки запроса.
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
        "Блоки: slide1=" & Len(slide1) & "; slide2=" & Len(slide2) & _
        "; slide3=" & Len(slide3) & "; slide4=" & Len(slide4)
End Function

' Целое число для JSON без локали.
Private Function JInt(v As Double) As String
    JInt = CStr(CLng(v))
End Function

' Слайд 1 в промпт: kpi + корзины времени + zn_type + defekt_type + «без поста» по неделям.
Private Function OverviewToJson() As String
    EnsureSnapshot
    Dim rw As Long
    rw = ReportWeekValue()
    Dim rwS As String
    rwS = CStr(rw)

    ' P0-2 (10.09.2026): столбец dateWeek появляется только в M v7. Если tbDATA
    ' собрана прежней версией запроса, столбца нет - и обращение к нему роняло
    ' ВЕСЬ запрос к ИИ (Err -2147221502), после чего все четыре вывода уходили
    ' в заглушки. Подтверждено журналом прогонов 09.09 8:52 и 10.09 0:19.
    ' Теперь метрики по неделе СОЗДАНИЯ наряда отдаются как null, остальной
    ' промпт собирается штатно. Полностью восстанавливается загрузкой M v7.
    Dim hasDW As Boolean
    hasDW = modAggregate.HasColumn("dateWeek")
    If Not hasDW Then
        modLog.WriteDebug 1, "Формирование отчёта", "OverviewToJson", _
            "Столбца dateWeek нет в tbDATA (данные загружены до M v7): " & _
            "opened/unsigned уходят в промпт как null, остальные метрики считаются."
    End If

    Dim opened As Double, closedN As Double
    If hasDW Then
        opened = DictVal(modAggregate.GroupCountDistinct(Array("dateWeek"), "number", FBase()), rwS & "|")
    End If
    closedN = DictVal(modAggregate.GroupCountDistinct(Array("yearWeek"), "number", _
        Array("ready_for=Готов к выбытию")), rwS & "|")

    Dim hasMed As Boolean
    Dim med As Double
    med = modAggregate.Percentile("deltaHours", 0.5, Array("yearWeek=" & rwS, "ready_for=Готов к выбытию"), hasMed)

    Dim armW As Object
    Set armW = modAggregate.GroupCount(Array("arm"), Array("yearWeek=" & rwS))
    Dim tabEv As Double, pcEv As Double
    tabEv = DictVal(armW, "ПЛАНШЕТ|")
    pcEv = DictVal(armW, "ПК|")

    Dim allD As Object, armD As Object
    Dim allEv As Double, unsEv As Double
    If hasDW Then
        Set allD = modAggregate.GroupCount(Array("dateWeek"), FBase())
        Set armD = modAggregate.GroupCount(Array("arm"), Array("dateWeek=" & rwS))
        allEv = DictVal(allD, rwS & "|")
        unsEv = DictVal(armD, "НЕ ПОДПИСАНО|")
    End If

    Dim kpi As String
    kpi = "{""opened"":" & IIf(hasDW, JInt(opened), "null") & ",""closed"":" & JInt(closedN) & "," & _
        """median_hours"":" & IIf(hasMed, FmtJson(med), "null") & "," & _
        """tablet_pct"":" & FmtJson(SafePctTwo(tabEv + pcEv, tabEv)) & "," & _
        """unsigned"":" & IIf(hasDW, JInt(unsEv), "null") & "," & _
        """unsigned_pct"":" & IIf(hasDW, FmtJson(SafePctTwo(allEv, unsEv)), "null") & "}"

    ' Корзины времени - те же границы, что BuildTimeHistogram (1/4/8/24/72 ч), обрезка p99.
    Dim buckets As String
    buckets = TimeBucketsToJson()

    Dim zn As Object, defk As Object
    Set zn = modAggregate.GroupCountDistinct(Array("zn_type"), "number", FBase())
    Set defk = modAggregate.GroupCountDistinct(Array("defekt_type"), "number", FBase())
    Dim sZn As String, sDef As String, first As Boolean, k As Variant
    sZn = "[": sDef = "["
    first = True
    For Each k In zn.Keys
        If Not first Then sZn = sZn & ","
        first = False
        Dim zt As String
        zt = KeyPart(k, 0)
        If zt = "" Then zt = "(не указан)"
        sZn = sZn & "{""zn_type"":""" & JsonEscape(zt) & """,""orders"":" & JInt(CDbl(zn(k))) & "}"
    Next k
    first = True
    For Each k In defk.Keys
        If Not first Then sDef = sDef & ","
        first = False
        Dim dt As String
        dt = KeyPart(k, 0)
        If dt = "" Then dt = "(раздел не указан)"
        sDef = sDef & "{""defekt_type"":""" & JsonEscape(dt) & """,""orders"":" & JInt(CDbl(defk(k))) & "}"
    Next k
    sZn = sZn & "]": sDef = sDef & "]"

    ' «Без поста» по неделям окна.
    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))
    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(windowSize, cnt)
    Dim totW As Object, npW As Object
    Set totW = modAggregate.GroupCountDistinct(Array("yearWeek"), "number", FSigned())
    Set npW = modAggregate.GroupCountDistinct(Array("yearWeek"), "number", AppendFilter(FSigned(), "postN="))
    Dim sNp As String, i As Long
    sNp = "["
    first = True
    If cnt > 0 Then
        For i = cnt To 1 Step -1
            Dim wkStr As String
            wkStr = CStr(KeyPart(weeks(i), 0))
            If Not first Then sNp = sNp & ","
            first = False
            sNp = sNp & "{""week"":""" & JsonEscape(WeekLabel(KeyPart(weeks(i), 0))) & """," & _
                """orders"":" & JInt(DictVal(totW, wkStr & "|")) & "," & _
                """nopost_pct"":" & FmtJson(SafePercent(npW, totW, wkStr & "|")) & "}"
        Next i
    End If
    sNp = sNp & "]"

    OverviewToJson = "{""kpi"":" & kpi & ",""time_buckets"":" & buckets & _
        ",""zn_types"":" & sZn & ",""defekt_sections"":" & sDef & ",""nopost_weekly"":" & sNp & "}"
End Function

' Корзины времени в ремзоне для промпта (границы совпадают с BuildTimeHistogram).
Private Function TimeBucketsToJson() As String
    EnsureSnapshot
    Dim has As Boolean
    Dim p99 As Double
    p99 = modAggregate.Percentile("deltaHours", 0.99, Empty, has)
    If Not has Then TimeBucketsToJson = "[]": Exit Function

    Dim vals(1 To 6) As Double, i As Long
    For i = 1 To 6: vals(i) = 0: Next i
    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    For r = 1 To n
        Dim v As Variant
        v = modAggregate.CellRaw(r, "deltaHours")
        If IsNumeric(v) Then
            If CDbl(v) > p99 Then
            ElseIf CDbl(v) < 1 Then vals(1) = vals(1) + 1
            ElseIf CDbl(v) < 4 Then vals(2) = vals(2) + 1
            ElseIf CDbl(v) < 8 Then vals(3) = vals(3) + 1
            ElseIf CDbl(v) < 24 Then vals(4) = vals(4) + 1
            ElseIf CDbl(v) < 72 Then vals(5) = vals(5) + 1
            Else: vals(6) = vals(6) + 1
            End If
        End If
    Next r

    Dim names() As String
    ReDim names(1 To 6)
    names(1) = "< 1 ч": names(2) = "1–4 ч": names(3) = "4–8 ч"
    names(4) = "8–24 ч": names(5) = "1–3 сут": names(6) = "> 3 сут"
    Dim s As String
    s = "["
    For i = 1 To 6
        If i > 1 Then s = s & ","
        s = s & "{""bucket"":""" & JsonEscape(names(i)) & """,""pairs"":" & JInt(vals(i)) & "}"
    Next i
    TimeBucketsToJson = s & "]"
End Function

' Слайды 2-3 в промпт: недели, посты, люди (только псевдонимы), статистика подписания.
Private Function DirToJson(dir As String) As String
    EnsureSnapshot
    Dim f As Variant
    f = FDir(dir)

    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))
    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(windowSize, cnt)

    Dim totalW As Object, tabW As Object
    Set totalW = modAggregate.GroupCount(Array("yearWeek"), f)
    Set tabW = modAggregate.GroupCount(Array("yearWeek"), AppendFilter(FTablet(), "direction=" & dir))
    Dim sW As String, i As Long, first As Boolean
    sW = "["
    first = True
    If cnt > 0 Then
        For i = cnt To 1 Step -1
            Dim wkStr As String
            wkStr = CStr(KeyPart(weeks(i), 0))
            If Not first Then sW = sW & ","
            first = False
            Dim totV As Double, tabV As Double
            totV = DictVal(totalW, wkStr & "|")
            tabV = DictVal(tabW, wkStr & "|")
            sW = sW & "{""week"":""" & JsonEscape(WeekLabel(KeyPart(weeks(i), 0))) & """," & _
                """events"":" & JInt(totV) & ",""tablet"":" & JInt(tabV) & "," & _
                """pct"":" & FmtJson(SafePctTwo(totV, tabV)) & "}"
        Next i
    End If
    sW = sW & "]"

    ' Посты за отчётную неделю (единица счёта - наряд).
    Dim rw As Long
    rw = ReportWeekValue()
    Dim fw As Variant
    fw = AppendFilter(f, "yearWeek=" & CStr(rw))
    Dim rec As Object, tabRec As Object
    Set rec = modAggregate.GroupCountDistinct(Array("postN"), "number", fw)
    Set tabRec = modAggregate.GroupCountDistinct(Array("postN"), "number", _
        AppendFilter(AppendFilter(FTablet(), "direction=" & dir), "yearWeek=" & CStr(rw)))
    Dim sP As String, k As Variant
    sP = "["
    first = True
    For Each k In rec.Keys
        If Not first Then sP = sP & ","
        first = False
        Dim pn As String
        pn = KeyPart(k, 0)
        If pn = "" Then pn = "(пост не указан)"
        sP = sP & "{""post"":""" & JsonEscape(pn) & """,""records"":" & JInt(CDbl(rec(k))) & "," & _
            """pct"":" & FmtJson(SafePercent(tabRec, rec, CStr(k))) & "}"
    Next k
    sP = sP & "]"

    ' Люди: только псевдонимы «Сотрудник N», понедельные агрегаты за ПН-3..ПН.
    Dim pw As Variant, pc As Long
    pw = RecentWeeksUpTo(4, pc)
    Dim totP As Object, tabP As Object
    Set totP = modAggregate.GroupCountDistinct(Array("employee", "yearWeek"), "Key", f)
    Set tabP = modAggregate.GroupCountDistinct(Array("employee", "yearWeek"), "Key", _
        AppendFilter(f, "arm=ПЛАНШЕТ"))
    Dim empSet As Object
    Set empSet = CreateObject("Scripting.Dictionary")
    For Each k In totP.Keys
        Dim wi As Long
        For wi = 1 To pc
            If KeyPart(k, 1) = CStr(KeyPart(pw(wi), 0)) Then
                empSet(KeyPart(k, 0)) = True
                Exit For
            End If
        Next wi
    Next k
    Dim sPe As String, e As Variant
    sPe = "["
    first = True
    For Each e In empSet.Keys
        If Not first Then sPe = sPe & ","
        first = False
        sPe = sPe & "{""alias"":""" & JsonEscape(AliasOf(CStr(e))) & """,""weeks"":["
        Dim wj As Long, firstW As Boolean
        firstW = True
        For wj = pc To 1 Step -1
            Dim keyP As String
            keyP = CStr(e) & "|" & CStr(KeyPart(pw(wj), 0)) & "|"
            If Not firstW Then sPe = sPe & ","
            firstW = False
            Dim totE As Double, tabE As Double
            totE = DictVal(totP, keyP)
            tabE = DictVal(tabP, keyP)
            sPe = sPe & "{""week"":""" & JsonEscape(WeekLabel(KeyPart(pw(wj), 0))) & """," & _
                """pct"":" & FmtJson(SafePctTwo(totE, tabE)) & ",""signs"":" & JInt(totE) & "," & _
                """tablet"":" & JInt(tabE) & "}"
        Next wj
        sPe = sPe & "]}"
    Next e
    sPe = sPe & "]"

    ' Статистика подписания.
    EnsureSignData dir
    Dim sd As Object
    Set sd = mSignData(dir)
    Dim both As Double, accOnly As Double, leaOnly As Double, none As Double
    both = 0: accOnly = 0: leaOnly = 0: none = 0
    For Each k In sd("all").Keys
        Dim hA As Boolean, hL As Boolean
        hA = sd("acc").Exists(k)
        hL = sd("lea").Exists(k)
        If hA And hL Then
            both = both + 1
        ElseIf hA Then
            accOnly = accOnly + 1
        ElseIf hL Then
            leaOnly = leaOnly + 1
        Else
            none = none + 1
        End If
    Next k
    Dim sS As String
    sS = "{""total"":" & JInt(CDbl(sd("all").Count)) & ",""both"":" & JInt(both) & _
        ",""accept_only"":" & JInt(accOnly) & ",""leave_only"":" & JInt(leaOnly) & _
        ",""none"":" & JInt(none) & "}"

    DirToJson = "{""weeks"":" & sW & ",""posts"":" & sP & ",""people"":" & sPe & ",""sign_stat"":" & sS & "}"
End Function

' Слайд 4 в промпт: KPI, старение, вид ремонта неподписанных.
Private Function UnsignedToJson() As String
    EnsureUnsignedData
    Dim d As Object
    Set d = mUnsigned

    Dim none As Double, part As Double, older As Double
    none = 0: part = 0: older = 0
    Dim k As Variant
    For Each k In d("all").Keys
        Dim s As Double, rows As Double
        s = DictVal(d("signed"), CStr(k))
        rows = DictVal(d("rows"), CStr(k))
        If s = 0 Then
            none = none + 1
            If d("dates").Exists(k) Then
                If CDbl(d("maxDate")) - CDbl(d("dates")(k)) > 14 Then older = older + 1
            End If
        ElseIf rows - s >= 1 Then
            part = part + 1
        End If
    Next k

    ' Старение по корзинам.
    Dim labels() As String, agesN() As Double, agesP() As Double
    ReDim labels(1 To 5): ReDim agesN(1 To 5): ReDim agesP(1 To 5)
    labels(1) = "0–7 дней": labels(2) = "8–14": labels(3) = "15–30": labels(4) = "31–90": labels(5) = "> 90"
    Dim i As Long
    For i = 1 To 5: agesN(i) = 0: agesP(i) = 0: Next i
    For Each k In d("all").Keys
        Dim age As Double
        age = 0
        If d("dates").Exists(k) Then age = CDbl(d("maxDate")) - CDbl(d("dates")(k))
        Dim bi As Long
        If age <= 7 Then
            bi = 1
        ElseIf age <= 14 Then
            bi = 2
        ElseIf age <= 30 Then
            bi = 3
        ElseIf age <= 90 Then
            bi = 4
        Else
            bi = 5
        End If
        s = DictVal(d("signed"), CStr(k))
        rows = DictVal(d("rows"), CStr(k))
        If s = 0 Then agesN(bi) = agesN(bi) + 1
        If rows - s >= 1 And s >= 1 Then agesP(bi) = agesP(bi) + 1
    Next k
    Dim sA As String
    sA = "["
    For i = 1 To 5
        If i > 1 Then sA = sA & ","
        sA = sA & "{""bucket"":""" & JsonEscape(labels(i)) & """,""none"":" & JInt(agesN(i)) & ",""part"":" & JInt(agesP(i)) & "}"
    Next i
    sA = sA & "]"

    ' Вид ремонта неподписанных.
    Dim cnt As Object
    Set cnt = CreateObject("Scripting.Dictionary")
    For Each k In d("all").Keys
        If DictVal(d("signed"), CStr(k)) = 0 Then
            Dim zt As String
            zt = ""
            If d("zn").Exists(k) Then zt = CStr(d("zn")(k))
            If zt = "" Then zt = "(не указан)"
            If cnt.Exists(zt) Then cnt(zt) = CDbl(cnt(zt)) + 1 Else cnt(zt) = 1
        End If
    Next k
    Dim sZ As String, first As Boolean
    sZ = "["
    first = True
    For Each k In cnt.Keys
        If Not first Then sZ = sZ & ","
        first = False
        sZ = sZ & "{""zn_type"":""" & JsonEscape(CStr(k)) & """,""orders"":" & JInt(CDbl(cnt(k))) & "}"
    Next k
    sZ = sZ & "]"

    Dim kpi As String
    kpi = "{""none"":" & JInt(none) & ",""part"":" & JInt(part) & "," & _
        """events_unsigned"":" & JInt(CDbl(d("unsEvents"))) & ",""older_14d"":" & JInt(older) & "}"
    UnsignedToJson = "{""kpi"":" & kpi & ",""aging"":" & sA & ",""by_zn_type"":" & sZ & "}"
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
' Псевдонимы сотрудников (ТЗ v1.2, T3.3): ФИО <-> «Сотрудник N». Нумерация стабильная -
' сортировка ФИО по возрастанию (EmpList), N с 1. Кэш сбрасывается в ResetContentCaches
' вместе с остальными кэшами прогона.
' =====================================================================================
Private Sub BuildEmployeeAliases()
    If mAliasReady Then Exit Sub

    Set mAlias = CreateObject("Scripting.Dictionary")
    Dim emps As Variant
    emps = EmpList()
    Dim i As Long
    For i = LBound(emps) To UBound(emps)
        mAlias(CStr(emps(i))) = "Сотрудник " & CStr(i - LBound(emps) + 1)
    Next i
    mAliasReady = True
End Sub

Private Function AliasOf(ByVal employee As String) As String
    BuildEmployeeAliases
    If mAlias.Exists(employee) Then AliasOf = CStr(mAlias(employee)) Else AliasOf = employee
End Function

' Обратная замена «Сотрудник N» -> ФИО в тексте ИИ-выводов, по УБЫВАНИЮ N:
' иначе «Сотрудник 1» зацепит начало «Сотрудник 12» и оставит «ФИО2».
Private Function DeAlias(ByVal text As String) As String
    If Not mAliasReady Then DeAlias = text: Exit Function

    Dim emps As Variant
    emps = EmpList()
    Dim i As Long
    For i = UBound(emps) To LBound(emps) Step -1
        text = Replace(text, CStr(mAlias(CStr(emps(i)))), CStr(emps(i)))
    Next i
    DeAlias = text
End Function

' =====================================================================================
' Рендер-примитивы четырёх слайдов (ТЗ v1.2, T4). Все возвращают готовую HTML-строку.
' Всё, что пришло из 1С, проходит Esc (HtmlEscape) ДО склейки. Цвета SVG - ТОЛЬКО через
' CSS-переменные var(--...): хардкод hex ломает переключение тёмной/светлой темы.
' viewBox задан, width/height не фиксируются - масштабирование через CSS шаблона.
' role="img" + aria-label обязательны; пустой набор значений -> EmptyNote().
' =====================================================================================
' Число элементов массива (любые границы); не массив / нераспределённый -> 0.
Private Function ArrLen(a As Variant) As Long
    ArrLen = 0
    If Not IsArray(a) Then Exit Function
    On Error Resume Next
    ArrLen = UBound(a) - LBound(a) + 1
    If Err.Number <> 0 Then Err.Clear: ArrLen = 0
    On Error GoTo 0
End Function

Public Function EmptyNote() As String
    EmptyNote = "<div class='empty-note'>Нет данных</div>"
End Function

' Число для SVG-атрибута: десятичная ТОЧКА независимо от локали машины.
Private Function FmtN(v As Double) As String
    FmtN = Replace$(Format(v, "0.##"), ",", ".")
End Function

' 1158 -> «1 158» (неразрывный пробел U+00A0), независимо от локали.
Public Function FmtInt(v As Double) As String
    Dim s As String
    s = CStr(CLng(Abs(v)))
    Dim out As String, i As Long, grp As Long
    out = ""
    grp = 0
    For i = Len(s) To 1 Step -1
        out = Mid$(s, i, 1) & out
        grp = grp + 1
        If grp Mod 3 = 0 And i > 1 Then out = ChrW$(&HA0) & out
    Next i
    If v < 0 Then out = ChrW$(&H2212) & out   ' U+2212 MINUS SIGN, вне 1251
    FmtInt = out
End Function

' 0.723 -> «72,3» (десятичная запятая независимо от локали; без знака %).
Public Function FmtPct(v As Double) As String
    FmtPct = Replace$(Format(v * 100, "0.0"), ".", ",")
End Function

' Вертикальные столбики с подписью значения над столбцом (гистограмма времени в ремзоне).
Private Function SvgBarsV(labels As Variant, values As Variant, caption As String) As String
    If ArrLen(labels) = 0 Or ArrLen(values) = 0 Then SvgBarsV = EmptyNote(): Exit Function

    Dim lb As Long: lb = LBound(labels)
    Dim n As Long: n = ArrLen(labels)
    Dim mx As Double, i As Long
    mx = 0
    For i = lb To lb + n - 1
        If CDbl(values(i)) > mx Then mx = CDbl(values(i))
    Next i

    Dim w As Long, h As Long, baseY As Long, plotH As Long
    w = 70 + n * 64
    h = 290
    baseY = 240
    plotH = 190

    Dim svg As String
    svg = "<svg viewBox='0 0 " & w & " " & h & "' role='img' aria-label='" & Esc(caption) & "'>"
    For i = lb To lb + n - 1
        Dim bh As Double, bw As Long, x As Long, y As Double
        bw = 40
        x = 50 + (i - lb) * 64
        bh = 0
        If mx > 0 Then bh = CDbl(values(i)) / mx * plotH
        y = baseY - bh
        svg = svg & "<rect x='" & x & "' y='" & FmtN(y) & "' width='" & bw & "' height='" & FmtN(bh) & "' rx='3' fill='var(--s1)'/>"
        If CDbl(values(i)) > 0 Then
            svg = svg & "<text x='" & (x + bw \ 2) & "' y='" & FmtN(y - 6) & "' text-anchor='middle' font-size='11' fill='var(--ink)'>" & FmtInt(CDbl(values(i))) & "</text>"
        End If
        svg = svg & "<text x='" & (x + bw \ 2) & "' y='" & (baseY + 18) & "' text-anchor='middle' font-size='10' fill='var(--muted)'>" & Esc(CStr(labels(i))) & "</text>"
    Next i
    svg = svg & "<line x1='40' y1='" & baseY & "' x2='" & (w - 24) & "' y2='" & baseY & "' stroke='var(--line-strong)' stroke-width='1'/>"
    SvgBarsV = svg & "</svg>"
End Function

' Горизонтальные бары, две серии: totals - основная (var(--s1)), parts - «из них» (var(--s2)).
' highlightIdx - индекс строки labels (LBound-based), которую красить критическим цветом;
' -1 = нет подсветки.
Private Function SvgBarsH(labels As Variant, totals As Variant, parts As Variant, highlightIdx As Long) As String
    If ArrLen(labels) = 0 Then SvgBarsH = EmptyNote(): Exit Function

    Dim lb As Long: lb = LBound(labels)
    Dim n As Long: n = ArrLen(labels)
    Dim mx As Double, i As Long
    mx = 0
    For i = lb To lb + n - 1
        If CDbl(totals(i)) > mx Then mx = CDbl(totals(i))
    Next i
    If mx <= 0 Then SvgBarsH = EmptyNote(): Exit Function

    Dim rowH As Long: rowH = 32
    Dim w As Long: w = 700
    Dim h As Long: h = n * rowH + 20
    Dim labelW As Long: labelW = 180
    Dim barMaxW As Long: barMaxW = (w - labelW - 150) \ 2

    Dim svg As String
    svg = "<svg viewBox='0 0 " & w & " " & h & "' role='img' aria-label='горизонтальные бары'>"
    For i = lb To lb + n - 1
        Dim y0 As Long: y0 = 6 + (i - lb) * rowH
        Dim fillTxt As String, fillBar As String, fillPart As String
        fillTxt = "var(--ink)"
        fillBar = "var(--s1)"
        fillPart = "var(--s2)"
        If i = highlightIdx Then fillTxt = "var(--crit)": fillBar = "var(--crit)"
        Dim bwT As Double, bwP As Double
        bwT = CDbl(totals(i)) / mx * barMaxW
        bwP = CDbl(parts(i)) / mx * barMaxW
        svg = svg & "<text x='" & (labelW - 10) & "' y='" & (y0 + 13) & "' text-anchor='end' font-size='11' fill='" & fillTxt & "'>" & Esc(CStr(labels(i))) & "</text>"
        svg = svg & "<rect x='" & labelW & "' y='" & y0 & "' width='" & FmtN(bwT) & "' height='12' fill='" & fillBar & "'/>"
        svg = svg & "<rect x='" & labelW & "' y='" & (y0 + 15) & "' width='" & FmtN(bwP) & "' height='12' fill='" & fillPart & "'/>"
        svg = svg & "<text x='" & FmtN(labelW + bwT + 6) & "' y='" & (y0 + 11) & "' font-size='10' fill='var(--ink-2)'>" & FmtInt(CDbl(totals(i))) & "</text>"
        svg = svg & "<text x='" & FmtN(labelW + bwP + 6) & "' y='" & (y0 + 26) & "' font-size='10' fill='var(--ink-2)'>" & FmtInt(CDbl(parts(i))) & "</text>"
    Next i
    SvgBarsH = svg & "</svg>"
End Function

' Столбики объёма (var(--s1)) + линия доли 0..1 на второй оси (var(--s2)) с точками.
' Для «без поста по неделям»: bars - нарядов за неделю, linePct - доля без поста.
Private Function SvgBarsLine(labels As Variant, bars As Variant, linePct As Variant) As String
    If ArrLen(labels) = 0 Then SvgBarsLine = EmptyNote(): Exit Function

    Dim lb As Long: lb = LBound(labels)
    Dim n As Long: n = ArrLen(labels)
    Dim mx As Double, i As Long
    mx = 0
    For i = lb To lb + n - 1
        If CDbl(bars(i)) > mx Then mx = CDbl(bars(i))
    Next i
    If mx <= 0 Then SvgBarsLine = EmptyNote(): Exit Function

    Dim w As Long, h As Long, baseY As Long, plotH As Long
    w = 70 + n * 56
    h = 300
    baseY = 250
    plotH = 200

    Dim svg As String
    svg = "<svg viewBox='0 0 " & w & " " & h & "' role='img' aria-label='объём и доля без поста по неделям'>"
    svg = svg & "<text x='" & (w - 8) & "' y='" & (baseY - plotH + 4) & "' text-anchor='end' font-size='9' fill='var(--muted)'>100%</text>"
    svg = svg & "<text x='" & (w - 8) & "' y='" & (baseY + 4) & "' text-anchor='end' font-size='9' fill='var(--muted)'>0%</text>"

    Dim poly As String
    poly = ""
    For i = lb To lb + n - 1
        Dim x As Long, bw As Long, bh As Double, y As Double
        bw = 28
        x = 34 + (i - lb) * 56
        bh = CDbl(bars(i)) / mx * plotH
        y = baseY - bh
        svg = svg & "<rect x='" & x & "' y='" & FmtN(y) & "' width='" & bw & "' height='" & FmtN(bh) & "' fill='var(--s1)' opacity='0.85'/>"
        svg = svg & "<text x='" & (x + bw \ 2) & "' y='" & (baseY + 18) & "' text-anchor='middle' font-size='10' fill='var(--muted)'>" & Esc(CStr(labels(i))) & "</text>"
        Dim py As Double
        py = baseY - CDbl(linePct(i)) * plotH
        If poly = "" Then poly = FmtN(x + bw \ 2) & "," & FmtN(py) Else poly = poly & " " & FmtN(x + bw \ 2) & "," & FmtN(py)
        svg = svg & "<circle cx='" & (x + bw \ 2) & "' cy='" & FmtN(py) & "' r='3' fill='var(--s2)'/>"
    Next i
    svg = svg & "<polyline points='" & poly & "' fill='none' stroke='var(--s2)' stroke-width='2'/>"
    svg = svg & "<line x1='28' y1='" & baseY & "' x2='" & (w - 20) & "' y2='" & baseY & "' stroke='var(--line-strong)' stroke-width='1'/>"
    SvgBarsLine = svg & "</svg>"
End Function

' Спарклайн для плитки KPI (значения масштабируются по min..max). Пустой набор -> "".
Private Function SvgSpark(values As Variant, colorVar As String) As String
    If ArrLen(values) = 0 Then SvgSpark = "": Exit Function

    Dim lb As Long: lb = LBound(values)
    Dim n As Long: n = ArrLen(values)
    Dim mn As Double, mx As Double, i As Long, v As Double
    mn = CDbl(values(lb)): mx = CDbl(values(lb))
    For i = lb + 1 To lb + n - 1
        v = CDbl(values(i))
        If v < mn Then mn = v
        If v > mx Then mx = v
    Next i

    Dim W As Long: W = 100
    Dim H As Long: H = 28
    Dim poly As String
    poly = ""
    Dim range As Double
    range = mx - mn
    For i = lb To lb + n - 1
        Dim px As Double, py As Double
        If n > 1 Then px = 2 + (i - lb) * (W - 4) / (n - 1) Else px = W \ 2
        If range <= 0 Then py = H \ 2 Else py = H - 3 - (CDbl(values(i)) - mn) / range * (H - 6)
        If poly = "" Then poly = FmtN(px) & "," & FmtN(py) Else poly = poly & " " & FmtN(px) & "," & FmtN(py)
    Next i
    SvgSpark = "<svg class='spark' viewBox='0 0 " & W & " " & H & "' role='img' aria-label='тренд'>" & _
        "<polyline points='" & poly & "' fill='none' stroke='var(--" & colorVar & ")' stroke-width='2'/></svg>"
End Function

' Одна плитка KPI: подпись, значение, единица (small), дельта, спарклайн.
' deltaKind: "up" | "dn" | "flat" -> класс .delta.up/.dn/.flat.
Private Function KpiTile(label As String, value As String, unit As String, _
                         deltaText As String, deltaKind As String, spark As String) As String
    Dim html As String
    html = "<div class='kpi'><div class='lab'>" & Esc(label) & "</div>"
    html = html & "<div class='val num'>" & value
    If unit <> "" Then html = html & " <small>" & Esc(unit) & "</small>"
    html = html & "</div><div class='row'>"
    If deltaText <> "" Then
        html = html & "<span class='delta " & deltaKind & "'>" & Esc(deltaText) & "</span>"
    End If
    If spark <> "" Then html = html & spark
    html = html & "</div></div>"
    KpiTile = html
End Function

' =====================================================================================
' Блоки слайда 1 (ТЗ v1.2, T5). Единица счёта - заказ-наряд.
' =====================================================================================
' Подпись «как считается» под блоком (T8.2: .calc-note обязательна под каждым блоком).
Public Function CalcNote(text As String) As String
    CalcNote = "<p class='calc-note'>" & Esc(text) & "</p>"
End Function

' Число с десятичной запятой для HTML, независимо от локали.
Private Function FmtD(v As Double, fmt As String) As String
    FmtD = Replace$(Format(v, fmt), ".", ",")
End Function

' Среднее по столбцу снимка без фильтра (для чипов времени в ремзоне).
Private Function AvgSimple(valueCol As String, ByRef hasValue As Boolean) As Double
    AvgSimple = 0
    hasValue = False
    EnsureSnapshot
    Dim n As Long
    n = modAggregate.RowCount()
    Dim sum As Double, cnt As Long, r As Long
    sum = 0: cnt = 0
    For r = 1 To n
        Dim v As Variant
        v = modAggregate.CellRaw(r, valueCol)
        If IsNumeric(v) Then
            sum = sum + CDbl(v)
            cnt = cnt + 1
        End If
    Next r
    If cnt = 0 Then Exit Function
    hasValue = True
    AvgSimple = sum / cnt
End Function

' Текст дельты плитки: «(треугольник вверх) 5 к пр. нед.» / «(вниз) 1,2 п.п.» / «±0».
' Пусто, если предыдущей недели нет.
Private Function DeltaText(cur As Double, prev As Double, unit As String, hasPrev As Boolean) As String
    DeltaText = ""
    If Not hasPrev Then Exit Function
    Dim d As Double
    d = cur - prev
    If d = 0 Then
        DeltaText = IIf(unit = "pct", "±0,0 п.п.", "±0")
        Exit Function
    End If
    Dim arrow As String
    ' ChrW: символы вне ANSI-1251 нельзя держать в исходнике - при импорте в VBE
    ' (UTF-8 -> 1251) они превращаются в "?". См. tools/vba_lint_v1.0.
    arrow = IIf(d > 0, ChrW$(&H25B2), ChrW$(&H25BC))   ' U+25B2 / U+25BC
    If unit = "pct" Then
        DeltaText = arrow & " " & FmtPct(Abs(d)) & " п.п."
    ElseIf unit = "hhmm" Then
        DeltaText = arrow & " " & FormatHHMM(Abs(d))
    Else
        DeltaText = arrow & " " & FmtInt(Abs(d)) & " к пр. нед."
    End If
End Function

' Класс дельты: рост «хорошей» метрики - up, «плохой» - dn; flat при равенстве/отсутствии.
Private Function DeltaKind(cur As Double, prev As Double, betterWhenUp As Boolean, hasPrev As Boolean) As String
    DeltaKind = "flat"
    If Not hasPrev Then Exit Function
    If cur = prev Then Exit Function
    If cur > prev Then
        DeltaKind = IIf(betterWhenUp, "up", "dn")
    Else
        DeltaKind = IIf(betterWhenUp, "dn", "up")
    End If
End Function

' {{FACTS}} - четыре факта выгрузки в шапке (постановка §3.1).
Private Function BuildFacts() As String
    EnsureSnapshot
    Dim n As Long
    n = modAggregate.RowCount()

    Dim nums As Object
    Set nums = modAggregate.DistinctValues("number", FBase())
    Dim noPostNums As Object
    Set noPostNums = CreateObject("Scripting.Dictionary")

    Dim dMin As Double, dMax As Double
    dMin = 0: dMax = 0
    Dim r As Long
    For r = 1 To n
        Dim dte As Variant
        dte = modAggregate.CellRaw(r, "date")
        If IsNumeric(dte) Then
            If dMin = 0 Or CDbl(dte) < dMin Then dMin = CDbl(dte)
            If CDbl(dte) > dMax Then dMax = CDbl(dte)
        End If
        If modAggregate.CellText(r, "postN") = "" Then
            Dim num As String
            num = modAggregate.CellText(r, "number")
            If num <> "" Then noPostNums(num) = True
        End If
    Next r

    Dim months As String
    months = "—"
    If dMin > 0 And dMax >= dMin Then
        months = "~" & CStr(DateDiff("m", CDate(dMin), CDate(dMax))) & " мес"
    End If
    Dim noPostPct As String
    If nums.Count > 0 Then
        noPostPct = FmtPct(CDbl(noPostNums.Count) / CDbl(nums.Count)) & " %"
    Else
        noPostPct = "—"
    End If

    Dim html As String
    html = "<div><dt>Событий</dt><dd class='num'>" & FmtInt(CDbl(n)) & "</dd></div>"
    html = html & "<div><dt>Нарядов</dt><dd class='num'>" & FmtInt(CDbl(nums.Count)) & "</dd></div>"
    html = html & "<div><dt>История</dt><dd class='num'>" & Esc(months) & "</dd></div>"
    html = html & "<div><dt>Без поста</dt><dd class='num'>" & Esc(noPostPct) & "</dd></div>"
    BuildFacts = html
End Function

' {{KPI_OVERVIEW}} - семь плиток за отчётную неделю с дельтой к ПН-1 и спарклайном за окно.
Private Function BuildKpiOverview() As String
    EnsureSnapshot
    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))

    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(windowSize, cnt)
    If cnt = 0 Then BuildKpiOverview = EmptyNote(): Exit Function

    ' Карта yearWeek -> индекс окна (1 = ПН, самая свежая из окна).
    Dim wmap As Object
    Set wmap = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To cnt
        wmap(CStr(KeyPart(weeks(i), 0))) = i
    Next i

    ' Множества номеров нарядов и счётчики по неделям.
    Dim opened() As Object, closed() As Object, noPost() As Object
    ReDim opened(1 To cnt): ReDim closed(1 To cnt): ReDim noPost(1 To cnt)
    Dim allEv() As Double, unsEv() As Double, pctTot() As Double, pctTab() As Double
    ReDim allEv(1 To cnt): ReDim unsEv(1 To cnt): ReDim pctTot(1 To cnt): ReDim pctTab(1 To cnt)
    For i = 1 To cnt
        Set opened(i) = CreateObject("Scripting.Dictionary")
        Set closed(i) = CreateObject("Scripting.Dictionary")
        Set noPost(i) = CreateObject("Scripting.Dictionary")
        allEv(i) = 0: unsEv(i) = 0: pctTot(i) = 0: pctTab(i) = 0
    Next i

    Dim n As Long
    n = modAggregate.RowCount()
    Dim r As Long
    For r = 1 To n
        Dim wd As Long, ws As Long
        wd = WeekKeyOfSafe(modAggregate.CellRaw(r, "date"))
        ws = 0
        Dim ywT As String
        ywT = modAggregate.CellText(r, "yearWeek")
        If ywT <> "" And IsNumeric(ywT) Then ws = CLng(ywT)

        Dim armT As String
        armT = modAggregate.CellText(r, "arm")
        Dim num As String
        num = modAggregate.CellText(r, "number")
        Dim idxD As Long, idxS As Long
        idxD = 0: idxS = 0
        If wd > 0 Then
            If wmap.Exists(CStr(wd)) Then idxD = CLng(wmap(CStr(wd)))
        End If
        If ws > 0 Then
            If wmap.Exists(CStr(ws)) Then idxS = CLng(wmap(CStr(ws)))
        End If

        ' По дате создания: открытые, без поста, события недели (знаменатель «не подписано»).
        If idxD > 0 Then
            allEv(idxD) = allEv(idxD) + 1
            If num <> "" Then
                opened(idxD)(num) = True
                If modAggregate.CellText(r, "postN") = "" Then noPost(idxD)(num) = True
            End If
            If armT = "НЕ ПОДПИСАНО" Then unsEv(idxD) = unsEv(idxD) + 1
        End If
        ' По дате статуса: % планшета, закрытые.
        If idxS > 0 Then
            If armT = "ПК" Or armT = "ПЛАНШЕТ" Then
                pctTot(idxS) = pctTot(idxS) + 1
                If armT = "ПЛАНШЕТ" Then pctTab(idxS) = pctTab(idxS) + 1
            End If
            If IsStatusReadyToLeave(modAggregate.CellText(r, "ready_for")) Then
                If num <> "" Then closed(idxS)(num) = True
            End If
        End If
    Next r

    ' Медианы deltaHours по неделям (строки «Готов к выбытию»).
    Dim medD As Object
    Set medD = modAggregate.GroupPercentile(Array("yearWeek"), "deltaHours", 0.5, _
        Array("ready_for=Готов к выбытию"))
    Dim med() As Double
    ReDim med(1 To cnt)
    For i = 1 To cnt
        med(i) = DictVal(medD, CStr(KeyPart(weeks(i), 0)) & "|")
    Next i

    ' «Висит на конец недели» - накопительно по возрастанию недель окна.
    Dim hang() As Double
    ReDim hang(1 To cnt)
    Dim cumOp As Double, cumCl As Double, j As Long
    cumOp = 0: cumCl = 0
    For j = cnt To 1 Step -1
        cumOp = cumOp + CDbl(opened(j).Count)
        cumCl = cumCl + CDbl(closed(j).Count)
        hang(j) = cumOp - cumCl
    Next j

    ' Спарклайны (по возрастанию недель окна).
    Dim spOpen() As Double, spClose() As Double, spHang() As Double, spMed() As Double
    Dim spNoPost() As Double, spPct() As Double, spUns() As Double
    ReDim spOpen(1 To cnt): ReDim spClose(1 To cnt): ReDim spHang(1 To cnt): ReDim spMed(1 To cnt)
    ReDim spNoPost(1 To cnt): ReDim spPct(1 To cnt): ReDim spUns(1 To cnt)
    For j = 1 To cnt
        Dim src As Long
        src = cnt - j + 1
        spOpen(j) = CDbl(opened(src).Count)
        spClose(j) = CDbl(closed(src).Count)
        spHang(j) = hang(src)
        spMed(j) = med(src)
        spNoPost(j) = SafePercentNo(opened(src).Count, noPost(src).Count)
        spPct(j) = SafePctTwo(pctTot(src), pctTab(src))
        spUns(j) = SafePctTwo(allEv(src), unsEv(src))
    Next j

    Dim hasPrev As Boolean
    hasPrev = (cnt >= 2)

    Dim pctCur As Double, pctPrev As Double
    pctCur = SafePctTwo(pctTot(1), pctTab(1))
    pctPrev = 0
    If hasPrev Then pctPrev = SafePctTwo(pctTot(2), pctTab(2))
    Dim npCur As Double, npPrev As Double
    npCur = SafePercentNo(opened(1).Count, noPost(1).Count)
    npPrev = 0
    If hasPrev Then npPrev = SafePercentNo(opened(2).Count, noPost(2).Count)
    Dim unsCur As Double, unsPrev As Double
    unsCur = SafePctTwo(allEv(1), unsEv(1))
    unsPrev = 0
    If hasPrev Then unsPrev = SafePctTwo(allEv(2), unsEv(2))

    ' Защита от окна из одной недели: обращение к (2) допустимо только при hasPrev,
    ' иначе «Subscript out of range» (массивы открытых/закрытых имеют размер cnt).
    Dim opCur As Double, opPrev As Double, clCur As Double, clPrev As Double
    opCur = CDbl(opened(1).Count): opPrev = 0
    clCur = CDbl(closed(1).Count): clPrev = 0
    If hasPrev Then opPrev = CDbl(opened(2).Count): clPrev = CDbl(closed(2).Count)
    Dim hangPrev As Double
    hangPrev = 0
    If hasPrev Then hangPrev = hang(2)

    Dim medVal As String, medDelta As String, medKind As String
    Dim medPrev As Double
    medPrev = 0
    If hasPrev Then medPrev = med(2)
    If med(1) > 0 Then
        medVal = FormatHHMM(med(1))
        medDelta = DeltaText(med(1), medPrev, "hhmm", hasPrev)
        medKind = DeltaKind(med(1), medPrev, False, hasPrev)
    Else
        medVal = "н/д"
        medDelta = ""
        medKind = "flat"
    End If

    Dim html As String
    html = KpiTile("Нарядов открыто", FmtInt(opCur), "", _
        DeltaText(opCur, opPrev, "num", hasPrev), _
        DeltaKind(opCur, opPrev, True, hasPrev), SvgSpark(spOpen, "s1"))
    html = html & KpiTile("Закрыто нарядов", FmtInt(clCur), "", _
        DeltaText(clCur, clPrev, "num", hasPrev), _
        DeltaKind(clCur, clPrev, True, hasPrev), SvgSpark(spClose, "s3"))
    html = html & KpiTile("Висит на конец недели", FmtInt(hang(1)), "", _
        DeltaText(hang(1), hangPrev, "num", hasPrev), _
        DeltaKind(hang(1), hangPrev, False, hasPrev), SvgSpark(spHang, "s2"))
    html = html & KpiTile("Медиана в ремзоне", medVal, "", medDelta, medKind, SvgSpark(spMed, "s1"))
    html = html & KpiTile("Без поста ремзоны", FmtInt(CDbl(noPost(1).Count)), "· " & FmtPct(npCur) & " %", _
        DeltaText(npCur, npPrev, "pct", hasPrev), DeltaKind(npCur, npPrev, False, hasPrev), SvgSpark(spNoPost, "s4"))
    html = html & KpiTile("% планшета", FmtPct(pctCur), "%", _
        DeltaText(pctCur, pctPrev, "pct", hasPrev), DeltaKind(pctCur, pctPrev, True, hasPrev), SvgSpark(spPct, "s3"))
    html = html & KpiTile("Не подписано", FmtInt(unsEv(1)), "· " & FmtPct(unsCur) & " %", _
        DeltaText(unsCur, unsPrev, "pct", hasPrev), DeltaKind(unsCur, unsPrev, False, hasPrev), SvgSpark(spUns, "s2"))

    html = html & "<p class='calc-note'>" & ChrW$(&H26A0) & " Медиана меряет всё время от приёмки до подписания выбытия, включая очередь и ожидание запчастей; для «плана против факта» непригодна.</p>"
    html = html & CalcNote("За отчётную неделю (REPORT/WEEK). Открыто/без поста/не подписано - по дате создания (date), " & _
        "закрыто/% планшета/медиана - по дате статуса (status_date). Наряды - уникальные number; " & _
        "дельта - к предыдущей неделе окна, спарклайн - за WEEKS_WINDOW недель.")
    BuildKpiOverview = html
End Function

' Доля без деления на ноль; parts/count - счётчики (Double).
Private Function SafePercentNo(count As Double, parts As Double) As Double
    If count > 0 Then SafePercentNo = parts / count Else SafePercentNo = 0
End Function

Private Function SafePctTwo(total As Double, part As Double) As Double
    If total > 0 Then SafePctTwo = part / total Else SafePctTwo = 0
End Function

' {{BLOCK_TIME_STATS}} - чипы медиана/среднее/p90 по deltaHours (без обрезки).
Private Function BuildTimeStats() As String
    EnsureSnapshot
    Dim has As Boolean
    Dim med As Double, p90 As Double
    med = modAggregate.Percentile("deltaHours", 0.5, Empty, has)
    If Not has Then BuildTimeStats = EmptyNote(): Exit Function
    p90 = modAggregate.Percentile("deltaHours", 0.9, Empty, has)

    Dim avg As Double, avgHas As Boolean
    avg = AvgSimple("deltaHours", avgHas)

    Dim html As String
    html = "<div class='chips' style='margin:10px 0 6px'>"
    html = html & "<span class='chip'>медиана " & FormatHHMM(med) & "</span>"
    If avgHas Then html = html & "<span class='chip'>среднее " & FmtD(avg, "0.0") & " ч</span>"
    html = html & "<span class='chip'>p90 " & FmtD(p90, "0.0") & " ч</span>"
    html = html & "</div>"
    html = html & "<p class='calc-note'>Среднее перекошено хвостом распределения — ориентироваться на медиану.</p>"
    BuildTimeStats = html
End Function

' {{BLOCK_TIME_HIST}} - корзины времени в ремзоне, обрезка p99 (ТЗ T5.3).
Private Function BuildTimeHistogram() As String
    EnsureSnapshot
    Const BIN1 As Double = 1
    Const BIN4 As Double = 4
    Const BIN8 As Double = 8
    Const BIN24 As Double = 24
    Const BIN72 As Double = 72

    Dim has As Boolean
    Dim p99 As Double
    p99 = modAggregate.Percentile("deltaHours", 0.99, Empty, has)
    If Not has Then BuildTimeHistogram = EmptyNote(): Exit Function

    Dim labels() As Variant, values() As Double
    ReDim labels(1 To 6): ReDim values(1 To 6)
    labels(1) = "< 1 ч": labels(2) = "1–4 ч": labels(3) = "4–8 ч"
    labels(4) = "8–24 ч": labels(5) = "1–3 сут": labels(6) = "> 3 сут"
    Dim i As Long
    For i = 1 To 6: values(i) = 0: Next i

    Dim n As Long, r As Long, dropped As Long
    n = modAggregate.RowCount()
    dropped = 0
    For r = 1 To n
        Dim v As Variant
        v = modAggregate.CellRaw(r, "deltaHours")
        If IsNumeric(v) Then
            If CDbl(v) > p99 Then
                dropped = dropped + 1
            ElseIf CDbl(v) < BIN1 Then
                values(1) = values(1) + 1
            ElseIf CDbl(v) < BIN4 Then
                values(2) = values(2) + 1
            ElseIf CDbl(v) < BIN8 Then
                values(3) = values(3) + 1
            ElseIf CDbl(v) < BIN24 Then
                values(4) = values(4) + 1
            ElseIf CDbl(v) < BIN72 Then
                values(5) = values(5) + 1
            Else
                values(6) = values(6) + 1
            End If
        End If
    Next r

    Dim html As String
    html = SvgBarsV(labels, values, "Время в ремзоне по корзинам, часов")
    html = html & CalcNote("Пары (наряд " & ChrW$(&HD7) & " дирекция) с заполненным deltaHours; верхний 1 % значений обрезан " & _
        "(p99 = " & FmtD(p99, "0.0") & " ч) — отброшено " & FmtInt(CDbl(dropped)) & " строк.")
    BuildTimeHistogram = html
End Function

' {{BLOCK_FLOW_ZNTYPE}} - виды ремонта: нарядов и доля.
Private Function BuildFlowByZnType() As String
    EnsureSnapshot
    Dim counts As Object
    Set counts = modAggregate.GroupCountDistinct(Array("zn_type"), "number", FBase())
    If counts.Count = 0 Then BuildFlowByZnType = EmptyNote(): Exit Function

    Dim total As Double
    total = CDbl(modAggregate.DistinctValues("number", FBase()).Count)

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(counts, True)

    Dim html As String, i As Long
    html = "<table class='block-table'><thead><tr><th>Вид ремонта</th><th>Нарядов</th><th>%</th></tr></thead><tbody>"
    For i = LBound(sorted) To UBound(sorted)
        Dim k As String
        k = CStr(sorted(i))
        Dim title As String
        title = KeyPart(k, 0)
        If title = "" Then title = "(не указан)"
        html = html & "<tr><td>" & Esc(title) & "</td><td class='num'>" & FmtInt(CDbl(counts(k))) & "</td>"
        If total > 0 Then
            html = html & "<td class='num'>" & FmtPct(CDbl(counts(k)) / total) & "</td></tr>"
        Else
            html = html & "<td class='num'>—</td></tr>"
        End If
    Next i
    html = html & "</tbody></table>"
    html = html & CalcNote("Вид ремонта (zn_type) за весь период выгрузки; нарядов — уникальные number; % — от всех нарядов периода. " & _
        "Поток на три четверти — внеплановый ремонт; плановое ТО в ремзоне почти не видно — стоит проверить у заказчика.")
    BuildFlowByZnType = html
End Function

' {{BLOCK_FLOW_DEFEKT}} - бары по разделам дефекта; пустое - отдельной строкой.
Private Function BuildFlowByDefekt() As String
    EnsureSnapshot
    Dim counts As Object
    Set counts = modAggregate.GroupCountDistinct(Array("defekt_type"), "number", FBase())
    If counts.Count = 0 Then BuildFlowByDefekt = EmptyNote(): Exit Function

    Dim emptyCnt As Double
    emptyCnt = DictVal(counts, "|")
    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(counts, True)
    Dim nRows As Long, i As Long
    nRows = UBound(sorted) - LBound(sorted) + 1
    If emptyCnt > 0 Then nRows = nRows - 1
    If nRows <= 0 Then
        BuildFlowByDefekt = "<p class='calc-note'>Раздел не указан: " & FmtInt(emptyCnt) & " нарядов</p>" & _
            CalcNote("defekt_type — разделы классификатора, а не диагнозы (12 значений на весь массив); метрику повторного ремонта на них строить нельзя.")
        Exit Function
    End If

    Dim labels() As Variant, vals() As Double, parts() As Double
    ReDim labels(1 To nRows): ReDim vals(1 To nRows): ReDim parts(1 To nRows)
    Dim pos As Long
    pos = 0
    For i = LBound(sorted) To UBound(sorted)
        Dim k As String
        k = CStr(sorted(i))
        If KeyPart(k, 0) <> "" Then
            pos = pos + 1
            labels(pos) = KeyPart(k, 0)
            vals(pos) = CDbl(counts(k))
            parts(pos) = 0
        End If
    Next i

    Dim html As String
    html = SvgBarsH(labels, vals, parts, -1)
    If emptyCnt > 0 Then
        html = html & "<p class='calc-note'>Раздел не указан: " & FmtInt(emptyCnt) & " нарядов</p>"
    End If
    html = html & CalcNote("defekt_type — разделы классификатора, а не диагнозы (12 значений на весь массив); метрику повторного ремонта на них строить нельзя.")
    BuildFlowByDefekt = html
End Function

' {{BLOCK_NOPOST_WEEKLY}} - столбики объёма + линия доли без поста по неделям окна.
Private Function BuildNoPostWeekly() As String
    EnsureSnapshot
    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))

    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(windowSize, cnt)
    If cnt = 0 Then BuildNoPostWeekly = EmptyNote(): Exit Function

    ' Ось - yearWeek по status_date (постановка §3.5); строка с пустым постом - фильтр «postN=».
    Dim total As Object, noPost As Object
    Set total = modAggregate.GroupCountDistinct(Array("yearWeek"), "number", FSigned())
    Set noPost = modAggregate.GroupCountDistinct(Array("yearWeek"), "number", _
        AppendFilter(FSigned(), "postN="))
    If total.Count = 0 Then BuildNoPostWeekly = EmptyNote(): Exit Function

    Dim labels() As Variant, bars() As Double, pct() As Double
    ReDim labels(1 To cnt): ReDim bars(1 To cnt): ReDim pct(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        Dim wkStr As String
        wkStr = CStr(KeyPart(weeks(i), 0))
        labels(i) = WeekLabel(KeyPart(weeks(i), 0))
        bars(i) = DictVal(total, wkStr & "|")
        pct(i) = SafePercent(noPost, total, wkStr & "|")
    Next i

    Dim html As String
    html = SvgBarsLine(labels, bars, pct)
    html = html & "<p class='calc-note'>Доля без поста в последней неделе окна: " & FmtPct(pct(1)) & " %.</p>"
    html = html & CalcNote("Недели окна WEEKS_WINDOW по yearWeek (status_date), подписанные события; столбики — нарядов за неделю " & _
        "(уникальные number), линия — доля нарядов с пустым post на правой оси 0–100 %. «Без поста» и «не подписано» — по данным один дефект (мост к слайду 4).")
    BuildNoPostWeekly = html
End Function

' =====================================================================================
' Блоки слайдов 2-3 (ТЗ v1.2, T6). Все функции параметризуются дирекцией ("ДЭНТ" | "ДГМ"):
' двух копий кода быть не должно, слайд 3 отличается только значением параметра.
' =====================================================================================
' {{BLOCK_WEEKS_<DIR>}} / {{BLOCK_WEEKS_ZONES_<DIR>}} - понедельная таблица использования планшета.
Private Function BuildWeeksTable(dir As String, zonesFilter As String) As String
    EnsureSnapshot
    Dim f As Variant, ft As Variant
    f = FDir(dir)
    ft = AppendFilter(FTablet(), "direction=" & dir)
    If Trim$(zonesFilter) <> "" Then
        f = AppendFilter(f, "postN@=" & zonesFilter)
        ft = AppendFilter(ft, "postN@=" & zonesFilter)
    End If

    Dim total As Object, tablet As Object
    Set total = modAggregate.GroupCount(Array("yearWeek"), f)
    Set tablet = modAggregate.GroupCount(Array("yearWeek"), ft)
    If total.Count = 0 Then
        BuildWeeksTable = "<h3>" & Esc(dir) & " · по неделям</h3>" & EmptyNote()
        Exit Function
    End If

    ' Окно WEEKS_WINDOW недель <= отчётной, по убыванию (weeks(1) = ПН).
    Dim windowSize As Long
    windowSize = CLng(modMain.GetVariableDef("REPORT/WEEKS_WINDOW", "8"))
    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(windowSize, cnt)
    If cnt = 0 Then
        BuildWeeksTable = "<h3>" & Esc(dir) & " · по неделям</h3>" & EmptyNote()
        Exit Function
    End If

    Dim title As String
    title = dir
    If Trim$(zonesFilter) <> "" Then
        title = title & " · ремзоны: " & Esc(Replace(Trim$(zonesFilter), ";", ", "))
    Else
        title = title & " · все ремзоны"
    End If

    Dim html As String, wj As Long
    html = "<h3>" & Esc(title) & "</h3>"
    html = html & "<div class='scroll'><table class='block-table matrix'><thead><tr><th>" & Esc(dir) & "</th>"
    For wj = cnt To 1 Step -1
        html = html & "<th>" & Esc(WeekLabel(KeyPart(weeks(wj), 0))) & "</th>"
    Next wj
    html = html & "</tr></thead><tbody>"

    ' Строки: всего событий -> % -> ПЛАНШЕТ -> ПК.
    html = html & "<tr><th class='row-head'>" & Esc(dir) & "</th>"
    For wj = cnt To 1 Step -1
        Dim wkStr As String
        wkStr = CStr(KeyPart(weeks(wj), 0))
        html = html & "<td class='num'>" & FmtInt(DictVal(total, wkStr & "|")) & "</td>"
    Next wj
    html = html & "</tr>"

    html = html & "<tr class='pct-row'><th class='row-head'>%</th>"
    For wj = cnt To 1 Step -1
        wkStr = CStr(KeyPart(weeks(wj), 0))
        If DictVal(total, wkStr & "|") = 0 Then
            html = html & PctCell(0, False)
        Else
            html = html & PctCell(SafePercent(tablet, total, wkStr & "|"))
        End If
    Next wj
    html = html & "</tr>"

    html = html & "<tr><th class='row-head'>ПЛАНШЕТ</th>"
    For wj = cnt To 1 Step -1
        wkStr = CStr(KeyPart(weeks(wj), 0))
        html = html & "<td class='num'>" & FmtInt(DictVal(tablet, wkStr & "|")) & "</td>"
    Next wj
    html = html & "</tr>"

    html = html & "<tr><th class='row-head'>ПК</th>"
    For wj = cnt To 1 Step -1
        wkStr = CStr(KeyPart(weeks(wj), 0))
        html = html & "<td class='num'>" & FmtInt(DictVal(total, wkStr & "|") - DictVal(tablet, wkStr & "|")) & "</td>"
    Next wj
    html = html & "</tr>"

    html = html & "</tbody></table></div>"
    html = html & CalcNote("Окно WEEKS_WINDOW недель <= REPORT/WEEK, события с arm из {ПК, ПЛАНШЕТ}; % = ПЛАНШЕТ / (ПК + ПЛАНШЕТ) за неделю; пустые ячейки — нет событий.")
    BuildWeeksTable = html
End Function

' {{BLOCK_POSTS_<DIR>}} - бары по постам за отчётную неделю + оценки (pill).
Private Function BuildPostsChart(dir As String) As String
    EnsureSnapshot
    Dim rw As Long
    rw = ReportWeekValue()
    Dim f As Variant, ft As Variant
    f = AppendFilter(FDir(dir), "yearWeek=" & CStr(rw))
    ft = AppendFilter(FTablet(), "direction=" & dir)
    ft = AppendFilter(ft, "yearWeek=" & CStr(rw))

    ' Единица счёта - наряд: Count(Distinct number).
    Dim records As Object, tabRec As Object
    Set records = modAggregate.GroupCountDistinct(Array("postN"), "number", f)
    Set tabRec = modAggregate.GroupCountDistinct(Array("postN"), "number", ft)
    If records.Count = 0 Then
        BuildPostsChart = "<h3>" & Esc(dir) & " · по постам</h3>" & EmptyNote()
        Exit Function
    End If

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(records, True)
    Dim nRows As Long, i As Long
    nRows = UBound(sorted) - LBound(sorted) + 1

    Dim labels() As Variant, vals() As Double, parts() As Double
    ReDim labels(1 To nRows): ReDim vals(1 To nRows): ReDim parts(1 To nRows)
    Dim hlIdx As Long
    hlIdx = -1
    For i = 1 To nRows
        Dim k As String
        k = CStr(sorted(i - 1))
        Dim pn As String
        pn = KeyPart(k, 0)
        If pn = "" Then
            labels(i) = "Пост не указан"
            hlIdx = i
        Else
            labels(i) = pn
        End If
        vals(i) = DictVal(records, pn & "|")
        parts(i) = DictVal(tabRec, pn & "|")
    Next i

    Dim html As String
    html = "<h3>" & Esc(dir) & " · где подписывают с ПК (нед. " & Esc(WeekLabel(rw)) & ")</h3>"
    html = html & "<p class='calc-note'>Единица счёта — наряд (уникальные number), соседние блоки считают события.</p>"
    html = html & SvgBarsH(labels, vals, parts, hlIdx)

    ' Оценки постов (pill): норма/провал/мало данных.
    Dim minPost As Long, norma As Double, proval As Double
    minPost = CLng(modMain.GetVariableDef("REPORT/MIN_POST_RECORDS", "10"))
    norma = CDbl(Val(Replace(modMain.GetVariableDef("NormaForPlanshet", "90"), ",", ".")))
    proval = CDbl(Val(Replace(modMain.GetVariableDef("ProvalForPlanshet", "50"), ",", ".")))
    html = html & "<table class='block-table'><thead><tr><th>Пост</th><th>Нарядов</th><th>% планшет</th><th>Оценка</th></tr></thead><tbody>"
    For i = 1 To nRows
        Dim pct As Double
        pct = SafePercentNo(vals(i), parts(i))
        Dim grade As String
        If vals(i) < minPost Then
            grade = "<span class='pill warn'>мало данных</span>"
        ElseIf proval < norma And pct * 100 >= norma Then
            grade = "<span class='pill good'>норма</span>"
        ElseIf proval < norma And pct * 100 < proval Then
            grade = "<span class='pill crit'>провал</span>"
        Else
            grade = "—"
        End If
        html = html & "<tr><td>" & Esc(labels(i)) & "</td><td class='num'>" & FmtInt(vals(i)) & "</td>" & _
            PctCell(pct) & "<td>" & grade & "</td></tr>"
    Next i
    html = html & "</tbody></table>"
    html = html & CalcNote("За отчётную неделю (REPORT/WEEK, yearWeek статуса); пост — родитель поста, строки читаются как площадки; " & _
        "оценка: норма при % >= NormaForPlanshet, провал при % < ProvalForPlanshet, «мало данных» при нарядов < REPORT/MIN_POST_RECORDS.")
    BuildPostsChart = html
End Function

' {{BLOCK_PEOPLE_<DIR>}} / {{BLOCK_DEPS_<DIR>}} - понедельная раскладка ПН-3..ПН по людям.
' groupField = "employee" либо "emp_dep" - одна функция на обе таблицы.
Private Function BuildPeopleWeekly(dir As String, groupField As String) As String
    EnsureSnapshot
    Dim weeks As Variant, cnt As Long
    weeks = RecentWeeksUpTo(4, cnt)
    If cnt = 0 Then
        BuildPeopleWeekly = "<h4>По " & Esc(groupField) & "</h4>" & EmptyNote()
        Exit Function
    End If

    Dim f As Variant
    f = FDir(dir)

    ' Всего подписей = Count(Distinct Key) за неделю.
    Dim tot As Object, tabObj As Object, tabAcc As Object, tabLeave As Object, delAvg As Object
    Set tot = modAggregate.GroupCountDistinct(Array(groupField, "yearWeek"), "Key", f)
    Set tabObj = modAggregate.GroupCountDistinct(Array(groupField, "yearWeek"), "Key", _
        AppendFilter(f, "arm=ПЛАНШЕТ"))
    Set tabAcc = modAggregate.GroupCountDistinct(Array(groupField, "yearWeek"), "Key", _
        AppendFilter(f, "arm=ПЛАНШЕТ;ready_for=Готов к приемке"))
    Set tabLeave = modAggregate.GroupCountDistinct(Array(groupField, "yearWeek"), "Key", _
        AppendFilter(f, "arm=ПЛАНШЕТ;ready_for=Готов к выбытию"))
    Set delAvg = modAggregate.GroupAverage(Array(groupField, "yearWeek"), "deltaHours", f)

    If tot.Count = 0 Then
        BuildPeopleWeekly = "<h4>По " & Esc(groupField) & "</h4>" & EmptyNote()
        Exit Function
    End If

    ' Строки и их сортировка по «% планшет» за ПН (weeks(1)), по убыванию.
    Dim rowSet As Object, pctDict As Object
    Set rowSet = CreateObject("Scripting.Dictionary")
    Set pctDict = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In tot.Keys
        rowSet(KeyPart(k, 0)) = True
    Next k
    Dim wPN As String
    wPN = CStr(KeyPart(weeks(1), 0))
    Dim rv As Variant
    For Each rv In rowSet.Keys
        Dim keyTot As String
        keyTot = CStr(rv) & "|" & wPN & "|"
        If DictVal(tot, keyTot) = 0 Then
            pctDict(rv) = -1
        Else
            pctDict(rv) = SafePercent(tabObj, tot, keyTot)
        End If
    Next rv
    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(pctDict, True)

    Dim caption As String
    If groupField = "employee" Then caption = "Сотрудник" Else caption = "Подразделение"

    Dim html As String, i As Long, wj As Long
    html = "<h4>" & Esc(caption) & "</h4>"
    html = html & "<div class='scroll'><table class='block-table people'><thead><tr><th rowspan='2'>" & Esc(caption) & "</th>"
    For wj = cnt To 1 Step -1
        Dim grp As String
        If wj = 1 Then
            grp = "ПН (" & Esc(WeekLabel(KeyPart(weeks(wj), 0))) & ")"
        Else
            grp = "ПН-" & CStr(wj - 1) & " (" & Esc(WeekLabel(KeyPart(weeks(wj), 0))) & ")"
        End If
        html = html & "<th colspan='6'>" & grp & "</th>"
    Next wj
    html = html & "</tr><tr>"
    For wj = cnt To 1 Step -1
        html = html & "<th>% планшет</th><th>Всего подписей</th><th>Из них на планшете</th>" & _
            "<th>из них Готов к приемке</th><th>из них Готов к выбытию</th><th>Среднее время</th>"
    Next wj
    html = html & "</tr></thead><tbody>"

    For i = LBound(sorted) To UBound(sorted)
        Dim nameVal As String
        nameVal = CStr(sorted(i))
        html = html & "<tr><th class='row-head'>" & Esc(nameVal) & "</th>"
        For wj = cnt To 1 Step -1
            Dim wkStr As String
            wkStr = CStr(KeyPart(weeks(wj), 0))
            Dim key2 As String
            key2 = nameVal & "|" & wkStr & "|"
            If DictVal(tot, key2) = 0 Then
                html = html & PctCell(0, False)
            Else
                html = html & PctCell(SafePercent(tabObj, tot, key2))
            End If
            html = html & "<td class='num'>" & FmtInt(DictVal(tot, key2)) & "</td>"
            html = html & "<td class='num'>" & FmtInt(DictVal(tabObj, key2)) & "</td>"
            html = html & "<td class='num'>" & FmtInt(DictVal(tabAcc, key2)) & "</td>"
            html = html & "<td class='num'>" & FmtInt(DictVal(tabLeave, key2)) & "</td>"
            If delAvg.Exists(key2) Then
                html = html & "<td class='num'>" & FormatHHMM(CDbl(delAvg(key2))) & "</td>"
            Else
                html = html & "<td class='num'>—</td>"
            End If
        Next wj
        html = html & "</tr>"
    Next i
    html = html & "</tbody></table></div>"

    If groupField = "emp_dep" Then
        html = html & CalcNote("Те же метрики по подразделениям (emp_dep); % планшет = ПЛАНШЕТ / (ПК + ПЛАНШЕТ) событий за неделю, «Среднее время» — средняя deltaHours по неделе, «чч:мм».")
    Else
        html = html & CalcNote("Недели ПН-3…ПН, присутствующие в данных; % планшет = ПЛАНШЕТ / (ПК + ПЛАНШЕТ) событий сотрудника за неделю; " & _
            "всего подписей — уникальные Key; «Готов к приемке»/«Готов к выбытию» сравниваются с нормализацией «ё»" & ChrW$(&H2192) & "«е»; сортировка по % планшет за ПН, по убыванию.")
    End If
    BuildPeopleWeekly = html
End Function

' Кэш статистики подписания по дирекции: all/acc/lea/dates/maxDate (один проход на дирекцию).
Private Sub EnsureSignData(dir As String)
    If Not mSignReady Is Nothing Then
        If mSignReady.Exists(dir) Then Exit Sub
    End If
    EnsureSnapshot
    If mSignReady Is Nothing Then Set mSignReady = CreateObject("Scripting.Dictionary")
    If mSignData Is Nothing Then Set mSignData = CreateObject("Scripting.Dictionary")

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("all") = CreateObject("Scripting.Dictionary")
    Set d("acc") = CreateObject("Scripting.Dictionary")
    Set d("lea") = CreateObject("Scripting.Dictionary")
    Set d("dates") = CreateObject("Scripting.Dictionary")

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    Dim maxDate As Double
    maxDate = 0
    For r = 1 To n
        Dim dte As Variant
        dte = modAggregate.CellRaw(r, "date")
        If IsNumeric(dte) Then
            If CDbl(dte) > maxDate Then maxDate = CDbl(dte)
        End If

        If modAggregate.CellText(r, "direction") = dir Then
            Dim num As String
            num = modAggregate.CellText(r, "number")
            If num <> "" Then
                d("all")(num) = True
                If IsNumeric(dte) Then
                    If Not d("dates").Exists(num) Then
                        d("dates")(num) = CDbl(dte)
                    ElseIf CDbl(dte) < CDbl(d("dates")(num)) Then
                        d("dates")(num) = CDbl(dte)
                    End If
                End If
                Dim armT As String
                armT = modAggregate.CellText(r, "arm")
                If armT = "ПК" Or armT = "ПЛАНШЕТ" Then
                    Dim rf As String
                    rf = modAggregate.CellText(r, "ready_for")
                    If IsStatusReadyToAccept(rf) Then d("acc")(num) = True
                    If IsStatusReadyToLeave(rf) Then d("lea")(num) = True
                End If
            End If
        End If
    Next r
    d("maxDate") = maxDate
    Set mSignData(dir) = d
    mSignReady(dir) = True
End Sub

' {{BLOCK_SIGNSTAT_<DIR>}} - пять чисел по дирекции (весь период выгрузки, без отсечения «НЕ ПОДПИСАНО»).
Private Function BuildSignStat(dir As String) As String
    EnsureSignData dir
    Dim d As Object
    Set d = mSignData(dir)

    Dim both As Double, accOnly As Double, leaOnly As Double, none As Double
    both = 0: accOnly = 0: leaOnly = 0: none = 0
    Dim k As Variant
    For Each k In d("all").Keys
        Dim hasAcc As Boolean, hasLea As Boolean
        hasAcc = d("acc").Exists(k)
        hasLea = d("lea").Exists(k)
        If hasAcc And hasLea Then
            both = both + 1
        ElseIf hasAcc Then
            accOnly = accOnly + 1
        ElseIf hasLea Then
            leaOnly = leaOnly + 1
        Else
            none = none + 1
        End If
    Next k

    Dim html As String
    html = "<div class='kpis'>"
    html = html & KpiTile("Всего заказ-нарядов", FmtInt(CDbl(d("all").Count)), "", "", "flat", "")
    html = html & KpiTile("Подписаны полностью", FmtInt(both), "", "", "flat", "")
    html = html & KpiTile("Только «Готов к приемке»", FmtInt(accOnly), "", "", "flat", "")
    html = html & KpiTile("Только «Готов к выбытию»", FmtInt(leaOnly), "", "", "flat", "")
    html = html & KpiTile("Не подписаны ни разу", FmtInt(none), "", "", "flat", "")
    html = html & "</div>"
    html = html & CalcNote("Весь период выгрузки, направление " & Esc(dir) & ", без отсечения «НЕ ПОДПИСАНО» (они и есть предмет). " & _
        "«Полностью» — оба статуса с arm из {ПК, ПЛАНШЕТ}; «ни разу» — ни одного статуса с arm из {ПК, ПЛАНШЕТ}.")
    BuildSignStat = html
End Function

' {{BLOCK_UNSIGNED_AGE_<DIR>}} - старение неподписанных статусов по дирекции (ось — date).
Private Function BuildUnsignedAgeByDir(dir As String) As String
    EnsureSignData dir
    Dim d As Object
    Set d = mSignData(dir)
    Dim maxDate As Double
    maxDate = CDbl(d("maxDate"))
    If maxDate <= 0 Then BuildUnsignedAgeByDir = EmptyNote(): Exit Function

    ' Корзины: до 1 дня, 1-7, 7-30, >30 (постановка §4.4).
    Dim labels() As Variant, unAcc() As Double, unLea() As Double
    ReDim labels(1 To 4): ReDim unAcc(1 To 4): ReDim unLea(1 To 4)
    labels(1) = "до 1 дня": labels(2) = "1–7 дней": labels(3) = "7–30 дней": labels(4) = "> 30 дней"
    Dim i As Long
    For i = 1 To 4: unAcc(i) = 0: unLea(i) = 0: Next i

    Dim k As Variant
    For Each k In d("all").Keys
        Dim ageDays As Double
        ageDays = 0
        If d("dates").Exists(k) Then ageDays = maxDate - CDbl(d("dates")(k))
        Dim bi As Long
        If ageDays < 1 Then
            bi = 1
        ElseIf ageDays < 7 Then
            bi = 2
        ElseIf ageDays < 30 Then
            bi = 3
        Else
            bi = 4
        End If
        If Not d("acc").Exists(k) Then unAcc(bi) = unAcc(bi) + 1
        If Not d("lea").Exists(k) Then unLea(bi) = unLea(bi) + 1
    Next k

    Dim html As String
    html = SvgBarsH(labels, unAcc, unLea, -1)
    html = html & CalcNote("Возраст наряда от даты создания (date) до конца выгрузки; серии — «не подписана приёмка» и «не подписано выбытие»; ось — date, не yearWeek (у неподписанных yearWeek = null).")
    BuildUnsignedAgeByDir = html
End Function

' =====================================================================================
' Блоки слайда 4 (ТЗ v1.2, T7). Ось времени - ТОЛЬКО date (у неподписанных yearWeek = null).
' =====================================================================================
' Один проход по снимку: all/signed/rows/dates/noPost/zn по нарядам + счётчик событий
' «НЕ ПОДПИСАНО» + max(date).
Private Sub EnsureUnsignedData()
    If mUnsignedReady Then Exit Sub
    EnsureSnapshot

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("all") = CreateObject("Scripting.Dictionary")
    Set d("signed") = CreateObject("Scripting.Dictionary")
    Set d("rows") = CreateObject("Scripting.Dictionary")
    Set d("dates") = CreateObject("Scripting.Dictionary")
    Set d("noPost") = CreateObject("Scripting.Dictionary")
    Set d("zn") = CreateObject("Scripting.Dictionary")

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    Dim unsEvents As Double, maxDate As Double
    unsEvents = 0: maxDate = 0
    For r = 1 To n
        Dim dte As Variant
        dte = modAggregate.CellRaw(r, "date")
        If IsNumeric(dte) Then
            If CDbl(dte) > maxDate Then maxDate = CDbl(dte)
        End If
        Dim num As String
        num = modAggregate.CellText(r, "number")
        If num = "" Then GoTo NextUnsRow
        d("all")(num) = True
        AddCnt d("rows"), num, 1
        If IsNumeric(dte) Then
            If Not d("dates").Exists(num) Then
                d("dates")(num) = CDbl(dte)
            ElseIf CDbl(dte) < CDbl(d("dates")(num)) Then
                d("dates")(num) = CDbl(dte)
            End If
        End If
        If modAggregate.CellText(r, "postN") = "" Then d("noPost")(num) = True
        If Not d("zn").Exists(num) Then d("zn")(num) = modAggregate.CellText(r, "zn_type")
        Dim armT As String
        armT = modAggregate.CellText(r, "arm")
        If armT = "ПК" Or armT = "ПЛАНШЕТ" Then AddCnt d("signed"), num, 1
        If armT = "НЕ ПОДПИСАНО" Then unsEvents = unsEvents + 1
NextUnsRow:
    Next r
    d("unsEvents") = unsEvents
    d("maxDate") = maxDate
    Set mUnsigned = d
    mUnsignedReady = True
End Sub

' {{KPI_UNSIGNED}} - четыре KPI + плашка «Что это значит для отчёта» (§6.5, числа считаются).
Private Function BuildUnsignedKpi() As String
    EnsureUnsignedData
    Dim d As Object
    Set d = mUnsigned

    ' «Без единой подписи» = ни одной строки с arm из {ПК, ПЛАНШЕТ}; «частично» = есть подписанные и есть неподписанные.
    Dim none As Double, part As Double
    none = 0: part = 0
    Dim k As Variant
    For Each k In d("all").Keys
        Dim s As Double, rows As Double
        s = DictVal(d("signed"), CStr(k))
        rows = DictVal(d("rows"), CStr(k))
        If s = 0 Then
            none = none + 1
        ElseIf rows - s >= 1 Then
            part = part + 1
        End If
    Next k

    ' Старше двух недель: из «без единой подписи», возраст от date до max(date) > 14 дней.
    Dim older As Double
    older = 0
    For Each k In d("all").Keys
        If DictVal(d("signed"), CStr(k)) = 0 Then
            If d("dates").Exists(k) Then
                If CDbl(d("maxDate")) - CDbl(d("dates")(k)) > 14 Then older = older + 1
            End If
        End If
    Next k

    Dim totalNum As Double, totalRows As Double
    totalNum = CDbl(d("all").Count)
    totalRows = CDbl(modAggregate.RowCount())

    Dim html As String
    html = "<div class='kpis'>"
    html = html & KpiTile("ЗН без единой подписи", FmtInt(none), "· " & FmtPct(SafePercentNo(totalNum, none)) & " %", "", "flat", "")
    html = html & KpiTile("ЗН подписаны частично", FmtInt(part), "", "", "flat", "")
    html = html & KpiTile("Событий без подписи", FmtInt(CDbl(d("unsEvents"))), "· " & FmtPct(SafePercentNo(totalRows, CDbl(d("unsEvents")))) & " %", "", "flat", "")
    html = html & KpiTile("Старше двух недель", FmtInt(older), "", "", "flat", "")
    html = html & "</div>"
    html = html & CalcNote("ЗН без единой подписи — ни одна строка наряда не имеет arm из {ПК, ПЛАНШЕТ}; частично — часть строк подписана, часть нет; " & _
        "событий без подписи — строки arm = НЕ ПОДПИСАНО; «старше» — из «без единой подписи», возраст от date до конца выгрузки > 14 дней.")

    ' Плашка «Что это значит для отчёта» (постановка §6.5): оба процента считаются, не хардкодятся.
    Dim armCnt As Object
    Set armCnt = modAggregate.GroupCount(Array("arm"), FBase())
    Dim tabEv As Double, pcEv As Double
    tabEv = DictVal(armCnt, "ПЛАНШЕТ|")
    pcEv = DictVal(armCnt, "ПК|")
    If tabEv + pcEv > 0 Then
        Dim pctFiltered As Double, pctAll As Double
        pctFiltered = tabEv / (tabEv + pcEv)
        pctAll = SafePercentNo(totalRows, tabEv)
        html = html & "<div class='note'><strong>Что это значит для отчёта.</strong> При фильтре arm " & ChrW$(&H2208) & " {ПК, ПЛАНШЕТ} отчёт показывает " & _
            FmtPct(pctFiltered) & " % и выглядит удовлетворительно; с учётом неподписанных доля событий, прошедших через планшет, — " & _
            FmtPct(pctAll) & " %. Оба числа верные, но отвечают на разные вопросы: первое — «чем подписывают», второе — «подписывают ли вообще».</div>"
    End If
    BuildUnsignedKpi = html
End Function

' {{BLOCK_UNSIGNED_AGE}} - старение: корзины 0-7, 8-14, 15-30, 31-90, >90 дней, две серии.
Private Function BuildUnsignedAging() As String
    EnsureUnsignedData
    Dim d As Object
    Set d = mUnsigned
    Dim maxDate As Double
    maxDate = CDbl(d("maxDate"))
    If maxDate <= 0 Then BuildUnsignedAging = EmptyNote(): Exit Function

    Dim labels() As Variant, none() As Double, part() As Double
    ReDim labels(1 To 5): ReDim none(1 To 5): ReDim part(1 To 5)
    labels(1) = "0–7 дней": labels(2) = "8–14": labels(3) = "15–30": labels(4) = "31–90": labels(5) = "> 90"
    Dim i As Long
    For i = 1 To 5: none(i) = 0: part(i) = 0: Next i

    Dim k As Variant
    For Each k In d("all").Keys
        Dim age As Double
        age = 0
        If d("dates").Exists(k) Then age = maxDate - CDbl(d("dates")(k))
        Dim bi As Long
        If age <= 7 Then
            bi = 1
        ElseIf age <= 14 Then
            bi = 2
        ElseIf age <= 30 Then
            bi = 3
        ElseIf age <= 90 Then
            bi = 4
        Else
            bi = 5
        End If
        Dim s As Double, rows As Double
        s = DictVal(d("signed"), CStr(k))
        rows = DictVal(d("rows"), CStr(k))
        If s = 0 Then
            none(bi) = none(bi) + 1
        ElseIf rows - s >= 1 Then
            part(bi) = part(bi) + 1
        End If
    Next k

    Dim html As String
    html = SvgBarsH(labels, none, part, -1)
    html = html & CalcNote("Возраст от date до конца выгрузки; серии — «без единой подписи» и «подписан частично». " & _
        "Чем старше корзина, тем она больше: событие прошло, подписывать нечего.")
    BuildUnsignedAging = html
End Function

' {{BLOCK_UNSIGNED_SOURCE}} - «откуда берутся»: неподписанные с пустым постом.
Private Function BuildUnsignedSource() As String
    EnsureUnsignedData
    Dim d As Object
    Set d = mUnsigned
    Dim none As Double, noneNoPost As Double
    none = 0: noneNoPost = 0
    Dim k As Variant
    For Each k In d("all").Keys
        If DictVal(d("signed"), CStr(k)) = 0 Then
            none = none + 1
            If d("noPost").Exists(k) Then noneNoPost = noneNoPost + 1
        End If
    Next k
    If none = 0 Then BuildUnsignedSource = EmptyNote(): Exit Function

    Dim html As String
    html = "<p class='lede'><strong>" & FmtInt(noneNoPost) & " из " & FmtInt(none) & "</strong> неподписанных нарядов имеют пустой пост.</p>"
    html = html & CalcNote("«Без поста» и «не подписано» — один дефект, а не два: наряд, не привязанный к посту ремзоны, не подписывает никто.")
    BuildUnsignedSource = html
End Function

' {{BLOCK_UNSIGNED_ZNTYPE}} - неподписанные наряды по виду ремонта.
Private Function BuildUnsignedByZnType() As String
    EnsureUnsignedData
    Dim d As Object
    Set d = mUnsigned

    Dim cnt As Object
    Set cnt = CreateObject("Scripting.Dictionary")
    Dim none As Double
    none = 0
    Dim k As Variant
    For Each k In d("all").Keys
        If DictVal(d("signed"), CStr(k)) = 0 Then
            none = none + 1
            Dim zt As String
            zt = ""
            If d("zn").Exists(k) Then zt = CStr(d("zn")(k))
            If zt = "" Then zt = "(не указан)"
            If cnt.Exists(zt) Then cnt(zt) = CDbl(cnt(zt)) + 1 Else cnt(zt) = 1
        End If
    Next k
    If cnt.Count = 0 Then BuildUnsignedByZnType = EmptyNote(): Exit Function

    Dim sorted As Variant
    sorted = modAggregate.SortDictionaryKeysByValue(cnt, True)
    Dim html As String, i As Long
    html = "<table class='block-table'><thead><tr><th>Вид ремонта</th><th>ЗН</th><th>%</th></tr></thead><tbody>"
    For i = LBound(sorted) To UBound(sorted)
        Dim nameVal As String
        nameVal = CStr(sorted(i))
        html = html & "<tr><td>" & Esc(nameVal) & "</td><td class='num'>" & FmtInt(CDbl(cnt(nameVal))) & "</td>"
        If none > 0 Then
            html = html & "<td class='num'>" & FmtPct(CDbl(cnt(nameVal)) / none) & "</td></tr>"
        Else
            html = html & "<td class='num'>—</td></tr>"
        End If
    Next i
    html = html & "</tbody></table>"
    html = html & CalcNote("Наряды без единой подписи (ни в одной строке нет arm из {ПК, ПЛАНШЕТ}) по виду ремонта; % — от числа неподписанных нарядов.")
    BuildUnsignedByZnType = html
End Function

' =====================================================================================
' 12. ParseAIResponse - двухшаговый разбор ответа Chat Completions (P0-4).
'     Четыре ключа slide1..slide4_conclusions складываются в module-level кэш mInsights
'     (ТЗ v1.2, T9.2). Контракт Core не меняется: ByRef slide3/4/5 наружу возвращают
'     выводы слайдов 2/3/4, вывод слайда 1 - только через кэш mInsights("slide1").
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
    For i = 1 To 8
        v = JsonUnescape(ExtractJsonStringValue(payload, "slide" & CStr(i) & "_conclusions"))
        If v <> "" Then
            ins("slide" & CStr(i)) = v
        Else
            ins("slide" & CStr(i)) = AI_FALLBACK
            ok = False
            missing = missing & "slide" & CStr(i) & ", "
        End If
    Next i

    ' Контракт Core: параметрами наружу - выводы слайдов 2/3/4, слайд 1 - только кэш.
    slide3 = ins("slide2")
    slide4 = ins("slide3")
    slide5 = ins("slide4")

    If ok Then
        modLog.WriteDebug 1, "Формирование отчёта", "ParseAIResponse", _
            "Распознаны все 8 ключей slide1..slide8_conclusions"
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

    ' --- Шапка и подвал ---
    Dim rw As Long
    rw = modContentZone.ZoneReportWeek()
    d("REPORT_TITLE") = "Отчёт МТО"
    d("REPORT_WEEK_LABEL") = modContentZone.WeekCaption(rw)
    d("REPORT_LEDE") = BuildLede(rw)
    d("FACTS") = BuildFactsRef()
    d("REPORT_FOOTER") = BuildFooter(rw)

    ' --- Слайды 1 и 5-8: часть «Техника» ---
    modContentZone.FillZonePlaceholders d
    ' --- Слайды 2-4: часть «Дисциплина» ---
    modContentDisc.FillDiscPlaceholders d

    ' --- Выводы ИИ: кэш при mInsightsReady, иначе fallback.
    ' Контракт Core не менялся: aiSlide3/4/5 - выводы слайдов 2/3/4.
    Dim ai As Object
    Set ai = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To 8
        ai("slide" & CStr(i)) = AI_FALLBACK
    Next i
    If mInsightsReady Then
        For i = 1 To 8
            If mInsights.Exists("slide" & CStr(i)) Then
                ai("slide" & CStr(i)) = CStr(mInsights("slide" & CStr(i)))
            End If
        Next i
    Else
        ai("slide2") = aiSlide3
        ai("slide3") = aiSlide4
        ai("slide4") = aiSlide5
    End If
    For i = 1 To 8
        ' Обратная замена псевдонимов «Сотрудник N» -> ФИО, затем экранирование.
        d("AI_INSIGHT_SLIDE_" & CStr(i)) = AiList(DeAlias(CStr(ai("slide" & CStr(i)))))
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

' Подзаголовок шапки: из чего собран отчёт и какая неделя отчётная.
Private Function BuildLede(ByVal rw As Long) As String
    Dim a As Double, b As Double
    a = modContentZone.SnapFrom()
    b = modContentZone.SnapTo()
    Dim per As String
    If a > 0 And b > 0 Then
        per = Format$(CDate(a), "dd.mm.yyyy") & " " & ChrW$(&H2013) & " " & _
            Format$(CDate(b), "dd.mm.yyyy")
    Else
        per = "период не определён"
    End If
    BuildLede = "Восемь слайдов в двух частях: дисциплина подписания на планшете " & _
        "(слайды 1" & ChrW$(&H2013) & "4) и операционка ремзоны (слайды 5" & _
        ChrW$(&H2013) & "8). Числа посчитаны на выгрузке 1С за " & per & _
        "; отчётная неделя " & ChrW$(&H2014) & " " & modContentZone.WLab(rw) & _
        " (" & modContentZone.WeekRange(rw) & "), последняя полная неделя снимка."
End Function

' Четыре числа шапки. Считает modContentZone - чтобы шапка не разошлась со слайдами.
Private Function BuildFactsRef() As String
    Dim h As String
    h = "<div><dt>Событий</dt><dd class=""num"">" & _
        FmtInt(modContentZone.EventsCount()) & "</dd></div>"
    h = h & "<div><dt>Нарядов</dt><dd class=""num"">" & _
        FmtInt(modContentZone.OrdersCount()) & "</dd></div>"
    h = h & "<div><dt>Машин в парке</dt><dd class=""num"">" & _
        FmtInt(modContentZone.FleetCount()) & "</dd></div>"
    h = h & "<div><dt>Без поста</dt><dd class=""num"">" & _
        modContentZone.Pc(modContentZone.NoPostPct(), 1) & "</dd></div>"
    BuildFactsRef = h
End Function

Private Function BuildFooter(ByVal rw As Long) As String
    BuildFooter = "Отчёт МТО " & ChrW$(&HB7) & " " & modContentZone.WeekCaption(rw) & _
        " " & ChrW$(&HB7) & " автономный HTML: шрифты и графика встроены, внешних " & _
        "запросов нет. Часть 1 " & ChrW$(&H2014) & " трек А (событие подписания), " & _
        "часть 2 " & ChrW$(&H2014) & " трек Б (наряд, заезд, машина). Единицы счёта " & _
        "разных треков не складываются."
End Function

' Вывод ИИ -> список. Модель отдаёт текст; в шаблоне на этом месте <ul>.
Private Function AiList(ByVal t As String) As String
    Dim body As String
    body = Replace$(Replace$(CStr(t), vbCrLf, vbLf), vbCr, vbLf)
    Dim parts As Variant, i As Long, out As String
    parts = Split(body, vbLf)
    out = ""
    For i = LBound(parts) To UBound(parts)
        Dim ln As String
        ln = Trim$(CStr(parts(i)))
        ' Модель иногда ставит маркер списка сама - убираем, разметку даёт шаблон.
        Do While Len(ln) > 0
            Dim c1 As String
            c1 = Left$(ln, 1)
            If c1 = "-" Or c1 = "*" Or c1 = ChrW$(&H2022) Then
                ln = Trim$(Mid$(ln, 2))
            Else
                Exit Do
            End If
        Loop
        If Len(ln) > 0 Then out = out & "<li>" & Esc(ln) & "</li>"
    Next i
    If out = "" Then out = "<li>" & Esc(AI_FALLBACK) & "</li>"
    AiList = "<ul>" & out & "</ul>"
End Function

' Цвет рамки процента: 0 -> красный, 50 -> оранжевый, 75 -> оливковый, 100 -> зелёный.
' Шкала из docs/plans/MTO_контракт_шаблона_v1.0.md; Core (modColor) не менялся.
Public Function PctBorderColor(ByVal p As Double) As String
    Dim stopsP As Variant, stopsC As Variant
    stopsP = Array(0#, 50#, 75#, 100#)
    stopsC = Array("#e2483a", "#e9822a", "#9aa93a", "#1faf6a")
    Dim v As Double
    v = p
    If v < 0# Then v = 0#
    If v > 100# Then v = 100#
    Dim i As Long
    For i = 0 To 2
        If v <= CDbl(stopsP(i + 1)) Then
            Dim t As Double
            t = (v - CDbl(stopsP(i))) / (CDbl(stopsP(i + 1)) - CDbl(stopsP(i)))
            PctBorderColor = LCase$(modColor.InterpolateHex(CStr(stopsC(i)), _
                CStr(stopsC(i + 1)), t))
            Exit Function
        End If
    Next i
    PctBorderColor = "#1faf6a"
End Function

' Подпись периода выгрузки для шапки ({{REPORT_PERIOD}}).
Private Function ReportPeriodCaption() As String
    EnsureSnapshot
    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    Dim dMin As Double, dMax As Double
    dMin = 0: dMax = 0
    For r = 1 To n
        Dim dte As Variant
        dte = modAggregate.CellRaw(r, "date")
        If IsNumeric(dte) Then
            If dMin = 0 Or CDbl(dte) < dMin Then dMin = CDbl(dte)
            If CDbl(dte) > dMax Then dMax = CDbl(dte)
        End If
    Next r
    If dMin = 0 Then ReportPeriodCaption = "": Exit Function
    ReportPeriodCaption = "Выгрузка " & Format(CDate(dMin), "dd.mm.yyyy") & " — " & Format(CDate(dMax), "dd.mm.yyyy")
End Function

' «неделя 202635 (нед. 35/2026)» для {{REPORT_WEEK_LABEL}}.
Private Function WeekLabelCaption(rw As Long) As String
    If rw = 0 Then WeekLabelCaption = "отчётная неделя не определена": Exit Function
    WeekLabelCaption = "неделя " & CStr(rw) & " (нед. " & CStr(rw Mod 100) & "/" & CStr(rw \ 100) & ")"
End Function

' Отладочная сверка плейсхолдеров: каждый {{...}} шаблона есть в словаре и наоборот (приёмка T9).
' templatePath: пусто -> ThisWorkbook.Path & "\tmp_index.html". Результат - в Immediate.
Public Sub DebugCheckPlaceholders(Optional templatePath As String = "")
    Dim p As String
    p = templatePath
    If Trim$(p) = "" Then p = ThisWorkbook.Path & "\tmp_index.html"

    Dim html As String
    On Error GoTo ErrRead
    html = modHTMLEngine.ReadUtf8(p)
    On Error GoTo 0

    Dim tpl As Object
    Set tpl = CreateObject("Scripting.Dictionary")
    Dim pos As Long
    pos = InStr(html, "{{")
    Do While pos > 0
        Dim posEnd As Long
        posEnd = InStr(pos + 2, html, "}}")
        If posEnd = 0 Then Exit Do
        Dim name As String
        name = Mid$(html, pos + 2, posEnd - pos - 2)
        tpl(name) = True
        pos = InStr(posEnd + 2, html, "{{")
    Loop

    Dim d As Object
    Set d = BuildPlaceholders(AI_FALLBACK, AI_FALLBACK, AI_FALLBACK)

    Dim missing As String, extra As String
    missing = "": extra = ""
    Dim k As Variant
    For Each k In tpl.Keys
        If Not d.Exists(k) Then missing = missing & k & ", "
    Next k
    For Each k In d.Keys
        If Not tpl.Exists(k) Then extra = extra & k & ", "
    Next k

    If missing = "" And extra = "" Then
        Debug.Print "DebugCheckPlaceholders: OK (" & d.Count & " плейсхолдеров, расхождений нет)"
    Else
        Debug.Print "В шаблоне, но НЕ в словаре: " & IIf(missing = "", "нет", missing)
        Debug.Print "В словаре, но НЕ в шаблоне: " & IIf(extra = "", "нет", extra)
    End If
    Exit Sub
ErrRead:
    Debug.Print "DebugCheckPlaceholders: не удалось прочитать шаблон " & p & " - " & Err.Description
End Sub

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
