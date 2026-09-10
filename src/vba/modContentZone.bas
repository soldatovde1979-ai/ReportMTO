Attribute VB_Name = "modContentZone"
' modContentZone - CONTENT-слой части «Техника» (слайды 5-8) отчёта МТО.
'
' Версия 2.0 от 10.09.2026:
'   - написаны все блоки слайдов 5-8; разметка снята с эталона
'     temp/MTO_макет_отчета_v4.0.html, включая пояснения под блоками;
'   - классификатор описаний дефекта (характер работы + узел) перенесён в VBA;
'   - возвраты: помесячно с начала года и понедельно за 8 недель, окно 30 суток;
'   - неделя и месяц считаются ЗДЕСЬ из поля date, а не берутся из yearWeek:
'     в книге с устаревшим Power Query yearWeek приходит нулём, и ось времени
'     умирает молча. Своя ISO-неделя снимает эту зависимость.
' Версия 1.0 от 10.09.2026:
'   - первый выпуск: парк и заезды, три уровня модели данных за ОДИН проход.
'
' Core не изменялся: нужен только снимок modAggregate через CellRaw/CellText.
'
' Разметка блоков обязана совпадать с docs/plans/MTO_контракт_шаблона_v1.0.md
' и с эталоном. Классы вне словаря контракта не использовать: в шаблоне их нет.
'
' Типы полей (fnNormalizeFields v7): date и status_date приведены к datetime,
' а zn_closed и MadeYear остаются ТЕКСТОМ из JSON - разбираются здесь ISO-парсером.
'
' defect_desc читается только внутри модуля (классификатор, список повторов).
' Во внешнюю модель это поле не уходит: в нём встречаются ФИО.
'
' Файл хранится в UTF-8 + CRLF. Символы вне ANSI-1251 запрещены: при импорте в VBE
' (UTF-8 -> 1251) они превращаются в "?". Проверка: tools/vba_lint_v1.0/vba_lint.py.
Option Explicit

' Поля записи наряда (mZn: number -> Variant-массив)
Private Const Z_DATE As Long = 0        ' дата создания ЗН, серийная
Private Const Z_TYPE As Long = 1        ' zn_type
Private Const Z_DEFEKT As Long = 2      ' defekt_type (группа дефекта)
Private Const Z_VEH As Long = 3         ' vehicle_number, уникальный код машины
Private Const Z_VGROUP As Long = 4      ' vehicle_group
Private Const Z_MADE As Long = 5        ' MadeYear, серийная
Private Const Z_PARTS As Long = 6       ' cost_parts
Private Const Z_CLOSED As Long = 7      ' zn_closed, серийная
Private Const Z_TEK As Long = 8         ' TekStatusPoDoc
Private Const Z_POST As Long = 9        ' post
Private Const Z_OWNER As Long = 10      ' owner_dep
Private Const Z_ACCG As Long = 11       ' «Готов к приемке», ДГМ
Private Const Z_ACCD As Long = 12       ' «Готов к приемке», ДЭНТ
Private Const Z_LEVG As Long = 13       ' «Готов к выбытию», ДГМ
Private Const Z_LEVD As Long = 14       ' «Готов к выбытию», ДЭНТ
Private Const Z_SIGNED As Long = 15     ' сколько из 4 событий подписаны
Private Const Z_DESC As Long = 16       ' defect_desc, наружу не отдаётся
Private Const Z_KIND As Long = 17       ' характер работы (классификатор)
Private Const Z_NODE As Long = 18       ' узел внутри группы (классификатор)
Private Const Z_WEEK As Long = 19       ' ISO год*100 + неделя по date
Private Const Z_MONTH As Long = 20      ' год*100 + месяц по date
Private Const Z_FIELDS As Long = 21

' Поля записи машины (mVeh: vehicle_number -> Variant-массив)
Private Const V_ZN As Long = 0
Private Const V_HOURS As Long = 1
Private Const V_PARTS As Long = 2
Private Const V_VISITS As Long = 3
Private Const V_MADE As Long = 4
Private Const V_GROUP As Long = 5
Private Const V_FIELDS As Long = 6

Private Const GAP_HOURS As Double = 12#     ' порог склейки заездов, гипотеза (слайд 5)
Private Const MAX_DUR_H As Double = 8760#   ' > года между созданием и закрытием - брак выгрузки
Private Const RET_WINDOW As Long = 30       ' окно возврата, суток
Private Const UNCLS As String = "(не классифицировано)"
Private Const EMPTYD As String = "(описание пустое)"
Private Const NOSECT As String = "(раздел не указан)"

Private mReady As Boolean
Private mZn As Object            ' number -> Variant(Z_FIELDS)
Private mVeh As Object           ' veh -> Variant(V_FIELDS)
Private mVisitSizes As Object    ' порог, ч -> Collection размеров заездов
Private mVisitWeeks As Object    ' ISO-неделя первого наряда заезда -> число заездов
Private mBadClosed As Long       ' нарядов с zn_closed раньше date или > года
Private mRows As Long            ' событий в снимке
Private mArmNone As Long         ' событий «НЕ ПОДПИСАНО»
Private mNoCounter As Long       ' событий, где odometer и engine_hours пусты
Private mNoBounds As Long        ' событий с in_bounds = ложь
Private mSnapEnd As Double
Private mKindRules As Collection
Private mNodeRules As Object     ' группа -> Collection(Array(имя, «|ключ|ключ|»))
Private mClsDone As Boolean
Private mWeeks As Variant        ' 8 недель окна
Private mSignTot As Object       ' ISO-неделя даты статуса -> подписей с известным АРМ
Private mSignTab As Object       ' то же, только с планшета
Private mRepWeek As Long

' =====================================================================================
' Сброс и построение уровней
' =====================================================================================
Public Sub ResetZone()
    mReady = False
    mClsDone = False
    Set mZn = Nothing
    Set mVeh = Nothing
    Set mVisitSizes = Nothing
    Set mVisitWeeks = Nothing
    Set mNodeRules = Nothing
    Set mKindRules = Nothing
    Set mSignTot = Nothing
    Set mSignTab = Nothing
    mBadClosed = 0
    mSnapEnd = 0#
    mRepWeek = 0
    mWeeks = Empty
    mRetReady = False
    mFlowReady = False
    Set mCloseH = Nothing
    Set mStuck = Nothing
    Set mHang = Nothing
    Set mTail = Nothing
End Sub

' Один проход по снимку: события сворачиваются в наряды.
Private Sub EnsureZn()
    If mReady Then Exit Sub
    If Not modAggregate.IsReady() Then
        Err.Raise vbObjectError + 41, , "modContentZone: снимок tbDATA не создан"
    End If

    Set mZn = CreateObject("Scripting.Dictionary")
    Set mSignTot = CreateObject("Scripting.Dictionary")
    Set mSignTab = CreateObject("Scripting.Dictionary")

    Dim hasVeh As Boolean, hasParts As Boolean, hasClosed As Boolean
    Dim hasMade As Boolean, hasTek As Boolean, hasGroup As Boolean, hasOwner As Boolean
    Dim hasDesc As Boolean, hasOdo As Boolean, hasEng As Boolean, hasBounds As Boolean
    hasVeh = modAggregate.HasColumn("vehicle_number")
    hasParts = modAggregate.HasColumn("cost_parts")
    hasClosed = modAggregate.HasColumn("zn_closed")
    hasMade = modAggregate.HasColumn("MadeYear")
    hasTek = modAggregate.HasColumn("TekStatusPoDoc")
    hasGroup = modAggregate.HasColumn("vehicle_group")
    hasOwner = modAggregate.HasColumn("owner_dep")
    hasDesc = modAggregate.HasColumn("defect_desc")
    hasOdo = modAggregate.HasColumn("odometer")
    hasEng = modAggregate.HasColumn("engine_hours")
    hasBounds = modAggregate.HasColumn("in_bounds")

    mArmNone = 0: mNoCounter = 0: mNoBounds = 0

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    mRows = n

    For r = 1 To n
        Dim num As String
        num = modAggregate.CellText(r, "number")
        If Len(num) > 0 Then
            Dim z As Variant
            If mZn.Exists(num) Then
                z = mZn(num)
            Else
                ReDim z(0 To Z_FIELDS - 1)
                Dim dser As Double
                dser = ToSerial(modAggregate.CellRaw(r, "date"))
                z(Z_DATE) = dser
                z(Z_WEEK) = IsoYearWeek(dser)
                z(Z_MONTH) = YearMonth(dser)
                z(Z_TYPE) = modAggregate.CellText(r, "zn_type")
                z(Z_DEFEKT) = modAggregate.CellText(r, "defekt_type")
                z(Z_POST) = modAggregate.CellText(r, "post")
                z(Z_ACCG) = 0#: z(Z_ACCD) = 0#: z(Z_LEVG) = 0#: z(Z_LEVD) = 0#
                z(Z_SIGNED) = 0#
                z(Z_KIND) = "": z(Z_NODE) = ""
                If hasDesc Then z(Z_DESC) = modAggregate.CellText(r, "defect_desc") Else z(Z_DESC) = ""
                If hasVeh Then z(Z_VEH) = Trim$(modAggregate.CellText(r, "vehicle_number")) Else z(Z_VEH) = ""
                If hasGroup Then z(Z_VGROUP) = modAggregate.CellText(r, "vehicle_group") Else z(Z_VGROUP) = ""
                If hasTek Then z(Z_TEK) = modAggregate.CellText(r, "TekStatusPoDoc") Else z(Z_TEK) = ""
                If hasOwner Then z(Z_OWNER) = modAggregate.CellText(r, "owner_dep") Else z(Z_OWNER) = ""
                If hasMade Then z(Z_MADE) = ToSerial(modAggregate.CellRaw(r, "MadeYear")) Else z(Z_MADE) = 0#
                If hasClosed Then z(Z_CLOSED) = ToSerial(modAggregate.CellRaw(r, "zn_closed")) Else z(Z_CLOSED) = 0#
                If hasParts Then z(Z_PARTS) = ToNum(modAggregate.CellRaw(r, "cost_parts")) Else z(Z_PARTS) = 0#
            End If

            Dim sd As Double
            sd = ToSerial(modAggregate.CellRaw(r, "status_date"))
            If sd > 0# Then
                Dim rf As String, dr As String, fld As Long
                rf = modAggregate.CellText(r, "ready_for")
                dr = modAggregate.CellText(r, "direction")
                fld = -1
                ' Статусы в выгрузке пишутся через «е»: «Готов к приемке» / «Готов к выбытию».
                If InStr(1, rf, "приемке", vbTextCompare) > 0 Then
                    If InStr(1, dr, "ДГМ", vbTextCompare) > 0 Then fld = Z_ACCG Else fld = Z_ACCD
                ElseIf InStr(1, rf, "выбытию", vbTextCompare) > 0 Then
                    If InStr(1, dr, "ДГМ", vbTextCompare) > 0 Then fld = Z_LEVG Else fld = Z_LEVD
                End If
                If fld >= 0 Then
                    If CDbl(z(fld)) = 0# Or sd < CDbl(z(fld)) Then z(fld) = sd
                End If
            End If

            Dim arm As String
            arm = modAggregate.CellText(r, "arm")
            If arm = "ПК" Or arm = "ПЛАНШЕТ" Then
                z(Z_SIGNED) = CDbl(z(Z_SIGNED)) + 1#
                If sd > 0# Then
                    ' Ось «% планшета» - по дате статуса: подпись относится к той
                    ' неделе, когда её поставили, а не когда завели наряд.
                    AddCnt mSignTot, CStr(IsoYearWeek(sd)), 1#
                    If arm = "ПЛАНШЕТ" Then AddCnt mSignTab, CStr(IsoYearWeek(sd)), 1#
                End If
            Else
                mArmNone = mArmNone + 1
            End If

            If hasOdo And hasEng Then
                If ToNum(modAggregate.CellRaw(r, "odometer")) = 0# _
                   And ToNum(modAggregate.CellRaw(r, "engine_hours")) = 0# Then
                    mNoCounter = mNoCounter + 1
                End If
            End If
            If hasBounds Then
                If Not Truthy(modAggregate.CellRaw(r, "in_bounds")) Then mNoBounds = mNoBounds + 1
            End If

            mZn(num) = z
        End If
    Next r

    mReady = True
    modLog.WriteDebug 2, "Техника", "modContentZone.EnsureZn", _
        "Событий " & CStr(n) & " -> нарядов " & CStr(mZn.Count)
End Sub

' =====================================================================================
' Разбор значений
' =====================================================================================
' Значение ячейки -> серийная дата. Понимает Date, число и текст ISO
' («2026-08-03T11:47:18»): zn_closed и MadeYear приходят из JSON текстом.
' Пустое значение и «нулевая» дата 1С (0001-01-01) -> 0.
Public Function ToSerial(ByVal v As Variant) As Double
    ToSerial = 0#
    If IsEmpty(v) Or IsNull(v) Then Exit Function

    If IsDate(v) Then
        Dim dv As Date
        dv = CDate(v)
        If Year(dv) < 1900 Then Exit Function
        ToSerial = CDbl(dv)
        Exit Function
    End If

    If IsNumeric(v) Then
        Dim d As Double
        d = CDbl(v)
        If d < 1# Then Exit Function          ' до 31.12.1899 - «пусто» в терминах Excel
        ToSerial = d
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) < 10 Then Exit Function
    Dim yy As Long, mm As Long, dd As Long
    If Not IsNumeric(Mid$(s, 1, 4)) Then Exit Function
    yy = CLng(Mid$(s, 1, 4))
    If yy < 1900 Then Exit Function
    If Not IsNumeric(Mid$(s, 6, 2)) Then Exit Function
    mm = CLng(Mid$(s, 6, 2))
    If Not IsNumeric(Mid$(s, 9, 2)) Then Exit Function
    dd = CLng(Mid$(s, 9, 2))
    If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function

    Dim hh As Long, mi As Long, ss As Long
    hh = 0: mi = 0: ss = 0
    If Len(s) >= 19 Then
        If IsNumeric(Mid$(s, 12, 2)) Then hh = CLng(Mid$(s, 12, 2))
        If IsNumeric(Mid$(s, 15, 2)) Then mi = CLng(Mid$(s, 15, 2))
        If IsNumeric(Mid$(s, 18, 2)) Then ss = CLng(Mid$(s, 18, 2))
    End If
    ToSerial = CDbl(DateSerial(yy, mm, dd)) + CDbl(TimeSerial(hh, mi, ss))
End Function

Public Function ToNum(ByVal v As Variant) As Double
    ToNum = 0#
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then ToNum = CDbl(v)
End Function

' Часы «создание -> закрытие». Отрицательные и больше года - брак выгрузки,
' такие наряды в суммы часов не идут и считаются отдельно (реестр качества).
Private Function DurHours(ByVal z As Variant) As Double
    DurHours = -1#
    Dim a As Double, b As Double
    a = CDbl(z(Z_DATE)): b = CDbl(z(Z_CLOSED))
    If a <= 0# Or b <= 0# Then Exit Function
    Dim h As Double
    h = (b - a) * 24#
    If h < 0# Or h > MAX_DUR_H Then
        mBadClosed = mBadClosed + 1
        Exit Function
    End If
    DurHours = h
End Function

' =====================================================================================
' Сортировка и медиана
' =====================================================================================
Public Sub QSortD(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p
            i = i + 1
        Loop
        Do While a(j) > p
            j = j - 1
        Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    QSortD a, lo, j
    QSortD a, i, hi
End Sub

' Сортировка ключей вместе со значениями (для нарядов машины по дате).
Public Sub QSortPair(ByRef k() As Double, ByRef v() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long, p As Double, t As Double, s As String
    i = lo: j = hi
    p = k((lo + hi) \ 2)
    Do While i <= j
        Do While k(i) < p
            i = i + 1
        Loop
        Do While k(j) > p
            j = j - 1
        Loop
        If i <= j Then
            t = k(i): k(i) = k(j): k(j) = t
            s = v(i): v(i) = v(j): v(j) = s
            i = i + 1: j = j - 1
        End If
    Loop
    QSortPair k, v, lo, j
    QSortPair k, v, i, hi
End Sub

' Медиана по коллекции чисел. cnt = 0 -> результат не определён (hasValue = False).
Public Function MedianOf(ByVal col As Collection, ByRef hasValue As Boolean) As Double
    hasValue = False
    MedianOf = 0#
    If col Is Nothing Then Exit Function
    If col.Count = 0 Then Exit Function

    Dim a() As Double
    ReDim a(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        a(i - 1) = CDbl(col(i))
    Next i
    QSortD a, 0, UBound(a)

    Dim n As Long
    n = UBound(a) + 1
    If n Mod 2 = 1 Then
        MedianOf = a((n - 1) \ 2)
    Else
        MedianOf = (a(n \ 2 - 1) + a(n \ 2)) / 2#
    End If
    hasValue = True
End Function

' =====================================================================================
' Машины и заезды
' =====================================================================================
' Заезд - наряды одной машины, у которых разрыв между датами создания не больше
' порога. Порог 12 ч - ГИПОТЕЗА: признака фактического выезда с территории в
' выгрузке нет. Поэтому считаем сразу на трёх порогах и публикуем вместе с ними.
Private Sub EnsureVeh()
    If Not (mVeh Is Nothing) Then Exit Sub
    EnsureZn

    Set mVeh = CreateObject("Scripting.Dictionary")
    Set mVisitSizes = CreateObject("Scripting.Dictionary")
    Set mVisitWeeks = CreateObject("Scripting.Dictionary")

    Dim byVeh As Object
    Set byVeh = CreateObject("Scripting.Dictionary")

    Dim k As Variant, z As Variant
    For Each k In mZn.Keys
        z = mZn(k)
        Dim vn As String
        vn = CStr(z(Z_VEH))
        If Len(vn) > 0 Then
            Dim v As Variant
            If mVeh.Exists(vn) Then
                v = mVeh(vn)
            Else
                ReDim v(0 To V_FIELDS - 1)
                v(V_ZN) = 0#: v(V_HOURS) = 0#: v(V_PARTS) = 0#
                v(V_VISITS) = 0#: v(V_MADE) = 0#: v(V_GROUP) = CStr(z(Z_VGROUP))
            End If
            v(V_ZN) = CDbl(v(V_ZN)) + 1#
            v(V_PARTS) = CDbl(v(V_PARTS)) + CDbl(z(Z_PARTS))
            Dim h As Double
            h = DurHours(z)
            If h >= 0# Then v(V_HOURS) = CDbl(v(V_HOURS)) + h
            If CDbl(v(V_MADE)) = 0# Then v(V_MADE) = CDbl(z(Z_MADE))
            mVeh(vn) = v

            If CDbl(z(Z_DATE)) > 0# Then
                If Not byVeh.Exists(vn) Then byVeh.Add vn, New Collection
                byVeh(vn).Add CStr(k)
            End If
        End If
    Next k

    Dim gaps As Variant
    gaps = Array(8#, 12#, 24#)
    Dim gi As Long
    For gi = LBound(gaps) To UBound(gaps)
        mVisitSizes.Add CStr(gaps(gi)), New Collection
    Next gi

    For Each k In byVeh.Keys
        Dim col As Collection
        Set col = byVeh(k)
        Dim dates() As Double, nums() As String, i As Long
        ReDim dates(0 To col.Count - 1)
        ReDim nums(0 To col.Count - 1)
        For i = 1 To col.Count
            nums(i - 1) = CStr(col(i))
            dates(i - 1) = CDbl(mZn(nums(i - 1))(Z_DATE))
        Next i
        QSortPair dates, nums, 0, UBound(dates)

        For gi = LBound(gaps) To UBound(gaps)
            Dim gap As Double
            gap = CDbl(gaps(gi))
            Dim size As Long, visits As Long, startI As Long
            size = 1: visits = 1: startI = 0
            For i = 1 To UBound(dates)
                If (dates(i) - dates(i - 1)) * 24# <= gap Then
                    size = size + 1
                Else
                    mVisitSizes(CStr(gap)).Add size
                    If gap = GAP_HOURS Then AddCnt mVisitWeeks, CStr(IsoYearWeek(dates(startI))), 1#
                    size = 1: visits = visits + 1: startI = i
                End If
            Next i
            mVisitSizes(CStr(gap)).Add size
            If gap = GAP_HOURS Then
                AddCnt mVisitWeeks, CStr(IsoYearWeek(dates(startI))), 1#
                Dim vv As Variant
                vv = mVeh(CStr(k))
                vv(V_VISITS) = CDbl(visits)
                mVeh(CStr(k)) = vv
            End If
        Next gi
    Next k

    modLog.WriteDebug 2, "Техника", "modContentZone.EnsureVeh", _
        "Машин " & CStr(mVeh.Count) & ", заездов при пороге " & CStr(GAP_HOURS) & _
        " ч: " & CStr(mVisitSizes(CStr(GAP_HOURS)).Count)
End Sub

' Конец снимка - максимальная дата создания наряда.
Public Function SnapshotEnd() As Double
    EnsureZn
    If mSnapEnd > 0# Then SnapshotEnd = mSnapEnd: Exit Function
    Dim k As Variant, mx As Double
    mx = 0#
    For Each k In mZn.Keys
        If CDbl(mZn(k)(Z_DATE)) > mx Then mx = CDbl(mZn(k)(Z_DATE))
    Next k
    mSnapEnd = mx
    SnapshotEnd = mx
End Function

Private Function AgeYears(ByVal made As Double) As Double
    AgeYears = -1#
    If made <= 0# Then Exit Function
    AgeYears = (SnapshotEnd() - made) / 365.25
End Function

' Плановые виды ремонта: у них дефекта нет, пустой defekt_type - норма.
' Парето и кандидаты на повтор считаются только по внеплановым.
Public Function IsPlanned(ByVal znType As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(znType))
    IsPlanned = (Left$(t, 2) = "ТО") Or (Left$(t, 3) = "СТО") Or (Left$(t, 3) = "ЧТО") _
        Or (Left$(t, 3) = "ПТО") Or (Left$(t, 3) = "ГТО") _
        Or (t = "ОБСЛУЖИВАНИЕ ПРИ ВЫПУСКЕ") Or (t = "OMNICOMM")
End Function

' =====================================================================================
' Счётчики и сортировки поверх словарей
' =====================================================================================
Public Sub AddCnt(ByVal d As Object, ByVal key As String, ByVal v As Double)
    If d.Exists(key) Then d(key) = CDbl(d(key)) + v Else d.Add key, v
End Sub

Public Function DictVal(ByVal d As Object, ByVal key As String) As Double
    If d.Exists(key) Then DictVal = CDbl(d(key)) Else DictVal = 0#
End Function

' Ключи словаря по убыванию значения. limit <= 0 - все.
Public Sub TopKeys(ByVal d As Object, ByVal limit As Long, _
                    ByRef labs As Variant, ByRef vals As Variant)
    Dim n As Long
    n = d.Count
    If n = 0 Then
        labs = Array(): vals = Array()
        Exit Sub
    End If
    Dim kk() As Double, vv() As String, i As Long, k As Variant
    ReDim kk(0 To n - 1)
    ReDim vv(0 To n - 1)
    i = 0
    For Each k In d.Keys
        kk(i) = -CDbl(d(k))
        vv(i) = CStr(k)
        i = i + 1
    Next k
    QSortPair kk, vv, 0, n - 1

    Dim m As Long
    m = n
    If limit > 0 And limit < n Then m = limit
    Dim rl() As Variant, rv() As Variant
    ReDim rl(0 To m - 1)
    ReDim rv(0 To m - 1)
    For i = 0 To m - 1
        rl(i) = vv(i)
        rv(i) = -kk(i)
    Next i
    labs = rl
    vals = rv
End Sub

' Отчётная неделя части «Техника»: последняя неделя снимка, а если в ней меньше
' половины объёма предыдущей - предыдущая. Неполная неделя ломает сравнение.
Public Function ZoneReportWeek() As Long
    If mRepWeek > 0 Then ZoneReportWeek = mRepWeek: Exit Function
    EnsureZn
    Dim byW As Object
    Set byW = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In mZn.Keys
        If CLng(mZn(k)(Z_WEEK)) > 0 Then AddCnt byW, CStr(mZn(k)(Z_WEEK)), 1#
    Next k
    If byW.Count = 0 Then ZoneReportWeek = 0: Exit Function

    Dim mx As Long
    mx = 0
    For Each k In byW.Keys
        If CLng(k) > mx Then mx = CLng(k)
    Next k
    Dim prev As Long
    prev = PrevWeek(mx)
    If byW.Exists(CStr(prev)) Then
        If DictVal(byW, CStr(mx)) < DictVal(byW, CStr(prev)) * 0.5 Then mx = prev
    End If
    mRepWeek = mx
    ZoneReportWeek = mx
End Function

' Предыдущая ISO-неделя. Через дату, а не вычитанием единицы: на границе года
' арифметика по ключу даёт «202600».
Public Function PrevWeek(ByVal yw As Long) As Long
    Dim d As Date
    d = WeekMonday(yw) - 7
    PrevWeek = IsoYearWeek(CDbl(d))
End Function

Private Function WeekMonday(ByVal yw As Long) As Date
    Dim y As Long, w As Long, jan4 As Date, mon1 As Date
    y = yw \ 100: w = yw Mod 100
    jan4 = DateSerial(y, 1, 4)
    mon1 = jan4 - ((Weekday(jan4, vbMonday) - 1))
    WeekMonday = mon1 + (w - 1) * 7
End Function

' Окно из n недель, заканчивающееся отчётной.
Public Function WeekWindow(ByVal n As Long) As Variant
    Dim rw As Long
    rw = ZoneReportWeek()
    Dim a() As Variant, i As Long, cur As Long
    ReDim a(0 To n - 1)
    cur = rw
    For i = n - 1 To 0 Step -1
        a(i) = cur
        cur = PrevWeek(cur)
    Next i
    WeekWindow = a
End Function
Public Function FmtF(ByVal v As Double, ByVal digits As Long) As String
    Dim s As String
    s = Format$(v, "0" & IIf(digits > 0, "." & String$(digits, "0"), ""))
    FmtF = Replace$(s, ".", ",")
End Function

' Число в SVG-координате: точка как разделитель, иначе браузер не поймёт.
Public Function Cx(ByVal v As Double) As String
    Cx = Replace$(Format$(v, "0.0"), ",", ".")
End Function

Public Function Nb() As String
    Nb = ChrW$(&HA0)
End Function

Public Function Dash() As String
    Dash = ChrW$(&H2014)
End Function

Public Function Pc(ByVal v As Double, ByVal digits As Long) As String
    Pc = FmtF(v, digits) & Nb() & "%"
End Function

Public Function Rub(ByVal v As Double) As String
    If Abs(v) >= 1000000# Then
        Rub = FmtF(v / 1000000#, 2) & Nb() & "млн" & Nb() & ChrW$(&H20BD)
    ElseIf Abs(v) >= 1000# Then
        Rub = modContentMTO.FmtInt(v / 1000#) & Nb() & "тыс" & Nb() & ChrW$(&H20BD)
    Else
        Rub = modContentMTO.FmtInt(v) & Nb() & ChrW$(&H20BD)
    End If
End Function

' Часы в человеческом виде: минуты / часы / сутки.
Public Function Hh(ByVal v As Double) As String
    If v < 1# Then
        Hh = modContentMTO.FmtInt(v * 60#) & " мин"
    ElseIf v < 48# Then
        Hh = FmtF(v, 1) & " ч"
    Else
        Hh = FmtF(v / 24#, 1) & " сут"
    End If
End Function

' Ячейка процента: рамка, а не заливка. Шкала 0/50/75/100 - см. контракт шаблона.
Public Function PctTd(ByVal p As Double, Optional ByVal hasValue As Boolean = True) As String
    If Not hasValue Then
        PctTd = "<td class=""pct empty""><span>" & Dash() & "</span></td>"
        Exit Function
    End If
    PctTd = "<td class=""pct""><span style=""border-color:" & _
        modContentMTO.PctBorderColor(p) & """>" & FmtF(p, 0) & "</span></td>"
End Function

Public Function Pill(ByVal cls As String, ByVal txt As String) As String
    Pill = "<span class=""pill " & cls & """>" & modContentMTO.Esc(txt) & "</span>"
End Function

Public Function Truthy(ByVal v As Variant) As Boolean
    Truthy = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then Truthy = CBool(v): Exit Function
    If IsNumeric(v) Then Truthy = (CDbl(v) <> 0#): Exit Function
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    Truthy = (s = "TRUE" Or s = "ИСТИНА" Or s = "1" Or s = "ДА")
End Function

' ISO-неделя: год*100 + номер. Год берётся по ISO, а не календарный, иначе
' 31 декабря уезжает в первую неделю следующего года со старым годом в ключе.
Public Function IsoYearWeek(ByVal ser As Double) As Long
    IsoYearWeek = 0
    If ser <= 0# Then Exit Function
    Dim d As Date, w As Long, y As Long
    d = CDate(Int(ser))
    w = DatePart("ww", d, vbMonday, vbFirstFourDays)
    y = Year(d)
    If w >= 52 And Month(d) = 1 Then
        y = y - 1
    ElseIf w = 1 And Month(d) = 12 Then
        y = y + 1
    End If
    IsoYearWeek = y * 100 + w
End Function

Public Function YearMonth(ByVal ser As Double) As Long
    YearMonth = 0
    If ser <= 0# Then Exit Function
    Dim d As Date
    d = CDate(Int(ser))
    YearMonth = Year(d) * 100 + Month(d)
End Function

' Минимальная НЕнулевая из двух отметок: подписи одной дирекции может не быть.
Private Function MinPos(ByVal a As Double, ByVal b As Double) As Double
    If a <= 0# Then MinPos = b: Exit Function
    If b <= 0# Then MinPos = a: Exit Function
    If a < b Then MinPos = a Else MinPos = b
End Function

Private Function ZAcc(ByVal z As Variant) As Double
    ZAcc = MinPos(CDbl(z(Z_ACCG)), CDbl(z(Z_ACCD)))
End Function

Private Function ZLev(ByVal z As Variant) As Double
    ZLev = MinPos(CDbl(z(Z_LEVG)), CDbl(z(Z_LEVD)))
End Function

Private Function DMon(ByVal ser As Double) As String
    If ser <= 0# Then DMon = Dash() Else DMon = Format$(CDate(ser), "dd.mm")
End Function

' Ярлык недели для оси: «2026-13». Год обязателен - иначе недели разных лет
' схлопываются в одну ось и сравнение с прошлым годом невозможно.
Public Function WLab(ByVal yw As Long) As String
    WLab = CStr(yw \ 100) & "-" & Format$(yw Mod 100, "00")
End Function

Public Function MLab(ByVal ym As Long) As String
    Dim m As Variant
    m = Array("янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек")
    MLab = CStr(m((ym Mod 100) - 1)) & " " & CStr(ym \ 100)
End Function

' =====================================================================================
' Примитивы графики. Инлайн-SVG, без xmlns, цвета - только переменными темы.
' =====================================================================================
' Горизонтальные полосы: подпись, полоса, число. Подпись обрезается по ширине
' колонки - иначе длинные названия статусов налезают на полосу.
Public Function HBars(ByVal labels As Variant, ByVal values As Variant, _
                       ByVal w As Long, ByVal labw As Long, ByVal rowh As Long, _
                       Optional ByVal moneyFmt As Boolean = False) As String
    Dim i As Long, mx As Double
    mx = 0#
    For i = LBound(values) To UBound(values)
        If CDbl(values(i)) > mx Then mx = CDbl(values(i))
    Next i
    If mx <= 0# Then mx = 1#

    Dim cnt As Long, hgt As Long
    cnt = UBound(values) - LBound(values) + 1
    hgt = rowh * cnt + 6

    Dim maxch As Long
    maxch = Int((labw - 10) / 6.4)
    If maxch < 8 Then maxch = 8

    Dim s As String
    s = "<svg viewBox=""0 0 " & CStr(w) & " " & CStr(hgt) & """ role=""img"" " & _
        "aria-label=""распределение"">"
    For i = LBound(values) To UBound(values)
        Dim y As Double, bw As Double
        y = 8# + (i - LBound(values)) * rowh
        bw = (w - labw - 90) * CDbl(values(i)) / mx

        Dim lab As String, muted As Boolean
        lab = CStr(labels(i))
        muted = (Left$(lab, 1) = "(")
        If Len(lab) > maxch Then lab = Left$(lab, maxch - 1) & ChrW$(&H2026)

        Dim fill As String, ink As String
        If muted Then
            fill = "var(--muted)": ink = "var(--muted)"
        Else
            fill = "var(--s1)": ink = "var(--ink)"
        End If

        s = s & "<text x=""0"" y=""" & Cx(y + 12.1) & """ font-size=""12.5"" " & _
            "font-family=""IBM Plex Sans,sans-serif"" fill=""" & ink & """>" & _
            modContentMTO.Esc(lab) & "</text>"
        s = s & "<rect x=""" & CStr(labw) & """ y=""" & Cx(y) & """ width=""" & Cx(bw) & _
            """ height=""14"" fill=""" & fill & """/>"
        s = s & "<text x=""" & Cx(labw + bw + 8) & """ y=""" & Cx(y + 12.1) & """ font-size=""12"" " & _
            "font-family=""IBM Plex Mono,monospace"" fill=""var(--ink-2)"">" & _
            IIf(moneyFmt, Rub(CDbl(values(i))), modContentMTO.FmtInt(CDbl(values(i)))) & "</text>"
    Next i
    HBars = s & "</svg>"
End Function

' Вертикальные столбики с подписью значения сверху и категории снизу.
Public Function Cols(ByVal labels As Variant, ByVal values As Variant, _
                      ByVal w As Long, ByVal h As Long) As String
    Dim i As Long, mx As Double
    mx = 0#
    For i = LBound(values) To UBound(values)
        If CDbl(values(i)) > mx Then mx = CDbl(values(i))
    Next i
    If mx <= 0# Then mx = 1#

    Dim cnt As Long, baseY As Double, topY As Double, stepX As Double, bw As Double
    cnt = UBound(values) - LBound(values) + 1
    baseY = h - 40: topY = 24
    stepX = w / cnt
    bw = stepX * 0.58

    Dim s As String
    s = "<svg viewBox=""0 0 " & CStr(w) & " " & CStr(h) & """ role=""img"">"
    s = s & "<line x1=""0"" y1=""" & Cx(baseY) & """ x2=""" & CStr(w) & """ y2=""" & Cx(baseY) & _
        """ stroke=""var(--line-strong)""/>"
    For i = LBound(values) To UBound(values)
        Dim x As Double, bh As Double
        x = stepX * (i - LBound(values)) + (stepX - bw) / 2#
        bh = (baseY - topY) * CDbl(values(i)) / mx
        s = s & "<rect x=""" & Cx(x) & """ y=""" & Cx(baseY - bh) & """ width=""" & Cx(bw) & _
            """ height=""" & Cx(bh) & """ fill=""var(--s1)"" opacity=""0.85""/>"
        s = s & "<text x=""" & Cx(x + bw / 2#) & """ y=""" & Cx(baseY - bh - 6#) & _
            """ text-anchor=""middle"" font-size=""12"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--ink-2)"">" & modContentMTO.FmtInt(CDbl(values(i))) & "</text>"
        s = s & "<text x=""" & Cx(x + bw / 2#) & """ y=""" & Cx(baseY + 17#) & _
            """ text-anchor=""middle"" font-size=""11.5"" font-family=""IBM Plex Sans,sans-serif"" " & _
            "fill=""var(--muted)"">" & modContentMTO.Esc(CStr(labels(i))) & "</text>"
    Next i
    Cols = s & "</svg>"
End Function

' Столбики + линия по правой шкале. Используется для кривой старения парка.
' hasLine(i) = False -> точка пропускается (в когорте нет машин).
Private Function LFmt(ByVal v As Double, ByVal digits As Long, ByVal pctFmt As Boolean) As String
    If pctFmt Then LFmt = Pc(v, digits) Else LFmt = FmtF(v, digits)
End Function

Public Function Collines(ByVal labels As Variant, ByVal bars As Variant, _
                          ByVal ln As Variant, ByVal w As Long, ByVal h As Long, _
                          ByVal digits As Long, _
                          Optional ByVal pctFmt As Boolean = False) As String
    Dim i As Long, mxb As Double, lo As Double, hi As Double
    mxb = 0#
    For i = LBound(bars) To UBound(bars)
        If CDbl(bars(i)) > mxb Then mxb = CDbl(bars(i))
    Next i
    If mxb <= 0# Then mxb = 1#
    lo = CDbl(ln(LBound(ln))): hi = lo
    For i = LBound(ln) To UBound(ln)
        If CDbl(ln(i)) < lo Then lo = CDbl(ln(i))
        If CDbl(ln(i)) > hi Then hi = CDbl(ln(i))
    Next i
    Dim rng As Double
    rng = hi - lo
    If rng = 0# Then rng = 1#

    Dim baseY As Double, topY As Double, stepX As Double, bw As Double
    Dim cnt As Long
    cnt = UBound(bars) - LBound(bars) + 1
    baseY = h - 44: topY = 30
    stepX = w / cnt
    bw = stepX * 0.5

    Dim s As String, pts As String, cxs() As Double, cys() As Double
    ReDim cxs(0 To cnt - 1)
    ReDim cys(0 To cnt - 1)
    s = "<svg viewBox=""0 0 " & CStr(w) & " " & CStr(h) & """ role=""img"" aria-label="""">"
    s = s & "<line x1=""0"" y1=""" & CStr(CLng(baseY)) & """ x2=""" & CStr(w) & """ y2=""" & _
        CStr(CLng(baseY)) & """ stroke=""var(--line-strong)""/>"
    For i = 0 To cnt - 1
        Dim x As Double, bh As Double
        x = stepX * i + (stepX - bw) / 2#
        bh = (baseY - topY) * CDbl(bars(LBound(bars) + i)) / mxb
        s = s & "<rect x=""" & Cx(x) & """ y=""" & Cx(baseY - bh) & """ width=""" & Cx(bw) & _
            """ height=""" & Cx(bh) & """ fill=""var(--s1)"" opacity=""0.45""/>"
        cxs(i) = x + bw / 2#
        cys(i) = baseY - 14# - (baseY - topY - 28#) * (CDbl(ln(LBound(ln) + i)) - lo) / rng
        s = s & "<text x=""" & Cx(cxs(i)) & """ y=""" & Cx(baseY + 17#) & _
            """ text-anchor=""middle"" font-size=""11.5"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--muted)"">" & modContentMTO.Esc(CStr(labels(LBound(labels) + i))) & "</text>"
        If pts <> "" Then pts = pts & " "
        pts = pts & Cx(cxs(i)) & "," & Cx(cys(i))
    Next i
    s = s & "<polyline points=""" & pts & """ fill=""none"" stroke=""var(--s2)"" " & _
        "stroke-width=""2.4"" stroke-linejoin=""round""/>"
    For i = 0 To cnt - 1
        s = s & "<circle cx=""" & Cx(cxs(i)) & """ cy=""" & Cx(cys(i)) & """ r=""3.2"" fill=""var(--s2)""/>"
    Next i
    s = s & "<text x=""" & Cx(cxs(cnt - 1) - 6#) & """ y=""" & Cx(cys(cnt - 1) - 10#) & _
        """ text-anchor=""end"" font-size=""13"" font-family=""IBM Plex Mono,monospace"" " & _
        "fill=""var(--s2)"">" & LFmt(CDbl(ln(UBound(ln))), digits, pctFmt) & "</text>"
    s = s & "<text x=""0"" y=""" & Cx(cys(0) - 8#) & """ font-size=""11"" " & _
        "font-family=""IBM Plex Mono,monospace"" fill=""var(--muted)"">" & _
        LFmt(CDbl(ln(LBound(ln))), digits, pctFmt) & "</text>"
    Collines = s & "</svg>"
End Function

' Возвраты по периодам: столбики - знаменатель, линия - доля возвратов.
' Периоды, у которых окно в 30 суток ещё не истекло, приглушены и вынесены
' за пунктирную границу: без пометки падение на правом краю читается как рост
' качества, хотя это артефакт окна.
Public Function RetChart(ByVal labels As Variant, ByVal dens As Variant, _
                          ByVal pcts As Variant, ByVal nClosed As Long) As String
    Const W As Long = 940
    Const H As Long = 250
    Dim baseY As Double, topY As Double
    baseY = 204#: topY = 30#

    Dim cnt As Long, i As Long
    cnt = UBound(dens) - LBound(dens) + 1
    If cnt = 0 Then RetChart = modContentMTO.EmptyNote(): Exit Function

    Dim mxb As Double, lo As Double, hi As Double
    mxb = 0#
    For i = 0 To cnt - 1
        If CDbl(dens(LBound(dens) + i)) > mxb Then mxb = CDbl(dens(LBound(dens) + i))
    Next i
    If mxb <= 0# Then mxb = 1#
    lo = CDbl(pcts(LBound(pcts))): hi = lo
    For i = 0 To cnt - 1
        Dim pv As Double
        pv = CDbl(pcts(LBound(pcts) + i))
        If pv < lo Then lo = pv
        If pv > hi Then hi = pv
    Next i
    Dim rng As Double
    rng = hi - lo
    If rng = 0# Then rng = 1#

    Dim stepX As Double, bw As Double
    stepX = W / cnt
    bw = stepX * 0.5

    Dim cxs() As Double, cys() As Double
    ReDim cxs(0 To cnt - 1)
    ReDim cys(0 To cnt - 1)

    Dim s As String
    s = "<svg viewBox=""0 0 " & CStr(W) & " " & CStr(H) & """ role=""img"">"
    s = s & "<line x1=""0"" y1=""" & CStr(CLng(baseY)) & """ x2=""" & CStr(W) & """ y2=""" & _
        CStr(CLng(baseY)) & """ stroke=""var(--line-strong)""/>"
    For i = 0 To cnt - 1
        Dim x As Double, bh As Double, op As String
        x = stepX * i + (stepX - bw) / 2#
        bh = (baseY - topY) * CDbl(dens(LBound(dens) + i)) / mxb
        op = IIf(i < nClosed, "0.45", "0.18")
        s = s & "<rect x=""" & Cx(x) & """ y=""" & Cx(baseY - bh) & """ width=""" & Cx(bw) & _
            """ height=""" & Cx(bh) & """ fill=""var(--s1)"" opacity=""" & op & """/>"
        cxs(i) = x + bw / 2#
        cys(i) = baseY - 14# - (baseY - topY - 28#) * (CDbl(pcts(LBound(pcts) + i)) - lo) / rng
        s = s & "<text x=""" & Cx(cxs(i)) & """ y=""" & Cx(baseY + 17#) & _
            """ text-anchor=""middle"" font-size=""11"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--muted)"">" & modContentMTO.Esc(CStr(labels(LBound(labels) + i))) & "</text>"
    Next i

    Dim solid As String, dashed As String
    solid = "": dashed = ""
    For i = 0 To cnt - 1
        If i < nClosed Then
            If solid <> "" Then solid = solid & " "
            solid = solid & Cx(cxs(i)) & "," & Cx(cys(i))
        End If
        If i >= nClosed - 1 And nClosed >= 1 Then
            If dashed <> "" Then dashed = dashed & " "
            dashed = dashed & Cx(cxs(i)) & "," & Cx(cys(i))
        End If
    Next i
    If InStr(solid, " ") > 0 Then
        s = s & "<polyline points=""" & solid & """ fill=""none"" stroke=""var(--s2)"" " & _
            "stroke-width=""2.4"" stroke-linejoin=""round""/>"
    End If
    If InStr(dashed, " ") > 0 Then
        s = s & "<polyline points=""" & dashed & """ fill=""none"" stroke=""var(--s2)"" " & _
            "stroke-width=""2"" stroke-dasharray=""5 4"" opacity=""0.55"" stroke-linejoin=""round""/>"
    End If
    For i = 0 To cnt - 1
        Dim po As String
        po = IIf(i < nClosed, "1", "0.5")
        s = s & "<circle cx=""" & Cx(cxs(i)) & """ cy=""" & Cx(cys(i)) & """ r=""3.2"" " & _
            "fill=""var(--s2)"" opacity=""" & po & """/>"
        s = s & "<text x=""" & Cx(cxs(i)) & """ y=""" & Cx(cys(i) - 9#) & _
            """ text-anchor=""middle"" font-size=""11"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--s2)"" opacity=""" & po & """>" & _
            FmtF(CDbl(pcts(LBound(pcts) + i)), 1) & "</text>"
    Next i
    If nClosed > 0 And nClosed < cnt Then
        Dim dx As Double
        dx = stepX * nClosed
        s = s & "<line x1=""" & Cx(dx) & """ y1=""24"" x2=""" & Cx(dx) & """ y2=""" & _
            CStr(CLng(baseY)) & """ stroke=""var(--line-strong)"" stroke-dasharray=""3 3""/>"
        s = s & "<text x=""" & Cx(dx + 6#) & """ y=""32"" font-size=""10.5"" " & _
            "font-family=""IBM Plex Mono,monospace"" fill=""var(--muted)"">окно не закрыто</text>"
    End If
    RetChart = s & "</svg>"
End Function

' Три фазы наряда накопительно по неделям. Средняя полоса неразложима:
' отделить очередь от ремонта можно только по истории смены поста, её нет.
Public Function PhasesChart(ByVal weeks As Variant, ByVal a1 As Variant, _
                             ByVal a2 As Variant, ByVal a3 As Variant) As String
    Const W As Long = 940
    Const H As Long = 250
    Dim baseY As Double, topY As Double
    baseY = H - 44: topY = 24

    Dim cnt As Long, i As Long, mx As Double
    cnt = UBound(weeks) - LBound(weeks) + 1
    mx = 0#
    For i = 0 To cnt - 1
        Dim tt As Double
        tt = CDbl(a1(LBound(a1) + i)) + CDbl(a2(LBound(a2) + i)) + CDbl(a3(LBound(a3) + i))
        If tt > mx Then mx = tt
    Next i
    If mx <= 0# Then mx = 1#

    Dim stepX As Double, bw As Double
    stepX = W / cnt
    bw = stepX * 0.5

    Dim s As String
    s = "<svg viewBox=""0 0 " & CStr(W) & " " & CStr(H) & """ role=""img"">"
    s = s & "<line x1=""0"" y1=""" & CStr(CLng(baseY)) & """ x2=""" & CStr(W) & """ y2=""" & _
        CStr(CLng(baseY)) & """ stroke=""var(--line-strong)""/>"
    For i = 0 To cnt - 1
        Dim x As Double, y As Double, j As Long
        x = stepX * i + (stepX - bw) / 2#
        y = baseY
        Dim vals As Variant, colr As Variant
        vals = Array(CDbl(a1(LBound(a1) + i)), CDbl(a2(LBound(a2) + i)), CDbl(a3(LBound(a3) + i)))
        colr = Array("var(--s4)", "var(--s2)", "var(--s7)")
        For j = 0 To 2
            Dim bh As Double
            bh = (baseY - topY) * CDbl(vals(j)) / mx
            s = s & "<rect x=""" & Cx(x) & """ y=""" & Cx(y - bh) & """ width=""" & Cx(bw) & _
                """ height=""" & Cx(bh) & """ fill=""" & CStr(colr(j)) & """/>"
            y = y - bh
        Next j
        s = s & "<text x=""" & Cx(x + bw / 2#) & """ y=""" & Cx(y - 6#) & _
            """ text-anchor=""middle"" font-size=""11.5"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--ink-2)"">" & _
            FmtF(CDbl(vals(0)) + CDbl(vals(1)) + CDbl(vals(2)), 1) & "</text>"
        s = s & "<text x=""" & Cx(x + bw / 2#) & """ y=""" & Cx(baseY + 17#) & _
            """ text-anchor=""middle"" font-size=""11.5"" font-family=""IBM Plex Mono,monospace"" " & _
            "fill=""var(--muted)"">" & Format$(CLng(weeks(LBound(weeks) + i)) Mod 100, "00") & "</text>"
    Next i
    s = s & "</svg>"
    s = s & "<div class=""legend""><span><i style=""background:var(--s4)""></i>" & _
        "постановка: создание " & ChrW$(&H2192) & " приёмка</span>"
    s = s & "<span><i style=""background:var(--s2)""></i>ремзона: очередь + ремонт</span>"
    s = s & "<span><i style=""background:var(--s7)""></i>закрытие: выбытие " & _
        ChrW$(&H2192) & " zn_closed</span></div>"
    PhasesChart = s
End Function

' =====================================================================================
' Классификатор описаний дефекта: характер работы (уровень 0) и узел (уровень 1).
' Правила узла привязаны к группе дефекта - свободных категорий нет.
' Остаток «(не классифицировано)» показывается числом, а не прячется.
' Источник правил: tools/mockup_v1.0/classify.py, версия от 09.09.2026.
' =====================================================================================
Private Sub AddKind(ByRef c As Collection, ByVal nm As String, ByVal keys As String)
    c.Add Array(nm, Split(keys, "|"))
End Sub

Private Sub AddNode(ByVal grp As String, ByVal nm As String, ByVal keys As String)
    If Not mNodeRules.Exists(grp) Then mNodeRules.Add grp, New Collection
    mNodeRules(grp).Add Array(nm, Split(keys, "|"))
End Sub

Private Sub BuildRules()
    If Not (mNodeRules Is Nothing) Then Exit Sub
    Set mNodeRules = CreateObject("Scripting.Dictionary")

    ' Порядок важен: первое совпадение выигрывает. Оформление и обслуживание идут
    ' раньше отказа, иначе «требуется замена» утащит плановую работу в отказы.
    Dim c As Collection
    Set c = New Collection
    AddKind c, "Оформление документов", _
        "акт тех|акт технич|акт техсост|составить акт|создать акт|списан|формирования акта"
    AddKind c, "Обслуживание", _
        "слив конденсат|хлорирован|хлорк|долив|долит|подкачк|протяжк|цеп|подготовк|" & _
        "заправ|смазк|промывк|очистк|мойк|регулировк|обслуживан|уровень масла|" & _
        "уровень антифриз|масла мин|сезонн|затяжк|то км|то проведен"
    AddKind c, "Дефектовка и осмотр", _
        "дефектовк|осмотр|проверк|диагностик|выявлен|тех состояни"
    AddKind c, "Отказ", _
        "не работа|не гор|не заводит|не запускает|нет запуска|нет хода|пропал ход|" & _
        "пропадает ход|глохнет|течь|утечк|обрыв|порвал|неисправн|ошибк|авари|сел |" & _
        "разряж|не выходит|не включает|не поднима|не опуска|не открыва|не закрыва|" & _
        "не включа|не фиксир|не едет|не льет|не убирает|не устанавлива|нет |перегорел|" & _
        "заклинил|треснул|сломан|разбит|порыв|люфт|стук|вибрац|дым|перегрев|не тормоз|" & _
        "не крутит|пропал|отсутств|повреж|спущен|спустил|слетел|отвалил|не выдал|не держит"
    AddKind c, "Изготовление и замена узла", _
        "изготовлен|замена|заменит|установит|монтаж|сварочн|ремонт|кронштейн"

    Set mKindRules = c

    Const G1 As String = "Ходовая часть / Управление / Тормозная система"
    AddNode G1, "Пневмосистема", "конденсат|воздух|пневмо|осушител|ресивер"
    AddNode G1, "Тормоза", "тормоз|колодк|суппорт|растормаж"
    AddNode G1, "Колёса, шины и цепи", _
        "колес|шин|покрышк|диск|ниппел|камер|цеп|спущен|спустил|подкачк|протяжк"
    AddNode G1, "Ход и трансмиссия", _
        "ход|кардан|мост|редуктор|полуос|кпп|коробк|сцеплен|раздатк|дифференциал|шасси"
    AddNode G1, "Подвеска", "рессор|амортизат|пружин|подвеск|сайлент|рычаг"
    AddNode G1, "Рулевое управление", "рулев|гур|гидроусилител|рейк|наконечник|тяг"
    AddNode G1, "Гидравлика", "рвд|гидравл|гидроцилиндр|шланг"
    AddNode G1, "Отмечено вне своей группы", "заводит|запуск|двс|масл|крен"

    Const G2 As String = "Электрооборудование"
    AddNode G2, "Освещение и сигнализация", _
        "фара|габарит|света|свет|лампа|маяк|проблеск|поворотник|подсветк|сигнал|плафон|фар"
    AddNode G2, "Аккумулятор и зарядка", "аккумулятор|акб|сел |разряж|зарядк|генератор"
    AddNode G2, "Пусковая система", "стартер|запуск|заводит|зажиган"
    AddNode G2, "Проводка и предохранители", _
        "проводк|провод|предохранител|клемм|разъем|замыкан|жгут"
    AddNode G2, "Приборы и электроника", _
        "датчик|приборн|блок управлен|дисплей|контроллер|тахограф|ошибк|реле"
    AddNode G2, "Стеклоочистители", "дворник|стеклоочистит|омыват"
    AddNode G2, "Отопитель и климат", "отопител|печк|кондиционер|подогрев"
    AddNode G2, "Прочая электрика", "счетчик|камер|плафон|напряжен|кабел|электр"

    Const G3 As String = "ДВС"
    AddNode G3, "Пуск и работа двигателя", _
        "запуск|заводит|запускает|глохнет|оборот|не тянет|мощност|троит|двигател|двс"
    AddNode G3, "Смазка", "масл"
    AddNode G3, "Система охлаждения", _
        "антифриз|охлажд|радиатор|помп|термостат|перегрев|тосол|патрубок"
    AddNode G3, "Топливная система", "топлив|форсунк|тнвд|солярк|дизел|бак"
    AddNode G3, "Ремённый привод", "ремен|ремн|ролик|натяжител"
    AddNode G3, "Впуск и выпуск", "турбин|выхлоп|глушител|сажев|интеркулер"

    Const G4 As String = "Спецоборудование"
    AddNode G4, "Санитарное оборудование", _
        "хлорирован|хлорк|вод|ассениз|туалет|смешиван"
    AddNode G4, "Гидравлика", _
        "гидравл|гидроцилиндр|рвд|гидрораспределит|гж|маслостанц|гидромотор"
    AddNode G4, "Подъёмный механизм", _
        "стрел|подъем|вышк|платформ|аутригер|телескоп|лебедк|люльк"
    AddNode G4, "Транспортёр и лента", "лент|транспортер|конвейер"
    AddNode G4, "Насосы и агрегаты", _
        "насос|компрессор|вентилятор|горелк|котел|бойлер|теплообменник|установк|" & _
        "щетк|подогреват"
    AddNode G4, "Ёмкости и арматура", "емкост|бак|клапан|вентил|фитинг|кран|заслонк"
    AddNode G4, "Оборудование в целом", _
        "спецоборудован|спец оборудован|оборудован|озп|влп|форсунк|шланг|кабел|ремен|оборот"

    Const G5 As String = "Кондиционер / отопитель"
    AddNode G5, "Отопитель", "отопител|печк|тэн|горелк|подогрев|вебасто|планар"
    AddNode G5, "Кондиционер", "кондиционер|фреон|испарител|охлажд"
    AddNode G5, "Вентиляция", "вентилятор|обдув|заслонк|воздуховод|воздух"

    Const G6 As String = "Целостность и внешний вид ТС и СТ/СНО"
    AddNode G6, "Кузов и облицовка", _
        "кузов|облицов|крыл|бампер|капот|панел|обшивк|коррози|покрас|вмятин|" & _
        "подножк|кронштейн|коврик|сварочн"
    AddNode G6, "Двери и люки", "двер|люк|замок|петл"
    AddNode G6, "Остекление и зеркала", "стекл|зеркал|лобов|форточк|дворник|стеклоочистит"
    AddNode G6, "Тент и уплотнения", "тент|уплотнител|резинк|пыльник"
    AddNode G6, "Колёса и шины", "колес|спущен|спустил|шин"
    AddNode G6, "Документы и оформление", "акт|списан|дтп|падени|дефектовк"

    Const G7 As String = "Рабочее место / Органы управления спецоборудованием"
    AddNode G7, "Органы управления", _
        "джойстик|пульт|рычаг|кнопк|тумблер|педал|вал|передач|раздатк|фиксатор|" & _
        "капюшон|марш|бордачк"
    AddNode G7, "Кресло и кабина", "кресл|сиден|кабин|ремень безопасн|двер"
    AddNode G7, "Приборная панель", "приборн|индикатор|дисплей|спидометр|ошибк"
    AddNode G7, "Стеклоочистители", "дворник|стеклоочистит|омыват"
    AddNode G7, "Спецагрегаты рабочего места", _
        "установк|насос|аутригер|люльк|печк|теплообменник|заслонк|течь|колокол"
    AddNode G7, "Документы и оформление", "акт|списан|дтп|событи"

    Const G8 As String = "Сцепное устройство"
    AddNode G8, "Сцепное устройство", "сцепн|фаркоп|дышл|петл|палец|крюк"
End Sub

' Нормализация описания: нижний регистр, ё -> е, цифры и знаки -> пробел.
Private Function NormDesc(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, code As Long
    Dim out As String, prevSp As Boolean
    t = LCase$(Replace$(CStr(s), ChrW$(&H451), ChrW$(&H435)))
    t = Replace$(t, ChrW$(&H401), ChrW$(&H435))
    out = "": prevSp = True
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        code = AscW(ch)
        Dim keep As Boolean
        keep = False
        If code >= &H430 And code <= &H44F Then keep = True
        If code >= 97 And code <= 122 Then keep = True
        If keep Then
            out = out & ch
            prevSp = False
        Else
            If Not prevSp Then
                out = out & " "
                prevSp = True
            End If
        End If
    Next i
    NormDesc = Trim$(out)
End Function

Private Function MatchRules(ByVal rules As Collection, ByVal t As String) As String
    MatchRules = ""
    If rules Is Nothing Then Exit Function
    Dim it As Variant, keys As Variant, j As Long
    For Each it In rules
        keys = it(1)
        For j = LBound(keys) To UBound(keys)
            If InStr(1, t, CStr(keys(j)), vbBinaryCompare) > 0 Then
                MatchRules = CStr(it(0))
                Exit Function
            End If
        Next j
    Next it
End Function

Private Sub EnsureCls()
    If mClsDone Then Exit Sub
    EnsureZn
    BuildRules

    Dim k As Variant, z As Variant, t As String
    For Each k In mZn.Keys
        z = mZn(k)
        t = NormDesc(CStr(z(Z_DESC)))
        If Len(t) = 0 Then
            z(Z_KIND) = EMPTYD
            z(Z_NODE) = EMPTYD
        Else
            Dim kd As String, nd As String
            kd = MatchRules(mKindRules, t)
            If kd = "" Then kd = UNCLS
            nd = ""
            If mNodeRules.Exists(CStr(z(Z_DEFEKT))) Then
                nd = MatchRules(mNodeRules(CStr(z(Z_DEFEKT))), t)
            End If
            If nd = "" Then nd = UNCLS
            z(Z_KIND) = kd
            z(Z_NODE) = nd
        End If
        mZn(k) = z
    Next k
    mClsDone = True
End Sub

' =====================================================================================
' Общие обёртки разметки
' =====================================================================================
Public Function KpiTile(ByVal lab As String, ByVal val As String, _
                         ByVal cls As String, ByVal note As String) As String
    Dim c As String
    If cls <> "" Then c = " " & cls Else c = ""
    KpiTile = "<div class=""kpi""><div class=""lab"">" & lab & "</div>" & _
        "<div class=""val num" & c & """>" & val & "</div>" & _
        "<div class=""row""><span class=""delta flat"">" & note & "</span></div></div>"
End Function

Public Function NoteBlk(ByVal t As String) As String
    NoteBlk = "<div class=""calc-note"">" & t & "</div>"
End Function

Public Function MockLabel(ByVal t As String) As String
    MockLabel = "<div class=""mock-label"">" & modContentMTO.Esc(t) & "</div>"
End Function

' Когорта возраста: 0 = до 5 лет ... 4 = 20+; -1 = год выпуска неизвестен.
Public Function CohortOf(ByVal ageY As Double) As Long
    CohortOf = -1
    If ageY < 0# Then Exit Function
    If ageY < 5# Then
        CohortOf = 0
    ElseIf ageY < 10# Then
        CohortOf = 1
    ElseIf ageY < 15# Then
        CohortOf = 2
    ElseIf ageY < 20# Then
        CohortOf = 3
    Else
        CohortOf = 4
    End If
End Function

Public Function CohortLabels() As Variant
    CohortLabels = Array("до 5 лет", "5" & ChrW$(&H2013) & "10", _
        "10" & ChrW$(&H2013) & "15", "15" & ChrW$(&H2013) & "20", "20 +")
End Function

' =====================================================================================
' СЛАЙД 5. Парк и заезды
' =====================================================================================
Public Function BuildKpiFleet() As String
    EnsureVeh
    Dim rw As Long, pw As Long
    rw = ZoneReportWeek()
    pw = PrevWeek(rw)

    Dim visWeek As Double, visPrev As Double
    visWeek = DictVal(mVisitWeeks, CStr(rw))
    visPrev = DictVal(mVisitWeeks, CStr(pw))

    Dim col As Collection, i As Long, multi As Long, tot As Long
    Set col = mVisitSizes(CStr(GAP_HOURS))
    tot = col.Count
    multi = 0
    For i = 1 To col.Count
        If CLng(col(i)) > 1 Then multi = multi + 1
    Next i

    Dim ages As New Collection, k As Variant, old15 As Long, a As Double
    old15 = 0
    For Each k In mVeh.Keys
        a = AgeYears(CDbl(mVeh(k)(V_MADE)))
        If a >= 0# Then
            ages.Add a
            If a >= 15# Then old15 = old15 + 1
        End If
    Next k
    Dim hasMed As Boolean, medAge As Double
    medAge = MedianOf(ages, hasMed)

    Dim s As String
    s = "<div class=""kpis"">"
    s = s & KpiTile("Машин в парке", modContentMTO.FmtInt(CDbl(mVeh.Count)), "", _
        "уникальных гаражных номеров")
    s = s & KpiTile("Заездов за неделю", modContentMTO.FmtInt(visWeek), "", _
        "было " & modContentMTO.FmtInt(visPrev) & " неделей раньше")
    s = s & KpiTile("Заездов пакетом", Pc(SafePct(multi, tot), 1), "", _
        "больше одного наряда за заезд")
    s = s & KpiTile("Возраст парка", IIf(hasMed, FmtF(medAge, 1), Dash()) & _
        " <small>лет</small>", "", _
        modContentMTO.FmtInt(CDbl(old15)) & " машин старше 15 лет")
    BuildKpiFleet = s & "</div>"
End Function

Public Function SafePct(ByVal a As Double, ByVal b As Double) As Double
    If b = 0# Then SafePct = 0# Else SafePct = 100# * a / b
End Function

Public Function BuildPostsWeek() As String
    EnsureZn
    Dim rw As Long
    rw = ZoneReportWeek()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant, p As String
    For Each k In mZn.Keys
        z = mZn(k)
        If CLng(z(Z_WEEK)) = rw Then
            p = Trim$(CStr(z(Z_POST)))
            If p = "" Then p = "(пост не указан)"
            AddCnt d, p, 1#
        End If
    Next k
    If d.Count = 0 Then BuildPostsWeek = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant
    TopKeys d, 8, labs, vals
    BuildPostsWeek = HBars(labs, vals, 680, 250, 24) & NoteBlk( _
        "«Пост не указан» стоит отдельной строкой, а не растворён в прочих: " & _
        "пока строка видна руководителю, её чинят.")
End Function

Public Function BuildAgeCurve() As String
    EnsureVeh
    ' Кривая идёт по ГОДАМ выпуска, а не по пятилетним когортам: на когортах
    ' перелом размазывается. Годы, где машин меньше трёх, отброшены -
    ' там среднее не значит ничего.
    Dim cars As Object, unp As Object
    Set cars = CreateObject("Scripting.Dictionary")
    Set unp = CreateObject("Scripting.Dictionary")

    Dim k As Variant, yv As Long
    For Each k In mVeh.Keys
        If CDbl(mVeh(k)(V_MADE)) > 0# Then
            yv = Year(CDate(CDbl(mVeh(k)(V_MADE))))
            AddCnt cars, CStr(yv), 1#
        End If
    Next k
    Dim z As Variant
    For Each k In mZn.Keys
        z = mZn(k)
        If Len(CStr(z(Z_VEH))) > 0 And Not IsPlanned(CStr(z(Z_TYPE))) Then
            If CDbl(z(Z_MADE)) > 0# Then
                yv = Year(CDate(CDbl(z(Z_MADE))))
                AddCnt unp, CStr(yv), 1#
            End If
        End If
    Next k

    Dim yrs() As Double, tmp() As String, n As Long, i As Long
    n = cars.Count
    If n = 0 Then BuildAgeCurve = modContentMTO.EmptyNote(): Exit Function
    ReDim yrs(0 To n - 1)
    ReDim tmp(0 To n - 1)
    i = 0
    For Each k In cars.Keys
        yrs(i) = CDbl(k): tmp(i) = CStr(k): i = i + 1
    Next i
    QSortPair yrs, tmp, 0, n - 1

    Dim labs() As Variant, bars() As Variant, ln() As Variant, m As Long
    ReDim labs(0 To n - 1)
    ReDim bars(0 To n - 1)
    ReDim ln(0 To n - 1)
    m = 0
    For i = 0 To n - 1
        Dim c As Double
        c = DictVal(cars, tmp(i))
        If c >= 3# Then
            labs(m) = tmp(i)
            bars(m) = c
            ln(m) = DictVal(unp, tmp(i)) / c
            m = m + 1
        End If
    Next i
    If m < 2 Then BuildAgeCurve = modContentMTO.EmptyNote(): Exit Function
    ReDim Preserve labs(0 To m - 1)
    ReDim Preserve bars(0 To m - 1)
    ReDim Preserve ln(0 To m - 1)

    Dim wi As Long, wv As Double
    wi = 0: wv = CDbl(ln(0))
    For i = 1 To m - 1
        If CDbl(ln(i)) > wv Then wv = CDbl(ln(i)): wi = i
    Next i
    Dim newV As Double
    newV = CDbl(ln(m - 1))

    Dim s As String
    s = "<figure>" & Collines(labs, bars, ln, 940, 250, 1)
    s = s & "<figcaption>Столбики " & ChrW$(&H2014) & _
        " сколько машин этого года выпуска в парке, линия " & ChrW$(&H2014) & _
        " <b>внеплановых нарядов на одну машину</b>. Годы, где машин меньше трёх, " & _
        "не показаны: там среднее не значит ничего.</figcaption></figure>"
    s = s & "<div class=""legend""><span><i style=""background:var(--s1);opacity:.45""></i>" & _
        "машин в парке</span><span><i style=""background:var(--s2)""></i>" & _
        "внеплановых нарядов на машину</span></div>"

    Dim ratio As String
    If newV > 0# Then ratio = FmtF(wv / newV, 1) Else ratio = Dash()
    s = s & NoteBlk("Перелом виден: " & CStr(labs(m - 1)) & " года выпуска " & _
        ChrW$(&H2014) & " " & FmtF(newV, 1) & " внеплановых наряда на машину, " & _
        CStr(labs(wi)) & " " & ChrW$(&H2014) & " уже " & FmtF(wv, 1) & _
        ", разница в " & ratio & " раза. Кривая идёт по годам, а не по пятилетним " & _
        "когортам: на когортах этот перелом размазывается и его не видно. Плато после " & _
        "десяти лет " & ChrW$(&H2014) & " не признак надёжности, а следствие того, что " & _
        "самые проблемные машины к этому возрасту уже списаны.")
    BuildAgeCurve = s
End Function

' Первый сегмент названия группы дефекта: в шапке матрицы полное название
' не помещается, а «/» внутри «СТ/СНО» резать нельзя - режем по « / ».
Public Function ShortGrp(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, " / ", vbBinaryCompare)
    If p > 0 Then ShortGrp = Left$(s, p - 1) Else ShortGrp = s
End Function

Public Function BuildAgeMatrix() As String
    EnsureVeh
    Dim labs As Variant
    labs = CohortLabels()

    ' Пять самых частых групп дефекта по всему снимку - это колонки матрицы.
    Dim gAll As Object
    Set gAll = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) Then
            If Len(Trim$(CStr(z(Z_DEFEKT)))) > 0 Then AddCnt gAll, CStr(z(Z_DEFEKT)), 1#
        End If
    Next k
    Dim gl As Variant, gv As Variant
    TopKeys gAll, 5, gl, gv
    If UBound(gl) < 0 Then BuildAgeMatrix = modContentMTO.EmptyNote(): Exit Function

    Dim cars(0 To 4) As Double, unp(0 To 4) As Double, pln(0 To 4) As Double
    Dim mat() As Double, i As Long, j As Long
    ReDim mat(0 To 4, 0 To UBound(gl))

    For Each k In mVeh.Keys
        i = CohortOf(AgeYears(CDbl(mVeh(k)(V_MADE))))
        If i >= 0 Then cars(i) = cars(i) + 1#
    Next k

    For Each k In mZn.Keys
        z = mZn(k)
        If Len(CStr(z(Z_VEH))) > 0 Then
            i = CohortOf(AgeYears(CDbl(z(Z_MADE))))
            If i >= 0 Then
                If IsPlanned(CStr(z(Z_TYPE))) Then
                    pln(i) = pln(i) + 1#
                Else
                    unp(i) = unp(i) + 1#
                    For j = 0 To UBound(gl)
                        If CStr(z(Z_DEFEKT)) = CStr(gl(j)) Then
                            mat(i, j) = mat(i, j) + 1#
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
    Next k

    Dim s As String
    s = "<table class=""matrix""><thead><tr><th rowspan=""2"">Когорта</th>" & _
        "<th class=""n"" rowspan=""2"">Машин</th>" & _
        "<th class=""n"" rowspan=""2"">Внеплановых<br>на машину</th>" & _
        "<th class=""n"" rowspan=""2"">Плановых<br>на машину</th>" & _
        "<th class=""grp sep-l"" colspan=""" & CStr(UBound(gl) + 1) & """>" & _
        "Что именно ломается, % внеплановых нарядов когорты</th>" & _
        "<th class=""n"" rowspan=""2"">Внеплановых<br>всего</th></tr><tr>"
    For j = 0 To UBound(gl)
        s = s & "<th class=""n"">" & modContentMTO.Esc(ShortGrp(CStr(gl(j)))) & "</th>"
    Next j
    s = s & "</tr></thead><tbody>"

    For i = 0 To 4
        If cars(i) > 0# Then
            s = s & "<tr><td class=""head"">" & modContentMTO.Esc(CStr(labs(i))) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(cars(i)) & "</td>"
            s = s & "<td class=""n"">" & FmtF(unp(i) / cars(i), 1) & "</td>"
            s = s & "<td class=""n"">" & FmtF(pln(i) / cars(i), 1) & "</td>"
            For j = 0 To UBound(gl)
                s = s & PctTd(SafePct(mat(i, j), unp(i)), unp(i) > 0#)
            Next j
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(unp(i)) & "</td></tr>"
        End If
    Next i
    s = s & "</tbody></table>"

    Dim g0 As String
    g0 = ShortGrp(CStr(gl(0)))
    s = s & NoteBlk("Строка читается слева направо: сколько машин, сколько на каждую " & _
        "приходится внеплановых и плановых нарядов, и из чего эти внеплановые состоят. " & _
        "<b>" & modContentMTO.Esc(g0) & "</b> " & ChrW$(&H2014) & " первая по объёму " & _
        "группа отказов; смотреть на неё надо по строкам, а не по одной когорте. " & _
        "<b>Ограничение:</b> <code>model_type</code> " & ChrW$(&H2014) & " константа по " & _
        "построению выгрузки, поэтому в одной когорте лежат погрузчик и легковая; " & _
        "разрез по видам техники возможен через <code>vehicle_group</code>, но там " & _
        "~100 значений и нужен укрупняющий справочник.")
    BuildAgeMatrix = s
End Function

Public Function BuildAging() As String
    EnsureVeh
    Dim labs As Variant
    labs = CohortLabels()
    Dim cars(0 To 4) As Double, vis(0 To 4) As Double
    Dim prt(0 To 4) As Double, hrs(0 To 4) As Double
    Dim k As Variant, i As Long
    For Each k In mVeh.Keys
        i = CohortOf(AgeYears(CDbl(mVeh(k)(V_MADE))))
        If i >= 0 Then
            cars(i) = cars(i) + 1#
            vis(i) = vis(i) + CDbl(mVeh(k)(V_VISITS))
            prt(i) = prt(i) + CDbl(mVeh(k)(V_PARTS))
            hrs(i) = hrs(i) + CDbl(mVeh(k)(V_HOURS))
        End If
    Next k

    Dim s As String
    s = "<table><thead><tr><th>Когорта</th><th class=""n"">Машин</th>" & _
        "<th class=""n"">Заездов / машину</th><th class=""n"">Материалы / машину</th>" & _
        "<th class=""n"">Часов в ремзоне / машину</th></tr></thead><tbody>"
    Dim first As Long, last As Long
    first = -1: last = -1
    For i = 0 To 4
        If cars(i) > 0# Then
            If first < 0 Then first = i
            last = i
            s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(cars(i)) & "</td>"
            s = s & "<td class=""n"">" & FmtF(vis(i) / cars(i), 1) & "</td>"
            s = s & "<td class=""n"">" & Rub(prt(i) / cars(i)) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(hrs(i) / cars(i)) & "</td></tr>"
        End If
    Next i
    s = s & "</tbody></table>"
    If first >= 0 Then
        s = s & NoteBlk("Расход на материалы по когортам: " & _
            Rub(prt(first) / cars(first)) & " на машину «" & _
            modContentMTO.Esc(CStr(labs(first))) & "» против " & _
            Rub(prt(last) / cars(last)) & " у группы «" & _
            modContentMTO.Esc(CStr(labs(last))) & "». Отдельные годы дают выбросы " & _
            "(одна машина с дорогим ремонтом на десять в когорте сдвигает среднее), " & _
            "поэтому для защиты бюджета брать когорты, а не годы.")
    End If
    BuildAging = s
End Function

Public Function BuildPack() As String
    EnsureVeh
    Dim col As Collection, i As Long
    Set col = mVisitSizes(CStr(GAP_HOURS))

    Dim dist(0 To 3) As Double, b As Long
    For i = 1 To col.Count
        b = CLng(col(i))
        If b > 4 Then b = 4
        dist(b - 1) = dist(b - 1) + 1#
    Next i

    Dim dl As Variant, dv() As Variant
    dl = Array("1", "2", "3", "4 +")
    ReDim dv(0 To 3)
    For i = 0 To 3
        dv(i) = dist(i)
    Next i

    Dim one As Double
    one = SafePct(dist(0), CDbl(col.Count))

    Dim s As String
    s = "<div class=""two-col wide-l""><div>" & Cols(dl, dv, 420, 200)
    s = s & NoteBlk("Нарядов на заезд. Один наряд на заезд " & ChrW$(&H2014) & " " & _
        Pc(one, 1) & " случаев.") & "</div><div>"
    s = s & MockLabel("Чувствительность к порогу склейки")
    s = s & "<table><thead><tr><th>Разрыв</th><th class=""n"">Заездов</th>" & _
        "<th class=""n"">Пакетных</th></tr></thead><tbody>"

    Dim gaps As Variant, gi As Long, lo As Double, hi As Double
    gaps = Array(8#, 12#, 24#)
    lo = -1#: hi = -1#
    For gi = 0 To 2
        Dim c2 As Collection, m2 As Long
        Set c2 = mVisitSizes(CStr(gaps(gi)))
        m2 = 0
        For i = 1 To c2.Count
            If CLng(c2(i)) > 1 Then m2 = m2 + 1
        Next i
        Dim mp As Double
        mp = SafePct(CDbl(m2), CDbl(c2.Count))
        If gi = 0 Then lo = mp
        If gi = 2 Then hi = mp
        s = s & "<tr><td>" & CStr(CLng(gaps(gi))) & " ч</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(c2.Count)) & "</td><td class=""n"">" & _
            Pc(mp, 1) & "</td></tr>"
    Next gi
    s = s & "</tbody></table></div></div>"
    s = s & NoteBlk("<b>Порог " & CStr(CLng(GAP_HOURS)) & " ч " & ChrW$(&H2014) & _
        " гипотеза, а не измеренный факт.</b> Признака фактического выезда с " & _
        "территории в выгрузке нет. При 8 и 24 часах доля пакетных заездов меняется " & _
        "с " & Pc(lo, 1) & " до " & Pc(hi, 1) & " " & ChrW$(&H2014) & _
        " разница качественная, поэтому число публикуется только вместе с порогом.")
    BuildPack = s
End Function

' =====================================================================================
' Возвраты
'
' ВОЗВРАТ: повторный заход машины в ремзону по той же группе дефекта в течение
' RET_WINDOW суток ПОСЛЕ ЗАКРЫТИЯ предыдущего наряда.
'   1. Отсчёт от zn_closed, а не от создания: иначе долгий ремонт засчитывает сам себя.
'   2. Возврат относится к периоду ПЕРВОГО наряда - проверяется тот ремонт.
'   3. Только ближайший возврат в окне: иначе хроника утраивает долю.
'   4. Только внеплановые: у планового наряда дефекта нет.
' Уровни строгости: по группе, по подкатегории, по подкатегории и только отказы.
' =====================================================================================
Private mRetReady As Boolean
Private mRetNumM As Object, mRetNumW As Object      ' по группе
Private mRetDenM As Object, mRetDenW As Object
Private mFailNumM As Object, mFailNumW As Object    ' по подкатегории, только отказы
Private mFailDenM As Object, mFailDenW As Object
Private mRetTot As Long, mStrictTot As Long, mFailTot As Long
Private mDenAll As Long, mDenFail As Long
Private mGapMed As Double, mGapHas As Boolean
Private mNodeNum As Object, mNodeDen As Object      ' «группа|узел» -> счётчик

' mode: 0 - ключ = группа; 1 - группа+узел; 2 - группа+узел и только отказы.
Private Function RetKey(ByVal z As Variant, ByVal mode As Long) As String
    RetKey = ""
    If IsPlanned(CStr(z(Z_TYPE))) Then Exit Function
    If Len(CStr(z(Z_VEH))) = 0 Or CDbl(z(Z_DATE)) <= 0# Then Exit Function
    If mode >= 2 Then
        If CStr(z(Z_KIND)) <> "Отказ" Then Exit Function
    End If
    Dim g As String
    g = CStr(z(Z_DEFEKT))
    If mode = 0 Then
        If Len(g) = 0 Then Exit Function
        RetKey = g
    Else
        If CStr(z(Z_NODE)) = UNCLS Or CStr(z(Z_NODE)) = EMPTYD Then Exit Function
        RetKey = g & "|" & CStr(z(Z_NODE))
    End If
End Function

' Пары «наряд -> возврат» для заданного уровня строгости.
' gaps (если передана) наполняется интервалами в сутках, numM/numW - счётчиками
' по периоду ПЕРВОГО наряда, nodes - счётчиком по «группа|узел».
Private Function RetPairs(ByVal mode As Long, ByVal numM As Object, ByVal numW As Object, _
                          ByVal gaps As Collection, ByVal nodes As Object) As Long
    EnsureCls
    Dim byK As Object
    Set byK = CreateObject("Scripting.Dictionary")

    Dim k As Variant, z As Variant, kk As String
    For Each k In mZn.Keys
        z = mZn(k)
        kk = RetKey(z, mode)
        If Len(kk) > 0 Then
            kk = CStr(z(Z_VEH)) & Chr$(1) & kk
            If Not byK.Exists(kk) Then byK.Add kk, New Collection
            byK(kk).Add CStr(k)
        End If
    Next k

    Dim total As Long
    total = 0
    For Each k In byK.Keys
        Dim col As Collection
        Set col = byK(k)
        If col.Count > 1 Then
            Dim ds() As Double, ns() As String, i As Long, j As Long
            ReDim ds(0 To col.Count - 1)
            ReDim ns(0 To col.Count - 1)
            For i = 1 To col.Count
                ns(i - 1) = CStr(col(i))
                ds(i - 1) = CDbl(mZn(ns(i - 1))(Z_DATE))
            Next i
            QSortPair ds, ns, 0, UBound(ds)

            For i = 0 To UBound(ds) - 1
                Dim za As Variant, base As Double
                za = mZn(ns(i))
                base = CDbl(za(Z_CLOSED))
                If base <= 0# Then base = CDbl(za(Z_DATE))
                For j = i + 1 To UBound(ds)
                    Dim gp As Double
                    gp = Int(ds(j) - base)
                    If gp >= 0# Then
                        If gp > RET_WINDOW Then Exit For
                        total = total + 1
                        If Not numM Is Nothing Then AddCnt numM, CStr(za(Z_MONTH)), 1#
                        If Not numW Is Nothing Then AddCnt numW, CStr(za(Z_WEEK)), 1#
                        If Not gaps Is Nothing Then gaps.Add gp
                        If Not nodes Is Nothing Then
                            AddCnt nodes, CStr(za(Z_DEFEKT)) & "|" & CStr(za(Z_NODE)), 1#
                        End If
                        Exit For
                    End If
                Next j
            Next i
        End If
    Next k
    RetPairs = total
End Function

Private Sub EnsureRet()
    If mRetReady Then Exit Sub
    EnsureCls

    Set mRetNumM = CreateObject("Scripting.Dictionary")
    Set mRetNumW = CreateObject("Scripting.Dictionary")
    Set mRetDenM = CreateObject("Scripting.Dictionary")
    Set mRetDenW = CreateObject("Scripting.Dictionary")
    Set mFailNumM = CreateObject("Scripting.Dictionary")
    Set mFailNumW = CreateObject("Scripting.Dictionary")
    Set mFailDenM = CreateObject("Scripting.Dictionary")
    Set mFailDenW = CreateObject("Scripting.Dictionary")
    Set mNodeNum = CreateObject("Scripting.Dictionary")
    Set mNodeDen = CreateObject("Scripting.Dictionary")

    Dim k As Variant, z As Variant
    mDenAll = 0: mDenFail = 0
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) And CDbl(z(Z_DATE)) > 0# Then
            mDenAll = mDenAll + 1
            AddCnt mRetDenM, CStr(z(Z_MONTH)), 1#
            AddCnt mRetDenW, CStr(z(Z_WEEK)), 1#
            If CStr(z(Z_KIND)) = "Отказ" Then
                mDenFail = mDenFail + 1
                AddCnt mFailDenM, CStr(z(Z_MONTH)), 1#
                AddCnt mFailDenW, CStr(z(Z_WEEK)), 1#
                AddCnt mNodeDen, CStr(z(Z_DEFEKT)) & "|" & CStr(z(Z_NODE)), 1#
            End If
        End If
    Next k

    Dim gaps As Collection
    Set gaps = New Collection
    mRetTot = RetPairs(0, mRetNumM, mRetNumW, gaps, Nothing)
    mStrictTot = RetPairs(1, Nothing, Nothing, Nothing, Nothing)
    mFailTot = RetPairs(2, mFailNumM, mFailNumW, Nothing, mNodeNum)
    mGapMed = MedianOf(gaps, mGapHas)

    mRetReady = True
    modLog.WriteDebug 2, "Техника", "modContentZone.EnsureRet", _
        "Внеплановых " & CStr(mDenAll) & ", возвратов по группе " & CStr(mRetTot) & _
        ", по подкатегории " & CStr(mStrictTot) & ", по отказу " & CStr(mFailTot)
End Sub

' Граница «окно не закрыто»: у нарядов последних RET_WINDOW суток возврат
' физически не успел случиться, и доля занижена.
Private Function OpenEdgeMonth() As Long
    OpenEdgeMonth = YearMonth(SnapshotEnd() - RET_WINDOW)
End Function

Private Function OpenEdgeWeek() As Long
    OpenEdgeWeek = IsoYearWeek(SnapshotEnd() - RET_WINDOW)
End Function

Public Function BuildRetKpi() As String
    EnsureRet
    Dim s As String
    s = "<div class=""kpis"">"
    s = s & KpiTile("Возвратов по отказу", Pc(SafePct(mFailTot, mDenFail), 1), "crit", _
        modContentMTO.FmtInt(CDbl(mFailTot)) & " из " & _
        modContentMTO.FmtInt(CDbl(mDenFail)) & " отказов")
    s = s & KpiTile("По подкатегории", Pc(SafePct(mStrictTot, mDenAll), 1), "", _
        "без разделения на отказ и обслуживание")
    s = s & KpiTile("По группе дефекта", Pc(SafePct(mRetTot, mDenAll), 1), "", _
        "верхняя граница, для сравнения")
    s = s & KpiTile("Медиана интервала", _
        IIf(mGapHas, modContentMTO.FmtInt(mGapMed), Dash()) & " <small>сут</small>", "", _
        "окно " & CStr(RET_WINDOW) & " суток")
    BuildRetKpi = s & "</div>"
End Function

Public Function BuildRetMonth() As String
    EnsureRet
    Dim edge As Long, snapY As Long, i As Long
    edge = OpenEdgeMonth()
    snapY = Year(CDate(SnapshotEnd()))

    Dim labs() As Variant, dens() As Variant, pcts() As Variant
    ReDim labs(0 To 11)
    ReDim dens(0 To 11)
    ReDim pcts(0 To 11)
    Dim m As Long, nClosed As Long, ym As Long
    m = 0: nClosed = 0
    For i = 1 To 12
        ym = snapY * 100 + i
        If DictVal(mFailDenM, CStr(ym)) > 0# Then
            labs(m) = MLab(ym)
            dens(m) = DictVal(mFailDenM, CStr(ym))
            pcts(m) = SafePct(DictVal(mFailNumM, CStr(ym)), DictVal(mFailDenM, CStr(ym)))
            If ym < edge Then nClosed = nClosed + 1
            m = m + 1
        End If
    Next i
    If m = 0 Then BuildRetMonth = modContentMTO.EmptyNote(): Exit Function
    ReDim Preserve labs(0 To m - 1)
    ReDim Preserve dens(0 To m - 1)
    ReDim Preserve pcts(0 To m - 1)

    Dim s As String
    s = RetChart(labs, dens, pcts, nClosed)
    s = s & "<div class=""legend""><span><i style=""background:var(--s1);opacity:.45""></i>" & _
        "отказов за месяц</span><span><i style=""background:var(--s2)""></i>" & _
        "доля возвратов, %</span></div>"
    s = s & NoteBlk("Снимок обрезан " & Format$(CDate(SnapshotEnd()), "dd.mm.yyyy") & _
        ". У нарядов последних " & CStr(RET_WINDOW) & " суток окно возврата ещё не " & _
        "истекло " & ChrW$(&H2014) & " возврат физически не успел случиться, и доля " & _
        "занижена. Эти периоды приглушены и в тренд не входят: без пометки падение на " & _
        "правом краю читается как рост качества, хотя это артефакт окна.")
    BuildRetMonth = s
End Function

Public Function BuildRetWeek() As String
    EnsureRet
    Dim wk As Variant
    wk = WeekWindow(8)
    Dim edge As Long, i As Long, nClosed As Long
    edge = OpenEdgeWeek()

    Dim labs() As Variant, dens() As Variant, pcts() As Variant
    ReDim labs(0 To UBound(wk))
    ReDim dens(0 To UBound(wk))
    ReDim pcts(0 To UBound(wk))
    nClosed = 0
    For i = 0 To UBound(wk)
        labs(i) = WLab(CLng(wk(i)))
        dens(i) = DictVal(mFailDenW, CStr(wk(i)))
        pcts(i) = SafePct(DictVal(mFailNumW, CStr(wk(i))), DictVal(mFailDenW, CStr(wk(i))))
        If CLng(wk(i)) < edge Then nClosed = nClosed + 1
    Next i

    Dim s As String
    s = RetChart(labs, dens, pcts, nClosed)
    s = s & "<div class=""legend""><span><i style=""background:var(--s1);opacity:.45""></i>" & _
        "отказов за неделю</span><span><i style=""background:var(--s2)""></i>" & _
        "доля возвратов, %</span></div>"
    s = s & NoteBlk("Ось подписана как <b>год-неделя</b>, а не одним номером: иначе " & _
        "недели разных лет схлопываются в одну ось и сравнение с прошлым годом " & _
        "становится невозможным.")
    BuildRetWeek = s
End Function

Public Function BuildRetNode() As String
    EnsureRet
    Dim labs As Variant, vals As Variant
    TopKeys mNodeNum, 0, labs, vals
    If UBound(labs) < 0 Then BuildRetNode = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<table><thead><tr><th>Узел</th><th>Группа дефекта</th><th class=""n"">Отказов</th>" & _
        "<th class=""n"">Возвратов</th><th class=""n"">%</th></tr></thead><tbody>"
    Dim i As Long, shown As Long
    shown = 0
    For i = 0 To UBound(labs)
        Dim den As Double
        den = DictVal(mNodeDen, CStr(labs(i)))
        If den >= 40# Then
            Dim parts As Variant
            parts = Split(CStr(labs(i)), "|")
            s = s & "<tr><td class=""head"">" & modContentMTO.Esc(CStr(parts(1))) & "</td>"
            s = s & "<td style=""color:var(--ink-2)"">" & _
                modContentMTO.Esc(Left$(CStr(parts(0)), 34)) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(den) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(CDbl(vals(i))) & "</td>"
            s = s & PctTd(SafePct(CDbl(vals(i)), den)) & "</tr>"
            shown = shown + 1
            If shown >= 10 Then Exit For
        End If
    Next i
    s = s & "</tbody></table>"
    s = s & NoteBlk("Узел " & ChrW$(&H2014) & " вторая ступень классификации, она " & _
        "строится по описанию дефекта. Именно она превращает возврат из совпадения в " & _
        "повтор: возврат по той же <b>группе</b> даёт " & _
        Pc(SafePct(mRetTot, mDenAll), 1) & " и ничего не значит, по той же " & _
        "<b>подкатегории и только по отказам</b> " & ChrW$(&H2014) & " " & _
        Pc(SafePct(mFailTot, mDenFail), 1) & ".")
    BuildRetNode = s
End Function

' =====================================================================================
' СЛАЙД 6. Что ломается и что возвращается
' =====================================================================================
Public Function BuildChronics() As String
    EnsureVeh
    Dim names() As String, n As Long, i As Long, k As Variant
    n = 0
    For Each k In mVeh.Keys
        If CDbl(mVeh(k)(V_VISITS)) > 0# Then n = n + 1
    Next k
    If n = 0 Then BuildChronics = modContentMTO.EmptyNote(): Exit Function
    ReDim names(0 To n - 1)
    i = 0
    For Each k In mVeh.Keys
        If CDbl(mVeh(k)(V_VISITS)) > 0# Then names(i) = CStr(k): i = i + 1
    Next i

    Dim rVis As Object, rHrs As Object, rPrt As Object
    Set rVis = RankBy(names, V_VISITS)
    Set rHrs = RankBy(names, V_HOURS)
    Set rPrt = RankBy(names, V_PARTS)

    Dim best() As Double, nm() As String
    ReDim best(0 To n - 1)
    ReDim nm(0 To n - 1)
    For i = 0 To n - 1
        Dim a As Double, b As Double, c As Double
        a = DictVal(rVis, names(i)): b = DictVal(rHrs, names(i)): c = DictVal(rPrt, names(i))
        best(i) = a
        If b < best(i) Then best(i) = b
        If c < best(i) Then best(i) = c
        nm(i) = names(i)
    Next i
    QSortPair best, nm, 0, n - 1

    Dim top As Long
    top = 12
    If top > n Then top = n

    Dim totHours As Double, topHours As Double
    totHours = 0#
    For Each k In mVeh.Keys
        totHours = totHours + CDbl(mVeh(k)(V_HOURS))
    Next k

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Гар. " & ChrW$(&H2116) & "</th>" & _
        "<th>Группа техники</th><th class=""n"">Заездов</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">Часов в ремзоне</th><th class=""n"">Материалы</th>" & _
        "<th class=""n"">Возраст</th><th class=""n"">Ранг</th><th>Флаг</th>" & _
        "</tr></thead><tbody>"

    Dim firstName As String, firstPrt As Double
    firstName = "": firstPrt = 0#
    For i = 0 To top - 1
        Dim v As Variant, rv As Double, rh As Double, rp As Double
        v = mVeh(nm(i))
        rv = DictVal(rVis, nm(i)): rh = DictVal(rHrs, nm(i)): rp = DictVal(rPrt, nm(i))
        topHours = topHours + CDbl(v(V_HOURS))
        If i = 0 Then firstName = nm(i): firstPrt = rp

        Dim flg As String
        If rv <= rh And rv <= rp Then
            flg = Pill("crit", "по заездам")
        ElseIf rh <= rp Then
            flg = Pill("ser", "по времени")
        Else
            flg = Pill("warn", "по деньгам")
        End If

        Dim ag As Double
        ag = AgeYears(CDbl(v(V_MADE)))

        s = s & "<tr><td class=""n mono"" style=""text-align:left;font-size:14px"">" & _
            modContentMTO.Esc(nm(i)) & "</td>"
        s = s & "<td style=""color:var(--ink-2)"">" & modContentMTO.Esc(CStr(v(V_GROUP))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(CDbl(v(V_VISITS))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(CDbl(v(V_ZN))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(CDbl(v(V_HOURS))) & "</td>"
        s = s & "<td class=""n"">" & Rub(CDbl(v(V_PARTS))) & "</td>"
        s = s & "<td class=""n"">" & IIf(ag >= 0#, FmtF(ag, 1), Dash()) & "</td>"
        s = s & "<td class=""n mono"">" & modContentMTO.FmtInt(rv) & " / " & _
            modContentMTO.FmtInt(rh) & " / " & modContentMTO.FmtInt(rp) & "</td>"
        s = s & "<td>" & flg & "</td></tr>"
    Next i
    s = s & "</tbody></table></div>"

    s = s & NoteBlk("Три ранга в одной колонке: по заездам / по часам / по деньгам. " & _
        "Видно, что машина-рекордсмен по заездам и машина-рекордсмен по деньгам " & _
        ChrW$(&H2014) & " <b>разные машины</b>: " & modContentMTO.Esc(firstName) & _
        " первая по заездам и " & modContentMTO.FmtInt(firstPrt) & "-я по деньгам. " & _
        "Единый «индекс проблемности» этот сюжет скрывает. Колонка называется «часов " & _
        "в ремзоне», а не «время ремонта»: в интервале сидят очередь и ожидание " & _
        "запчастей. " & modContentMTO.FmtInt(CDbl(top)) & " машин из " & _
        modContentMTO.FmtInt(CDbl(mVeh.Count)) & " дают " & _
        Pc(SafePct(topHours, totHours), 1) & " всех часов в ремзоне. Часы суммируются " & _
        "по нарядам: при двух одновременно открытых нарядах на одной машине время " & _
        "считается дважды " & ChrW$(&H2014) & " это загрузка ремзоны, а не календарь машины.")
    BuildChronics = s
End Function

' Ранги 1..N по убыванию поля записи машины.
Private Function RankBy(ByRef names() As String, ByVal fld As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim n As Long, i As Long
    n = UBound(names) - LBound(names) + 1
    Dim kk() As Double, vv() As String
    ReDim kk(0 To n - 1)
    ReDim vv(0 To n - 1)
    For i = 0 To n - 1
        kk(i) = -CDbl(mVeh(names(LBound(names) + i))(fld))
        vv(i) = names(LBound(names) + i)
    Next i
    QSortPair kk, vv, 0, n - 1
    For i = 0 To n - 1
        d(vv(i)) = CDbl(i + 1)
    Next i
    Set RankBy = d
End Function

Public Function BuildPareto() As String
    EnsureZn
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant, g As String, tot As Double
    tot = 0#
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) Then
            g = Trim$(CStr(z(Z_DEFEKT)))
            If g = "" Then g = NOSECT
            AddCnt d, g, 1#
            tot = tot + 1#
        End If
    Next k
    If tot = 0# Then BuildPareto = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant
    TopKeys d, 8, labs, vals

    Dim s As String, i As Long, acc As Double, top3 As Double
    s = "<div class=""two-col wide-l""><div>" & HBars(labs, vals, 520, 230, 24) & "</div><div>"
    s = s & "<table><thead><tr><th>Группа дефекта</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">%</th><th class=""n"">Накоп.</th></tr></thead><tbody>"
    acc = 0#: top3 = 0#
    For i = 0 To UBound(labs)
        acc = acc + CDbl(vals(i))
        If i < 3 Then top3 = top3 + CDbl(vals(i))
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(vals(i))) & "</td><td class=""n"">" & _
            FmtF(SafePct(CDbl(vals(i)), tot), 1) & "</td><td class=""n"">" & _
            FmtF(SafePct(acc, tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table></div></div>"
    s = s & NoteBlk("Три верхние группы дают <b>" & Pc(SafePct(top3, tot), 1) & _
        "</b> внепланового потока. Плановое ТО из расчёта исключено: у планового " & _
        "наряда дефекта нет, и пустой <code>defekt_type</code> у него " & _
        ChrW$(&H2014) & " норма, а не брак. <code>defekt_type</code> " & _
        ChrW$(&H2014) & " это уже готовая группа отказа (двигатель, тормозная " & _
        "система, электрооборудование), отдельный справочник групп не нужен. Глубже " & _
        "группы Парето не идёт: детализация вынесена в блок ниже.")
    BuildPareto = s
End Function

Public Function BuildRepeats() As String
    EnsureZn
    Dim byK As Object
    Set byK = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant, kk As String
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) Then
            If Len(CStr(z(Z_VEH))) > 0 And CDbl(z(Z_DATE)) > 0# _
               And Len(Trim$(CStr(z(Z_DEFEKT)))) > 0 Then
                kk = CStr(z(Z_VEH)) & Chr$(1) & CStr(z(Z_DEFEKT))
                If Not byK.Exists(kk) Then byK.Add kk, New Collection
                byK(kk).Add CStr(k)
            End If
        End If
    Next k

    Dim gapsA() As Double, rowsA() As String, cnt As Long, cap As Long
    cap = 4096
    ReDim gapsA(0 To cap - 1)
    ReDim rowsA(0 To cap - 1)
    cnt = 0

    Dim unpTot As Double
    unpTot = 0#
    For Each k In mZn.Keys
        If Not IsPlanned(CStr(mZn(k)(Z_TYPE))) Then unpTot = unpTot + 1#
    Next k

    Dim nPairs As Long
    nPairs = 0
    For Each k In byK.Keys
        Dim col As Collection
        Set col = byK(k)
        If col.Count > 1 Then
            Dim ds() As Double, ns() As String, i As Long
            ReDim ds(0 To col.Count - 1)
            ReDim ns(0 To col.Count - 1)
            For i = 1 To col.Count
                ns(i - 1) = CStr(col(i))
                ds(i - 1) = CDbl(mZn(ns(i - 1))(Z_DATE))
            Next i
            QSortPair ds, ns, 0, UBound(ds)
            For i = 0 To UBound(ds) - 1
                Dim gp As Double
                gp = Int(ds(i + 1) - ds(i))
                If gp > 0# And gp <= 30# Then
                    nPairs = nPairs + 1
                    If cnt < cap Then
                        gapsA(cnt) = gp
                        rowsA(cnt) = ns(i) & Chr$(1) & ns(i + 1)
                        cnt = cnt + 1
                    End If
                End If
            Next i
        End If
    Next k

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Машина</th><th>Группа дефекта</th>" & _
        "<th class=""n"">Интервал</th><th>Первый / повторный</th>" & _
        "<th>Описание повторного</th><th>Итог разбора</th></tr></thead><tbody>"
    If cnt > 0 Then
        ReDim Preserve gapsA(0 To cnt - 1)
        ReDim Preserve rowsA(0 To cnt - 1)
        QSortPair gapsA, rowsA, 0, cnt - 1
        Dim shown As Long
        shown = cnt
        If shown > 6 Then shown = 6
        For i = 0 To shown - 1
            Dim pr As Variant, za As Variant, zb As Variant
            pr = Split(rowsA(i), Chr$(1))
            za = mZn(CStr(pr(0))): zb = mZn(CStr(pr(1)))
            s = s & "<tr><td class=""mono"">" & modContentMTO.Esc(CStr(za(Z_VEH))) & "</td>"
            s = s & "<td>" & modContentMTO.Esc(CStr(za(Z_DEFEKT))) & "</td>"
            s = s & "<td class=""n"">" & modContentMTO.FmtInt(gapsA(i)) & " сут</td>"
            s = s & "<td class=""mono"">" & ChrW$(&H2026) & Right$(CStr(pr(0)), 4) & " / " & _
                ChrW$(&H2026) & Right$(CStr(pr(1)), 4) & "</td>"
            s = s & "<td style=""color:var(--ink-2)"">" & _
                modContentMTO.Esc(Left$(CStr(zb(Z_DESC)), 70)) & "</td>"
            s = s & "<td>" & Dash() & "</td></tr>"
        Next i
    End If
    s = s & "</tbody></table></div>"
    s = s & NoteBlk("<b>Дисклеймер обязателен:</b> кандидаты на повтор, а не " & _
        "показатель качества ремонта. Формально под правило попадает " & _
        modContentMTO.FmtInt(CDbl(nPairs)) & " пар " & ChrW$(&H2014) & " " & _
        Pc(SafePct(CDbl(nPairs), unpTot), 1) & " внеплановых нарядов. Группа дефекта в " & _
        "выгрузке есть, справочника поверх неё заводить не нужно, но <b>группа " & _
        "крупная</b>: два разных отказа внутри «электрооборудования» правило считает " & _
        "повтором. Поэтому это верхняя граница кандидатов, а не доля брака. Отделить " & _
        "настоящий повтор можно только по <code>defect_desc</code> и глазами мастера " & _
        ChrW$(&H2014) & " отчёт живёт в жанре списка: 10" & ChrW$(&H2013) & "30 кейсов " & _
        "в неделю на разбор, колонка «итог разбора» заполняется руками.")
    BuildRepeats = s
End Function

Public Function Indent() As String
    Indent = "<span style=""color:var(--muted)"">&nbsp;&nbsp;&nbsp;&nbsp;</span>"
End Function

Public Function BuildDefectDetail() As String
    EnsureCls
    Dim kinds As Object, grps As Object
    Set kinds = CreateObject("Scripting.Dictionary")
    Set grps = CreateObject("Scripting.Dictionary")

    Dim k As Variant, z As Variant, tot As Double, unk As Double, g As String
    tot = 0#: unk = 0#
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) And CDbl(z(Z_DATE)) > 0# Then
            tot = tot + 1#
            AddCnt kinds, CStr(z(Z_KIND)), 1#
            g = Trim$(CStr(z(Z_DEFEKT)))
            If g = "" Then g = "(не указан)"
            AddCnt grps, g, 1#
            If CStr(z(Z_NODE)) = UNCLS Then unk = unk + 1#
        End If
    Next k
    If tot = 0# Then BuildDefectDetail = modContentMTO.EmptyNote(): Exit Function

    Dim kl As Variant, kv As Variant, i As Long
    TopKeys kinds, 0, kl, kv

    Dim s As String
    s = "<div class=""two-col""><div>" & MockLabel("Уровень 0 " & ChrW$(&HB7) & _
        " характер работы")
    s = s & "<table><thead><tr><th>Характер работы</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">%</th></tr></thead><tbody>"
    For i = 0 To UBound(kl)
        s = s & "<tr" & IIf(i = 0, " class=""total""", "") & "><td>" & _
            modContentMTO.Esc(CStr(kl(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(kv(i))) & "</td><td class=""n"">" & _
            FmtF(SafePct(CDbl(kv(i)), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & NoteBlk("Половина «внеплановых» " & ChrW$(&H2014) & " не отказы. «Слив " & _
        "конденсата», «хлорирование», «долив масла», «подкачка колёс» повторяются по " & _
        "регламенту, и в метрике возвратов они давали ложный повтор.") & "</div>"

    s = s & "<div>" & MockLabel("Уровень 1 " & ChrW$(&HB7) & " узел внутри группы")
    s = s & "<div class=""scroll"" style=""max-height:520px;overflow-y:auto"">"
    s = s & "<table><thead><tr><th>Группа и узел</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">%</th></tr></thead><tbody>"

    Dim gl As Variant, gv As Variant, j As Long
    TopKeys grps, 7, gl, gv
    For i = 0 To UBound(gl)
        s = s & "<tr class=""total""><td class=""head"">" & _
            modContentMTO.Esc(Left$(CStr(gl(i)), 44)) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(gv(i))) & "</td><td class=""n"">100,0</td></tr>"

        Dim nd As Object
        Set nd = CreateObject("Scripting.Dictionary")
        For Each k In mZn.Keys
            z = mZn(k)
            If Not IsPlanned(CStr(z(Z_TYPE))) And CDbl(z(Z_DATE)) > 0# Then
                g = Trim$(CStr(z(Z_DEFEKT)))
                If g = "" Then g = "(не указан)"
                If g = CStr(gl(i)) Then AddCnt nd, CStr(z(Z_NODE)), 1#
            End If
        Next k
        Dim nl As Variant, nv As Variant
        TopKeys nd, 6, nl, nv
        For j = 0 To UBound(nl)
            Dim mut As String
            If CStr(nl(j)) = UNCLS Or CStr(nl(j)) = EMPTYD Then
                mut = " style=""color:var(--muted)"""
            Else
                mut = ""
            End If
            s = s & "<tr><td" & mut & ">" & Indent() & _
                modContentMTO.Esc(CStr(nl(j))) & "</td><td class=""n"">" & _
                modContentMTO.FmtInt(CDbl(nv(j))) & "</td><td class=""n"">" & _
                FmtF(SafePct(CDbl(nv(j)), CDbl(gv(i))), 1) & "</td></tr>"
        Next j
    Next i
    s = s & "</tbody></table></div></div></div>"

    s = s & NoteBlk("Описание дефекта приводится к единому классификатору из двух " & _
        "ступеней: характер работы и узел. Узел разбирается правилами «ключевое слово " & _
        ChrW$(&H2192) & " узел», привязанными к своей группе: свободных категорий нет. " & _
        "<b>Остаток «" & UNCLS & "» " & ChrW$(&H2014) & " " & Pc(SafePct(unk, tot), 1) & _
        "</b> и показан отдельной строкой в каждой группе, а не спрятан. Словарь " & _
        "требует доработки на реальных текстах, либо детализацию надо брать из " & _
        "подчинённых уровней справочника <code>" & ChrW$(&H448) & _
        "хКлассификаторДефекта</code>, если они там есть " & ChrW$(&H2014) & _
        " это было бы авторитетнее любого разбора текста.")
    BuildDefectDetail = s
End Function

' =====================================================================================
' СЛАЙД 7. Фазы наряда, возврат техники, хвост незакрытого
' =====================================================================================
Private mFlowReady As Boolean
Private mCloseH As Collection      ' часы «выбытие -> zn_closed»
Private mStuck As Collection       ' приёмка есть, выбытия нет
Private mHang As Collection        ' выбытие есть, наряд не закрыт
Private mTail As Collection        ' нет zn_closed вообще
Private mPairBoth As Long, mPairEq As Long, mMaxAccLev As Double
Private mLong720 As Long

Private Sub EnsureFlow()
    If mFlowReady Then Exit Sub
    EnsureZn
    Set mCloseH = New Collection
    Set mStuck = New Collection
    Set mHang = New Collection
    Set mTail = New Collection
    mPairBoth = 0: mPairEq = 0: mMaxAccLev = 0#: mLong720 = 0

    Dim k As Variant, z As Variant, ac As Double, lv As Double, cl As Double
    For Each k In mZn.Keys
        z = mZn(k)
        ac = ZAcc(z): lv = ZLev(z): cl = CDbl(z(Z_CLOSED))

        If ac > 0# And lv <= 0# Then mStuck.Add CStr(k)
        If lv > 0# And cl <= 0# Then mHang.Add CStr(k)
        If cl <= 0# And CDbl(z(Z_DATE)) > 0# Then mTail.Add CStr(k)

        If lv > 0# And cl > 0# Then
            Dim h As Double
            h = (cl - lv) * 24#
            If h >= 0# And h <= MAX_DUR_H Then mCloseH.Add h
        End If

        If CDbl(z(Z_LEVG)) > 0# And CDbl(z(Z_LEVD)) > 0# Then
            mPairBoth = mPairBoth + 1
            If Abs(CDbl(z(Z_LEVG)) - CDbl(z(Z_LEVD))) < 0.0000116 Then mPairEq = mPairEq + 1
        End If
        If ac > 0# And lv > 0# Then
            If lv - ac > mMaxAccLev Then mMaxAccLev = lv - ac
            If (lv - ac) * 24# > 720# Then mLong720 = mLong720 + 1
        End If
    Next k
    mFlowReady = True
End Sub

' Возраст наряда в сутках на конец снимка.
Private Function AgeDays(ByVal z As Variant) As Double
    AgeDays = Int(SnapshotEnd() - CDbl(z(Z_DATE)))
End Function

Private Function MedDays(ByVal col As Collection) As String
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = 1 To col.Count
        c.Add AgeDays(mZn(col(i)))
    Next i
    Dim hasV As Boolean, m As Double
    m = MedianOf(c, hasV)
    MedDays = IIf(hasV, modContentMTO.FmtInt(m), Dash())
End Function

' Перцентиль по коллекции чисел (0..100).
Public Function PctlOf(ByVal col As Collection, ByVal p As Double, _
                        ByRef hasValue As Boolean) As Double
    hasValue = False
    PctlOf = 0#
    If col Is Nothing Then Exit Function
    If col.Count = 0 Then Exit Function
    Dim a() As Double, i As Long
    ReDim a(0 To col.Count - 1)
    For i = 1 To col.Count
        a(i - 1) = CDbl(col(i))
    Next i
    QSortD a, 0, UBound(a)
    Dim idx As Long
    idx = CLng((UBound(a)) * p / 100#)
    If idx < 0 Then idx = 0
    If idx > UBound(a) Then idx = UBound(a)
    PctlOf = a(idx)
    hasValue = True
End Function

Public Function BuildPhases() As String
    EnsureZn
    Dim wk As Variant
    wk = WeekWindow(8)
    Dim a1() As Variant, a2() As Variant, a3() As Variant, i As Long
    ReDim a1(0 To UBound(wk))
    ReDim a2(0 To UBound(wk))
    ReDim a3(0 To UBound(wk))

    Dim c1 As Collection, c2 As Collection, c3 As Collection
    Dim k As Variant, z As Variant, hasV As Boolean
    Dim s1 As Double, s2 As Double, s3 As Double
    For i = 0 To UBound(wk)
        Set c1 = New Collection: Set c2 = New Collection: Set c3 = New Collection
        For Each k In mZn.Keys
            z = mZn(k)
            If CLng(z(Z_WEEK)) = CLng(wk(i)) Then
                Dim ac As Double, lv As Double, cl As Double
                ac = ZAcc(z): lv = ZLev(z): cl = CDbl(z(Z_CLOSED))
                If ac > 0# And ac >= CDbl(z(Z_DATE)) Then c1.Add (ac - CDbl(z(Z_DATE))) * 24#
                If ac > 0# And lv > 0# And lv >= ac Then c2.Add (lv - ac) * 24#
                If lv > 0# And cl > 0# Then
                    Dim hh2 As Double
                    hh2 = (cl - lv) * 24#
                    If hh2 >= 0# And hh2 <= MAX_DUR_H Then c3.Add hh2
                End If
            End If
        Next k
        a1(i) = MedianOf(c1, hasV)
        a2(i) = MedianOf(c2, hasV)
        a3(i) = MedianOf(c3, hasV)
        If i = UBound(wk) Then
            s1 = CDbl(a1(i)): s2 = CDbl(a2(i)): s3 = CDbl(a3(i))
        End If
    Next i

    Dim s As String
    s = PhasesChart(wk, a1, a2, a3)
    s = s & NoteBlk("Средняя полоса <b>принципиально неразложима</b>: отделить " & _
        "очередь от самого ремонта можно только по истории смены поста, а её в " & _
        "выгрузке нет. До появления этой истории полоса подписана «очередь + ремонт» " & _
        "и в план/факт не превращается. Медианы недели " & _
        WLab(CLng(wk(UBound(wk)))) & ": постановка " & Hh(s1) & ", ремзона " & _
        Hh(s2) & ", закрытие " & Hh(s3) & ".")
    BuildPhases = s
End Function

Public Function BuildReturnKpi() As String
    EnsureFlow
    Dim hasV As Boolean, med As Double, p90 As Double, p99 As Double, mx As Double
    med = MedianOf(mCloseH, hasV)
    p90 = PctlOf(mCloseH, 90#, hasV)
    p99 = PctlOf(mCloseH, 99#, hasV)
    Dim i As Long, tail7 As Long
    tail7 = 0: mx = 0#
    For i = 1 To mCloseH.Count
        If CDbl(mCloseH(i)) > 168# Then tail7 = tail7 + 1
        If CDbl(mCloseH(i)) > mx Then mx = CDbl(mCloseH(i))
    Next i

    Dim s As String
    s = "<div class=""kpis"">"
    s = s & KpiTile("Застряли в ремзоне", modContentMTO.FmtInt(CDbl(mStuck.Count)), "crit", _
        "приёмка есть, выбытия нет " & ChrW$(&HB7) & " медиана " & MedDays(mStuck) & " сут")
    s = s & KpiTile("Готово, но не закрыто", modContentMTO.FmtInt(CDbl(mHang.Count)), "crit", _
        "выбытие подписано, наряд открыт " & ChrW$(&HB7) & " медиана " & MedDays(mHang) & " сут")
    s = s & KpiTile("Выбытие " & ChrW$(&H2192) & " закрытие", Hh(med), "", _
        "медиана; p90 " & Hh(p90) & ", p99 " & Hh(p99))
    s = s & KpiTile("Хвост дольше 7 суток", modContentMTO.FmtInt(CDbl(tail7)), "crit", _
        "максимум " & FmtF(mx / 24#, 1) & " сут")
    BuildReturnKpi = s & "</div>"
End Function

Public Function BuildReturnHist() As String
    EnsureFlow
    Dim labs As Variant, lim As Variant
    labs = Array("до 1 ч", "1" & ChrW$(&H2013) & "4 ч", "4" & ChrW$(&H2013) & "12 ч", _
        "12" & ChrW$(&H2013) & "24 ч", "1" & ChrW$(&H2013) & "3 сут", _
        "3" & ChrW$(&H2013) & "7 сут", "7 сут +")
    lim = Array(1#, 4#, 12#, 24#, 72#, 168#)

    Dim vals(0 To 6) As Double, i As Long, j As Long
    For i = 1 To mCloseH.Count
        Dim h As Double, b As Long
        h = CDbl(mCloseH(i))
        b = 6
        For j = 0 To 5
            If h < CDbl(lim(j)) Then
                b = j
                Exit For
            End If
        Next j
        vals(b) = vals(b) + 1#
    Next i
    Dim vv() As Variant
    ReDim vv(0 To 6)
    For i = 0 To 6
        vv(i) = vals(i)
    Next i

    Dim s As String
    s = "<div class=""two-col""><div>" & HBars(labs, vv, 560, 110, 25)
    s = s & NoteBlk(modContentMTO.FmtInt(vals(0)) & " из " & _
        modContentMTO.FmtInt(CDbl(mCloseH.Count)) & " нарядов закрываются в тот же час " & _
        ChrW$(&H2014) & " это оформление, а не простой техники.") & "</div><div>"
    s = s & MockLabel("По площадкам")
    s = s & "<table><thead><tr><th>Площадка</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">Медиана</th></tr></thead><tbody>"

    Dim byP As Object, cnt As Object
    Set byP = CreateObject("Scripting.Dictionary")
    Set cnt = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant, p As String
    For Each k In mZn.Keys
        z = mZn(k)
        Dim lv As Double, cl As Double
        lv = ZLev(z): cl = CDbl(z(Z_CLOSED))
        If lv > 0# And cl > 0# Then
            Dim hh3 As Double
            hh3 = (cl - lv) * 24#
            If hh3 >= 0# And hh3 <= MAX_DUR_H Then
                p = Trim$(CStr(z(Z_POST)))
                If p = "" Then p = "(пост не указан)"
                If Not byP.Exists(p) Then byP.Add p, New Collection
                byP(p).Add hh3
                AddCnt cnt, p, 1#
            End If
        End If
    Next k
    Dim pl As Variant, pv As Variant
    TopKeys cnt, 8, pl, pv
    For i = 0 To UBound(pl)
        Dim hasV As Boolean, m As Double
        m = MedianOf(byP(CStr(pl(i))), hasV)
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(pl(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(pv(i))) & "</td><td class=""n"">" & _
            IIf(hasV, Hh(m), Dash()) & "</td></tr>"
    Next i
    s = s & "</tbody></table></div></div>"
    s = s & NoteBlk("<b>Здесь нельзя посчитать главное.</b> «Готово, но не забрано» " & _
        ChrW$(&H2014) & " это разница между подписью выбытия ДГМ и ДЭНТ, а в " & _
        "заказ-наряде на обе дирекции <b>одно поле даты</b> (<code>" & ChrW$(&H442) & _
        "1кДатаВыдачи</code>), различаются только сотрудники. Разность тождественно " & _
        "ноль во всех " & modContentMTO.FmtInt(CDbl(mPairEq)) & " парах " & _
        ChrW$(&H2014) & " по построению, а не по данным. Поэтому настоящий простой " & _
        "техники после ремонта виден только косвенно: через две таблицы ниже.")
    BuildReturnHist = s
End Function

' Первые n нарядов коллекции по возрастанию даты создания (самые старые сверху).
Private Function OldestOf(ByVal col As Collection, ByVal n As Long) As Variant
    Dim m As Long
    m = col.Count
    If m = 0 Then OldestOf = Array(): Exit Function
    Dim ds() As Double, ns() As String, i As Long
    ReDim ds(0 To m - 1)
    ReDim ns(0 To m - 1)
    For i = 1 To m
        ns(i - 1) = CStr(col(i))
        ds(i - 1) = CDbl(mZn(ns(i - 1))(Z_DATE))
    Next i
    QSortPair ds, ns, 0, m - 1
    Dim take As Long
    take = n
    If take > m Then take = m
    Dim r() As Variant
    ReDim r(0 To take - 1)
    For i = 0 To take - 1
        r(i) = ns(i)
    Next i
    OldestOf = r
End Function

Private Function PostOr(ByVal z As Variant) As String
    Dim p As String
    p = Trim$(CStr(z(Z_POST)))
    If p = "" Then p = "(не указан)"
    PostOr = modContentMTO.Esc(p)
End Function

Public Function BuildReturnStuck() As String
    EnsureFlow
    Dim rows As Variant
    rows = OldestOf(mStuck, 8)
    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Наряд</th><th>Машина</th>" & _
        "<th class=""n"">Возраст</th><th>Приёмка</th><th>Пост</th>" & _
        "<th>Статус по документу</th></tr></thead><tbody>"
    Dim i As Long, z As Variant
    For i = 0 To UBound(rows)
        z = mZn(CStr(rows(i)))
        s = s & "<tr><td class=""mono"">" & modContentMTO.Esc(CStr(rows(i))) & "</td>"
        s = s & "<td class=""mono"">" & modContentMTO.Esc(CStr(z(Z_VEH))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(AgeDays(z)) & " сут</td>"
        s = s & "<td>" & DMon(ZAcc(z)) & "</td><td>" & PostOr(z) & "</td>"
        s = s & "<td style=""color:var(--ink-2)"">" & modContentMTO.Esc(CStr(z(Z_TEK))) & _
            "</td></tr>"
    Next i
    s = s & "</tbody></table></div>"
    s = s & NoteBlk("Всего таких " & modContentMTO.FmtInt(CDbl(mStuck.Count)) & _
        ", медиана возраста " & MedDays(mStuck) & " суток. Это техника, которую " & _
        "приняли в ремзону и по документам не выпустили. Либо она физически стоит, " & _
        "либо её отдали без подписи " & ChrW$(&H2014) & " и то и другое разбирается " & _
        "поимённо, а не процентом.")
    BuildReturnStuck = s
End Function

Public Function BuildReturnHang() As String
    EnsureFlow
    Dim rows As Variant
    rows = OldestOf(mHang, 8)
    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Наряд</th><th>Машина</th>" & _
        "<th class=""n"">Возраст</th><th>Выбытие</th><th>Пост</th>" & _
        "<th>Статус по документу</th></tr></thead><tbody>"
    Dim i As Long, z As Variant
    For i = 0 To UBound(rows)
        z = mZn(CStr(rows(i)))
        s = s & "<tr><td class=""mono"">" & modContentMTO.Esc(CStr(rows(i))) & "</td>"
        s = s & "<td class=""mono"">" & modContentMTO.Esc(CStr(z(Z_VEH))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(AgeDays(z)) & " сут</td>"
        s = s & "<td>" & DMon(ZLev(z)) & "</td><td>" & PostOr(z) & "</td>"
        s = s & "<td style=""color:var(--ink-2)"">" & modContentMTO.Esc(CStr(z(Z_TEK))) & _
            "</td></tr>"
    Next i
    s = s & "</tbody></table></div>"
    s = s & NoteBlk("Всего " & modContentMTO.FmtInt(CDbl(mHang.Count)) & ", медиана " & _
        "возраста " & MedDays(mHang) & " суток. Выбытие подписано, значит техника у " & _
        "владельца, но наряд в 1С остался открытым. Для ремзоны это мусор в " & _
        "отчётности, для экономики " & ChrW$(&H2014) & " незакрытые затраты. Адресат " & _
        ChrW$(&H2014) & " не мастер, а тот, кто закрывает документы.")
    BuildReturnHang = s
End Function

Public Function BuildTailAge() As String
    EnsureFlow
    Dim labs As Variant
    labs = Array("0" & ChrW$(&H2013) & "3", "3" & ChrW$(&H2013) & "7", _
        "7" & ChrW$(&H2013) & "14", "14" & ChrW$(&H2013) & "30", "30 +")
    Dim vals(0 To 4) As Double, i As Long, a As Double
    For i = 1 To mTail.Count
        a = AgeDays(mZn(mTail(i)))
        If a < 3# Then
            vals(0) = vals(0) + 1#
        ElseIf a < 7# Then
            vals(1) = vals(1) + 1#
        ElseIf a < 14# Then
            vals(2) = vals(2) + 1#
        ElseIf a < 30# Then
            vals(3) = vals(3) + 1#
        Else
            vals(4) = vals(4) + 1#
        End If
    Next i
    Dim vv() As Variant
    ReDim vv(0 To 4)
    For i = 0 To 4
        vv(i) = vals(i)
    Next i
    BuildTailAge = Cols(labs, vv, 440, 200) & NoteBlk("Всего без <code>zn_closed</code> " & _
        ChrW$(&H2014) & " " & modContentMTO.FmtInt(CDbl(mTail.Count)) & " нарядов, из них " & _
        modContentMTO.FmtInt(vals(4)) & " старше 30 суток.")
End Function

Public Function BuildTailWhy() As String
    EnsureFlow
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, t As String
    For i = 1 To mTail.Count
        t = Trim$(CStr(mZn(mTail(i))(Z_TEK)))
        If t = "" Then t = "(пусто)"
        AddCnt d, t, 1#
    Next i
    If d.Count = 0 Then BuildTailWhy = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    TopKeys d, 6, labs, vals
    BuildTailWhy = HBars(labs, vals, 560, 300, 24) & NoteBlk( _
        "«Отменен, требует повторного планирования» " & ChrW$(&H2014) & " не задержка " & _
        "ремонта, а брошенный наряд: его надо закрывать в 1С, а не разбирать на " & _
        "планёрке. Три верхние строки " & ChrW$(&H2014) & " три разных адресата и три " & _
        "разных разговора.")
End Function

' Последнее подписанное событие наряда - это диагноз задержки.
Private Function LastEv(ByVal z As Variant) As String
    Dim t(0 To 3) As Double, l(0 To 3) As String, i As Long
    t(0) = CDbl(z(Z_ACCG)): l(0) = "приёмка ДГМ"
    t(1) = CDbl(z(Z_ACCD)): l(1) = "приёмка ДЭНТ"
    t(2) = CDbl(z(Z_LEVG)): l(2) = "выбытие ДГМ"
    t(3) = CDbl(z(Z_LEVD)): l(3) = "выбытие ДЭНТ"
    Dim bi As Long, bv As Double
    bi = -1: bv = 0#
    For i = 0 To 3
        If t(i) > bv Then bv = t(i): bi = i
    Next i
    If bi < 0 Then
        LastEv = "подписей нет"
    Else
        LastEv = l(bi) & ", " & Format$(CDate(bv), "dd.mm")
    End If
End Function

Public Function BuildTailRows() As String
    EnsureFlow
    Dim old As Collection, i As Long
    Set old = New Collection
    For i = 1 To mTail.Count
        If AgeDays(mZn(mTail(i))) >= 14# Then old.Add mTail(i)
    Next i
    Dim rows As Variant
    rows = OldestOf(old, 8)

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Наряд</th><th>Машина</th>" & _
        "<th class=""n"">Возраст</th><th>Последнее событие</th><th>Пост</th>" & _
        "<th>Статус по документу</th></tr></thead><tbody>"
    Dim z As Variant
    For i = 0 To UBound(rows)
        z = mZn(CStr(rows(i)))
        s = s & "<tr><td class=""mono"">" & modContentMTO.Esc(CStr(rows(i))) & "</td>"
        s = s & "<td class=""head"">" & modContentMTO.Esc(CStr(z(Z_VEH))) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(AgeDays(z)) & " сут</td>"
        s = s & "<td>" & modContentMTO.Esc(LastEv(z)) & "</td><td>" & PostOr(z) & "</td>"
        s = s & "<td style=""color:var(--ink-2)"">" & modContentMTO.Esc(CStr(z(Z_TEK))) & _
            "</td></tr>"
    Next i
    s = s & "</tbody></table></div>"
    s = s & NoteBlk("Колонка «последнее событие» " & ChrW$(&H2014) & " это диагноз " & _
        "задержки: ждём ремонт, ждём подпись владельца или наряд вообще не заведён в " & _
        "работу. Разбирать надо " & modContentMTO.FmtInt(CDbl(old.Count)) & _
        " нарядов, а не " & modContentMTO.FmtInt(CDbl(mTail.Count)) & ".")
    BuildTailRows = s
End Function

' Отчёты, снятые с публикации: пока строка открыта, отчёт не показывается
' с оговоркой мелким шрифтом, а не публикуется.
Public Function BuildLimits() As String
    EnsureFlow
    Dim s As String
    s = "<table><thead><tr><th>Отчёт</th><th>Почему</th><th>Что нужно от 1С</th>" & _
        "</tr></thead><tbody>"
    s = s & "<tr><td class=""head"">Готово, но не забрано</td><td>в заказ-наряде " & _
        "<b>одно поле даты на обе дирекции</b>: <code>" & ChrW$(&H442) & _
        "1кДатаПриемки</code> и <code>" & ChrW$(&H442) & "1кДатаВыдачи</code>. " & _
        "Различаются только сотрудники " & ChrW$(&H2014) & " <code>" & ChrW$(&H442) & _
        "1кПринялВПриемку</code> против <code>" & ChrW$(&H442) & _
        "1кСдалВПриемку</code>. Разность подписей ноль по построению, а не по данным: " & _
        modContentMTO.FmtInt(CDbl(mPairEq)) & " пар из " & _
        modContentMTO.FmtInt(CDbl(mPairBoth)) & "</td>" & _
        "<td>отдельная отметка времени у каждой дирекции в документе ЗН</td></tr>"
    s = s & "<tr><td class=""head"">План против факта длительности</td><td>план есть " & _
        ChrW$(&H2014) & " <code>hourdlit</code>, плановая длительность в часах. Не " & _
        "хватает факта: интервал приёмка " & ChrW$(&H2192) & " выбытие включает " & _
        "очередь и ожидание запчастей, максимум по снимку " & _
        FmtF(mMaxAccLev, 1) & " сут. Сравнение дало бы не срыв срока, а свойство " & _
        "метрики</td><td>отметки начала и конца работ либо история смены поста</td></tr>"
    s = s & "<tr><td class=""head"">Наработка и межремонтный интервал</td><td><code>" & _
        "odometer</code> и <code>engine_hours</code> " & ChrW$(&H2014) & " <b>срез " & _
        "последних</b> показаний счётчика на момент выгрузки, а не на момент наряда, " & _
        "и оба падают в 0, если показаний нет. Разница между двумя ремонтами по ним " & _
        "не считается</td><td>показание счётчика на дату заказ-наряда + тип " & _
        "применимого счётчика</td></tr>"
    s = s & "</tbody></table>"
    s = s & NoteBlk("Строки этой таблицы не декоративны: пока они открыты, отчёты, " & _
        "которые на них опираются, из презентации сняты, а не показаны с оговоркой " & _
        "мелким шрифтом.")
    BuildLimits = s
End Function

' =====================================================================================
' СЛАЙД 8. Материалы и качество учёта
' =====================================================================================
Private Function AbcSorted(ByRef nm() As String, ByRef sm() As Double) As Long
    EnsureVeh
    Dim n As Long, k As Variant
    n = 0
    For Each k In mVeh.Keys
        If CDbl(mVeh(k)(V_PARTS)) > 0# Then n = n + 1
    Next k
    AbcSorted = n
    If n = 0 Then Exit Function
    ReDim nm(0 To n - 1)
    ReDim sm(0 To n - 1)
    Dim kk() As Double, vv() As String, i As Long
    ReDim kk(0 To n - 1)
    ReDim vv(0 To n - 1)
    i = 0
    For Each k In mVeh.Keys
        If CDbl(mVeh(k)(V_PARTS)) > 0# Then
            kk(i) = -CDbl(mVeh(k)(V_PARTS))
            vv(i) = CStr(k)
            i = i + 1
        End If
    Next k
    QSortPair kk, vv, 0, n - 1
    For i = 0 To n - 1
        nm(i) = vv(i)
        sm(i) = -kk(i)
    Next i
End Function

Public Function BuildKpiParts() As String
    EnsureVeh
    Dim nm() As String, sm() As Double, n As Long, i As Long, tot As Double
    n = AbcSorted(nm, sm)
    If n = 0 Then BuildKpiParts = modContentMTO.EmptyNote(): Exit Function
    tot = 0#
    For i = 0 To n - 1
        tot = tot + sm(i)
    Next i

    Dim acc As Double, aCars As Long, aSum As Double
    acc = 0#: aCars = 0: aSum = 0#
    For i = 0 To n - 1
        If tot > 0# And acc / tot < 0.8 Then
            aCars = aCars + 1
            aSum = aSum + sm(i)
        End If
        acc = acc + sm(i)
    Next i

    Dim s As String
    s = "<div class=""kpis"">"
    s = s & KpiTile("Материалы за снимок", Rub(tot), "", "единственное денежное поле выгрузки")
    s = s & KpiTile("Машин с расходом", modContentMTO.FmtInt(CDbl(n)), "", _
        "из " & modContentMTO.FmtInt(CDbl(mVeh.Count)) & " в парке")
    s = s & KpiTile("Группа A", modContentMTO.FmtInt(CDbl(aCars)) & " <small>машин</small>", _
        "crit", Pc(SafePct(aSum, tot), 1) & " всех денег")
    s = s & KpiTile("Самая дорогая", "<span class=""mono"" style=""font-size:26px"">" & _
        modContentMTO.Esc(nm(0)) & "</span>", "crit", Rub(sm(0)))
    BuildKpiParts = s & "</div>"
End Function

Public Function BuildAbc() As String
    EnsureVeh
    Dim nm() As String, sm() As Double, n As Long, i As Long, tot As Double
    n = AbcSorted(nm, sm)
    If n = 0 Then BuildAbc = modContentMTO.EmptyNote(): Exit Function
    tot = 0#
    For i = 0 To n - 1
        tot = tot + sm(i)
    Next i

    Dim cars(0 To 2) As Double, sums(0 To 2) As Double, acc As Double, g As Long
    acc = 0#
    For i = 0 To n - 1
        Dim share As Double
        If tot > 0# Then share = acc / tot Else share = 0#
        If share < 0.8 Then
            g = 0
        ElseIf share < 0.95 Then
            g = 1
        Else
            g = 2
        End If
        cars(g) = cars(g) + 1#
        sums(g) = sums(g) + sm(i)
        acc = acc + sm(i)
    Next i

    Dim s As String, gl As Variant
    gl = Array("A", "B", "C")
    s = "<div class=""two-col""><div><table><thead><tr><th>Группа</th>" & _
        "<th class=""n"">Машин</th><th class=""n"">Материалы</th>" & _
        "<th class=""n"">% расхода</th></tr></thead><tbody>"
    For i = 0 To 2
        s = s & "<tr><td class=""head"">" & CStr(gl(i)) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(cars(i)) & "</td><td class=""n"">" & Rub(sums(i)) & _
            "</td><td class=""n"">" & FmtF(SafePct(sums(i), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & NoteBlk(modContentMTO.FmtInt(cars(0)) & " машин из " & _
        modContentMTO.FmtInt(CDbl(n)) & " дают " & Pc(SafePct(sums(0), tot), 1) & _
        " закупки " & ChrW$(&H2014) & " заявка формируется по ним.") & "</div>"

    s = s & "<div>" & MockLabel("Топ по расходу материалов")
    s = s & "<table><thead><tr><th>Гар. " & ChrW$(&H2116) & "</th><th>Группа техники</th>" & _
        "<th class=""n"">Возраст</th><th class=""n"">Материалы</th></tr></thead><tbody>"
    Dim top As Long
    top = 8
    If top > n Then top = n
    For i = 0 To top - 1
        Dim v As Variant, ag As Double
        v = mVeh(nm(i))
        ag = AgeYears(CDbl(v(V_MADE)))
        s = s & "<tr><td class=""n mono"" style=""text-align:left;font-size:14px"">" & _
            modContentMTO.Esc(nm(i)) & "</td><td style=""color:var(--ink-2)"">" & _
            modContentMTO.Esc(CStr(v(V_GROUP))) & "</td><td class=""n"">" & _
            IIf(ag >= 0#, FmtF(ag, 1), Dash()) & "</td><td class=""n"">" & Rub(sm(i)) & _
            "</td></tr>"
    Next i
    s = s & "</tbody></table></div></div>"
    s = s & NoteBlk("<b>Только материалы.</b> <code>cost_Trudozatrat</code> хранит " & _
        "часы, а не рубли: перевести их в деньги можно лишь умножением на ставку " & _
        "нормо-часа, которой в выгрузке нет. Поэтому отчёт нигде не называется " & _
        "«стоимостью ремонта» " & ChrW$(&H2014) & " иначе первое же сравнение с " & _
        "бухгалтерией его похоронит.")
    BuildAbc = s
End Function

Public Function BuildMoneyDefekt() As String
    EnsureZn
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, z As Variant, g As String, tot As Double
    tot = 0#
    For Each k In mZn.Keys
        z = mZn(k)
        If Not IsPlanned(CStr(z(Z_TYPE))) Then
            g = Trim$(CStr(z(Z_DEFEKT)))
            If g = "" Then g = NOSECT
            AddCnt d, g, CDbl(z(Z_PARTS))
            tot = tot + CDbl(z(Z_PARTS))
        End If
    Next k
    If tot <= 0# Then BuildMoneyDefekt = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant, i As Long
    TopKeys d, 6, labs, vals
    Dim s As String
    s = "<table><thead><tr><th>Раздел</th><th class=""n"">Материалы</th>" & _
        "<th class=""n"">%</th></tr></thead><tbody>"
    For i = 0 To UBound(labs)
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            Rub(CDbl(vals(i))) & "</td><td class=""n"">" & _
            FmtF(SafePct(CDbl(vals(i)), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & NoteBlk("Список «часто» (слайд 6) и список «дорого» " & ChrW$(&H2014) & _
        " разные. Складывать их в один рейтинг нельзя: у них разная единица счёта.")
    BuildMoneyDefekt = s
End Function

Private Function QRow(ByVal d As String, ByVal u As String, ByVal v As String, _
                      ByVal who As String, ByVal cls As String, ByVal st As String) As String
    QRow = "<tr><td>" & d & "</td><td style=""color:var(--muted)"">" & u & _
        "</td><td class=""n"">" & v & "</td><td>" & who & "</td><td>" & _
        Pill(cls, st) & "</td></tr>"
End Function

Public Function BuildQuality() As String
    EnsureFlow
    EnsureVeh
    Dim noPost As Double, noClosed As Double, k As Variant
    noPost = 0#: noClosed = 0#
    For Each k In mZn.Keys
        If Trim$(CStr(mZn(k)(Z_POST))) = "" Then noPost = noPost + 1#
        If CDbl(mZn(k)(Z_CLOSED)) <= 0# Then noClosed = noClosed + 1#
    Next k
    Dim nz As Double
    nz = CDbl(mZn.Count)

    Dim T As String
    T = ChrW$(&H442)

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Дефект</th><th>Единица</th>" & _
        "<th class=""n"">Сейчас</th><th>Кто чинит</th><th>Статус</th>" & _
        "</tr></thead><tbody>"
    s = s & QRow("Наряд без поста ремзоны (<code>post</code> = <code>" & T & _
        "1кПостРемзоны.Родитель</code>)", "наряд", Pc(SafePct(noPost, nz), 1), _
        "ДГМ + 1С", "crit", "открыт")
    s = s & QRow("<code>post</code> " & ChrW$(&H2014) & _
        " родитель поста, сам пост не выгружается", "поле", "100 %", _
        "1С (выгрузка)", "crit", "открыт")
    s = s & QRow("Одно поле даты на обе дирекции (<code>" & T & "1кДатаПриемки</code> / " & _
        "<code>" & T & "1кДатаВыдачи</code>)", "пара подписей", _
        modContentMTO.FmtInt(CDbl(mPairEq)) & " из " & _
        modContentMTO.FmtInt(CDbl(mPairBoth)), "1С (документ ЗН)", "crit", "открыт")
    s = s & QRow("<code>arm</code> = «НЕ ПОДПИСАНО» (нет даты статуса)", "событие", _
        Pc(SafePct(CDbl(mArmNone), CDbl(mRows)), 1), "дирекции", "warn", "в работе")
    s = s & QRow("<code>model_type</code> " & ChrW$(&H2014) & _
        " константа: выгрузка отбирает только Автотранспорт", "поле", "100 %", _
        "1С (условие отбора)", "", "по построению")
    s = s & QRow("<code>vehicle_group</code> без укрупняющего классификатора", _
        "справочник", "~100 значений", "ДГМ (НСИ)", "crit", "открыт")
    s = s & QRow("<code>odometer</code>/<code>engine_hours</code> " & ChrW$(&H2014) & _
        " срез последних, не на дату ЗН", "событие", _
        Pc(SafePct(CDbl(mNoCounter), CDbl(mRows)), 1) & " обе пусты", _
        "1С (выгрузка)", "crit", "открыт")
    s = s & QRow("<code>MadeYear</code> " & ChrW$(&H2014) & " дата, а не год", "поле", _
        "100 %", "1С (выгрузка)", "good", "обходится")
    s = s & QRow("«Отдел договорной работы» есть и в списке ДГМ, и в списке ДЭНТ " & _
        ChrW$(&H2014) & " <code>emp_dep</code> всегда даёт ДГМ", "подразделение", _
        "1 из 10", "1С (параметры запроса)", "crit", "открыт")
    s = s & QRow("<code>zn_closed</code> " & ChrW$(&H2014) & _
        " только при текущем статусе «Закрыт» (срез последних)", "наряд", _
        Pc(SafePct(noClosed, nz), 1) & " пусто", "1С (регистр статусов)", _
        "", "по построению")
    s = s & QRow("<code>zn_closed</code> раньше <code>date</code>", "наряд", _
        modContentMTO.FmtInt(CDbl(mBadClosed)), "1С (выгрузка)", "crit", "открыт")
    s = s & QRow("Интервал приёмка " & ChrW$(&H2192) & " выбытие &gt; 720 ч", "пара", _
        modContentMTO.FmtInt(CDbl(mLong720)), "ДГМ (разбор)", "warn", "в работе")
    s = s & QRow("<code>in_bounds</code> = «последняя отметка чекина раньше начала " & _
        "периода» " & ChrW$(&H2014) & " для отчёта смысла не несёт", "событие", _
        Pc(SafePct(CDbl(mNoBounds), CDbl(mRows)), 1), _
        "1С " & ChrW$(&H2014) & " уточнить смысл", "crit", "открыт")
    s = s & "</tbody></table></div>"
    s = s & NoteBlk("Единица счёта проставлена у каждой строки намеренно: «33 % " & _
        "нарядов» и «24 % событий» стоят рядом и не складываются. Без этой колонки " & _
        "таблица через месяц начнёт врать. Закрытая строка исчезает с панели; пока " & _
        "строка открыта, опирающиеся на неё отчёты публикуются с пометкой о " & _
        "недостоверности.")
    BuildQuality = s
End Function

Private Function RRow(ByVal n As Long, ByVal what As String, ByVal unlocks As String) As String
    RRow = "<tr><td class=""n"">" & CStr(n) & "</td><td>" & what & "</td><td>" & _
        unlocks & "</td></tr>"
End Function

' Заявка в 1С. Текст фиксирован: это не расчёт, а перечень открытых требований.
Public Function BuildRequest() As String
    Dim T As String
    T = ChrW$(&H442)
    Dim s As String
    s = "<table><thead><tr><th class=""n"">" & ChrW$(&H2116) & "</th><th>Что нужно</th>" & _
        "<th>Что разблокирует</th></tr></thead><tbody>"
    s = s & RRow(1, "Отдельная отметка времени подписи у каждой дирекции в документе " & _
        "ЗН. Сейчас на обе стоит одно поле " & ChrW$(&H2014) & " <code>" & T & _
        "1кДатаПриемки</code> и <code>" & T & "1кДатаВыдачи</code>, различаются только " & _
        "сотрудники", "«Готово, но не забрано», синхронность дирекций, реальный " & _
        "простой техники")
    s = s & RRow(2, "Выгружать сам <code>" & T & "1кПостРемзоны</code>, а не его " & _
        "<code>.Родитель</code>, и сделать поле обязательным в форме ЗН", _
        "все разрезы по постам на слайдах 2, 3, 5, 7; снимет «пост не указан»")
    s = s & RRow(3, "Укрупняющий классификатор поверх <code>" & T & _
        "1кГруппаТехники</code> " & ChrW$(&H2014) & " сейчас в нём ~100 значений " & _
        "вплоть до отдельных моделей водил", _
        "кривую старения и Парето в разрезе видов техники")
    s = s & RRow(4, "Показание счётчика <b>на дату заказ-наряда</b> плюс тип " & _
        "применимого счётчика. Сейчас выгружается срез последних на момент выгрузки", _
        "наработку и межремонтный интервал целиком")
    s = s & RRow(5, "Ставка нормо-часа к <code>ЧасыТрудозатрат</code>", _
        "полную стоимость ремонта вместо одних материалов")
    s = s & RRow(6, "История смены поста или отметки начала и конца работ", _
        "честный план/факт длительности вместо интервала с очередью внутри")
    s = s & RRow(7, "Убрать <code>in_bounds</code> либо задокументировать: сейчас это " & _
        "«последняя отметка чекина раньше начала периода отбора», к качеству данных " & _
        "отношения не имеет", "снимет неопределённость в базовом фильтре")
    s = s & RRow(8, "«Отдел договорной работы» стоит в списках и ДГМ, и ДЭНТ; в " & _
        "<code>ВЫБОР</code> первым проверяется ДГМ, поэтому сотрудники ОДР всегда " & _
        "попадают в ДГМ", "корректный <code>emp_dep</code>")
    BuildRequest = s & "</tbody></table>"
End Function

' =====================================================================================
' Точка входа: заполнение плейсхолдеров слайдов 5-8.
' Ключи и их состав зафиксированы в docs/plans/MTO_контракт_шаблона_v1.0.md.
' =====================================================================================
Public Sub FillZonePlaceholders(ByVal d As Object)
    Dim t0 As Single
    t0 = Timer

    d("KPI_OVERVIEW") = BuildKpiOverview()
    d("BLOCK_TIME_HIST") = BuildTimeHist()
    d("BLOCK_FLOW_ZNTYPE") = BuildFlowZnType()
    d("BLOCK_FLOW_DEFEKT") = BuildFlowDefekt()
    d("BLOCK_NOPOST_WEEKLY") = BuildNoPostWeekly()
    modLog.WriteDebug 1, "Техника", "FillZonePlaceholders", _
        "Слайд 1 готов: " & Round(Timer - t0, 2) & " c"

    d("KPI_FLEET") = BuildKpiFleet()
    d("BLOCK_POSTS_WEEK") = BuildPostsWeek()
    d("BLOCK_AGE_CURVE") = BuildAgeCurve()
    d("BLOCK_AGE_MATRIX") = BuildAgeMatrix()
    d("BLOCK_AGING") = BuildAging()
    d("BLOCK_PACK") = BuildPack()
    modLog.WriteDebug 1, "Техника", "FillZonePlaceholders", _
        "Слайд 5 готов: " & Round(Timer - t0, 2) & " c"

    d("BLOCK_CHRONICS") = BuildChronics()
    d("BLOCK_PARETO") = BuildPareto()
    d("BLOCK_DEFECT_DETAIL") = BuildDefectDetail()
    d("BLOCK_RET_KPI") = BuildRetKpi()
    d("BLOCK_RET_MONTH") = BuildRetMonth()
    d("BLOCK_RET_WEEK") = BuildRetWeek()
    d("BLOCK_RET_NODE") = BuildRetNode()
    d("BLOCK_REPEATS") = BuildRepeats()
    modLog.WriteDebug 1, "Техника", "FillZonePlaceholders", _
        "Слайд 6 готов: " & Round(Timer - t0, 2) & " c"

    d("BLOCK_PHASES") = BuildPhases()
    d("BLOCK_RETURN_KPI") = BuildReturnKpi()
    d("BLOCK_RETURN_HIST") = BuildReturnHist()
    d("BLOCK_RETURN_STUCK") = BuildReturnStuck()
    d("BLOCK_RETURN_HANG") = BuildReturnHang()
    d("BLOCK_TAIL_AGE") = BuildTailAge()
    d("BLOCK_TAIL_WHY") = BuildTailWhy()
    d("BLOCK_TAIL_ROWS") = BuildTailRows()
    d("BLOCK_LIMITS") = BuildLimits()
    modLog.WriteDebug 1, "Техника", "FillZonePlaceholders", _
        "Слайд 7 готов: " & Round(Timer - t0, 2) & " c"

    d("KPI_PARTS") = BuildKpiParts()
    d("BLOCK_ABC") = BuildAbc()
    d("BLOCK_MONEY_DEFEKT") = BuildMoneyDefekt()
    d("BLOCK_QUALITY") = BuildQuality()
    d("BLOCK_REQUEST") = BuildRequest()
    modLog.WriteDebug 1, "Техника", "FillZonePlaceholders", _
        "Слайд 8 готов: " & Round(Timer - t0, 2) & " c"
End Sub

' Сводка чисел части «Техника» для промпта ИИ. Только готовые числа:
' ИИ интерпретирует, но ничего не считает и не судит о людях.
Public Function ZoneFactsJson() As String
    EnsureVeh
    EnsureRet
    EnsureFlow
    Dim col As Collection, i As Long, multi As Long
    Set col = mVisitSizes(CStr(GAP_HOURS))
    multi = 0
    For i = 1 To col.Count
        If CLng(col(i)) > 1 Then multi = multi + 1
    Next i

    Dim s As String
    s = "{""veh"":" & CStr(mVeh.Count) & _
        ",""visits"":" & CStr(col.Count) & _
        ",""visits_multi_pct"":" & JNum(SafePct(CDbl(multi), CDbl(col.Count))) & _
        ",""unplanned"":" & CStr(mDenAll) & _
        ",""fail"":" & CStr(mDenFail) & _
        ",""ret_fail_pct"":" & JNum(SafePct(mFailTot, mDenFail)) & _
        ",""ret_grp_pct"":" & JNum(SafePct(mRetTot, mDenAll)) & _
        ",""stuck"":" & CStr(mStuck.Count) & _
        ",""hang"":" & CStr(mHang.Count) & _
        ",""tail"":" & CStr(mTail.Count) & _
        ",""report_week"":" & CStr(ZoneReportWeek()) & "}"
    ZoneFactsJson = s
End Function

Private Function JNum(ByVal v As Double) As String
    JNum = Replace$(Format$(v, "0.0"), ",", ".")
End Function

' =====================================================================================
' Примитивы, общие со слайдами 2-4 (модуль modContentDisc)
' =====================================================================================
' Спарклайн плитки KPI. Меньше двух точек - пусто.
Public Function Spark(ByVal series As Variant) As String
    Const W As Long = 92
    Const H As Long = 24
    Dim n As Long, i As Long, lo As Double, hi As Double
    n = UBound(series) - LBound(series) + 1
    If n < 2 Then Spark = "": Exit Function
    lo = CDbl(series(LBound(series))): hi = lo
    For i = LBound(series) To UBound(series)
        If CDbl(series(i)) < lo Then lo = CDbl(series(i))
        If CDbl(series(i)) > hi Then hi = CDbl(series(i))
    Next i
    Dim rng As Double
    rng = hi - lo
    If rng = 0# Then rng = 1#

    Dim pts As String, lx As Double, ly As Double
    pts = ""
    For i = 0 To n - 1
        Dim x As Double, y As Double
        x = 2# + (W - 4) * i / (n - 1)
        y = H - 5# - (H - 10) * (CDbl(series(LBound(series) + i)) - lo) / rng
        If pts <> "" Then pts = pts & " "
        pts = pts & Cx(x) & "," & Cx(y)
        lx = x: ly = y
    Next i
    Spark = "<svg viewBox=""0 0 " & CStr(W) & " " & CStr(H) & """ width=""" & CStr(W) & _
        """ height=""" & CStr(H) & """ role=""img"" aria-hidden=""true"">" & _
        "<polyline points=""" & pts & """ fill=""none"" stroke=""var(--s1)"" " & _
        "stroke-width=""2"" stroke-linejoin=""round""/><circle cx=""" & Cx(lx) & _
        """ cy=""" & Cx(ly) & """ r=""2.6"" fill=""var(--s1)""/></svg>"
End Function

' Дельта к предыдущей неделе. kind: 0 - число, 1 - п.п., 2 - % от базы, 3 - часы.
' upGood = True, если рост - это хорошо.
Public Function DeltaSpan(ByVal cur As Double, ByVal prev As Double, ByVal kind As Long, _
                          ByVal upGood As Boolean, ByVal hasPrev As Boolean) As String
    If Not hasPrev Then
        DeltaSpan = "<span class=""delta flat"">нет базы</span>"
        Exit Function
    End If
    Dim d As Double, arrow As String, txt As String
    d = cur - prev
    If d > 0# Then
        arrow = ChrW$(&H25B2) & " "
    ElseIf d < 0# Then
        arrow = ChrW$(&H25BC) & " "
    Else
        arrow = "= "
    End If
    Select Case kind
        Case 1
            txt = arrow & FmtF(Abs(d), 1) & " п.п."
        Case 2
            If prev <> 0# Then
                txt = arrow & Pc(Abs(d) / prev * 100#, 1)
            Else
                txt = arrow & Dash()
            End If
        Case 3
            txt = arrow & Hh(Abs(d))
        Case Else
            txt = arrow & modContentMTO.FmtInt(Abs(d))
    End Select
    Dim cls As String
    If Abs(d) < 0.000000001 Then
        cls = "flat"
    Else
        Dim good As Boolean
        If upGood Then good = (d > 0#) Else good = (d < 0#)
        cls = IIf(good, "up", "dn")
    End If
    DeltaSpan = "<span class=""delta " & cls & """>" & txt & " к пр." & Nb() & "нед.</span>"
End Function

' Плитка KPI со спарклайном и дельтой.
Public Function KpiTileD(ByVal lab As String, ByVal val As String, ByVal cls As String, _
                         ByVal delta As String, ByVal sp As String) As String
    Dim c As String
    If cls <> "" Then c = " " & cls Else c = ""
    KpiTileD = "<div class=""kpi""><div class=""lab"">" & lab & "</div>" & _
        "<div class=""val num" & c & """>" & val & "</div>" & _
        "<div class=""row"">" & delta & sp & "</div></div>"
End Function

' Оценка по проценту: норма / ниже нормы / провал; при малом объёме - «мало данных».
Public Function Grade(ByVal p As Double, ByVal nrec As Double, ByVal hasValue As Boolean) As String
    Dim minRec As Double, norma As Double, proval As Double
    minRec = ToNum(modMain.GetVariableDef("REPORT/MIN_POST_RECORDS", "10"))
    If minRec <= 0# Then minRec = 10#
    norma = ToNum(modMain.GetVariableDef("REPORT/NormaForPlanshet", "90"))
    If norma <= 0# Then norma = 90#
    proval = ToNum(modMain.GetVariableDef("REPORT/ProvalForPlanshet", "50"))
    If proval <= 0# Then proval = 50#

    If nrec < minRec Then
        Grade = "<span class=""pill"">мало данных</span>"
    ElseIf Not hasValue Then
        Grade = Dash()
    ElseIf p >= norma Then
        Grade = "<span class=""pill good"">норма</span>"
    ElseIf p < proval Then
        Grade = "<span class=""pill crit"">провал</span>"
    Else
        Grade = "<span class=""pill warn"">ниже нормы</span>"
    End If
End Function

' Парные столбики двух серий по одной оси категорий.
Public Function GroupBars(ByVal labels As Variant, ByVal s1v As Variant, ByVal s2v As Variant, _
                          ByVal w As Long, ByVal h As Long) As String
    Dim cnt As Long, i As Long, mx As Double
    cnt = UBound(labels) - LBound(labels) + 1
    mx = 0#
    For i = 0 To cnt - 1
        If CDbl(s1v(LBound(s1v) + i)) > mx Then mx = CDbl(s1v(LBound(s1v) + i))
        If CDbl(s2v(LBound(s2v) + i)) > mx Then mx = CDbl(s2v(LBound(s2v) + i))
    Next i
    If mx <= 0# Then mx = 1#

    Dim baseY As Double, topY As Double, stepX As Double, bw As Double
    baseY = h - 44: topY = 24
    stepX = w / cnt
    bw = stepX * 0.26

    Dim s As String
    s = "<svg viewBox=""0 0 " & CStr(w) & " " & CStr(h) & """ role=""img"">"
    s = s & "<line x1=""0"" y1=""" & CStr(CLng(baseY)) & """ x2=""" & CStr(w) & """ y2=""" & _
        CStr(CLng(baseY)) & """ stroke=""var(--line-strong)""/>"
    For i = 0 To cnt - 1
        Dim x0 As Double, j As Long
        x0 = stepX * i + (stepX - (2# * bw + 4#)) / 2#
        Dim vv As Variant, cc As Variant
        vv = Array(CDbl(s1v(LBound(s1v) + i)), CDbl(s2v(LBound(s2v) + i)))
        cc = Array("var(--s2)", "var(--s7)")
        For j = 0 To 1
            Dim bh As Double, x As Double
            bh = (baseY - topY) * CDbl(vv(j)) / mx
            x = x0 + bw * j
            s = s & "<rect x=""" & Cx(x) & """ y=""" & Cx(baseY - bh) & """ width=""" & _
                Cx(bw) & """ height=""" & Cx(bh) & """ fill=""" & CStr(cc(j)) & """/>"
            s = s & "<text x=""" & Cx(x + bw / 2#) & """ y=""" & Cx(baseY - bh - 5#) & _
                """ text-anchor=""middle"" font-size=""11"" " & _
                "font-family=""IBM Plex Mono,monospace"" fill=""var(--ink-2)"">" & _
                modContentMTO.FmtInt(CDbl(vv(j))) & "</text>"
        Next j
        s = s & "<text x=""" & Cx(stepX * i + stepX / 2#) & """ y=""" & Cx(baseY + 18#) & _
            """ text-anchor=""middle"" font-size=""11.5"" " & _
            "font-family=""IBM Plex Sans,sans-serif"" fill=""var(--muted)"">" & _
            modContentMTO.Esc(CStr(labels(LBound(labels) + i))) & "</text>"
    Next i
    GroupBars = s & "</svg>"
End Function

' =====================================================================================
' СЛАЙД 1. Обзор недели. Единица счёта - заказ-наряд (трек Б).
' =====================================================================================
' Конец недели (исключающая граница) для накопительных метрик.
Private Function WeekEnd(ByVal yw As Long) As Double
    WeekEnd = CDbl(WeekMonday(yw)) + 7#
End Function

' Восемь рядов слайда 1 за окно недель. Всё считается из даты создания наряда
' и даты закрытия: колонка yearWeek из Power Query здесь не участвует.
Private Sub Slide1Series(ByRef wk As Variant, ByRef opened() As Double, _
                         ByRef closedA() As Double, ByRef visitsA() As Double, _
                         ByRef hangA() As Double, ByRef medA() As Double, _
                         ByRef tail14 As Variant, ByRef partsA() As Double, _
                         ByRef tabPct() As Double)
    EnsureVeh
    wk = WeekWindow(8)
    Dim n As Long, i As Long
    n = UBound(wk) + 1
    ReDim opened(0 To n - 1)
    ReDim closedA(0 To n - 1)
    ReDim visitsA(0 To n - 1)
    ReDim hangA(0 To n - 1)
    ReDim medA(0 To n - 1)
    ReDim partsA(0 To n - 1)
    ReDim tabPct(0 To n - 1)
    Dim t14() As Double
    ReDim t14(0 To n - 1)

    Dim medCol() As Collection
    ReDim medCol(0 To n - 1)
    For i = 0 To n - 1
        Set medCol(i) = New Collection
        visitsA(i) = DictVal(mVisitWeeks, CStr(wk(i)))
        tabPct(i) = SafePct(DictVal(mSignTab, CStr(wk(i))), DictVal(mSignTot, CStr(wk(i))))
    Next i

    Dim k As Variant, z As Variant
    For Each k In mZn.Keys
        z = mZn(k)
        Dim dt As Double, cl As Double, ac As Double, lv As Double
        dt = CDbl(z(Z_DATE)): cl = CDbl(z(Z_CLOSED))
        ac = ZAcc(z): lv = ZLev(z)
        For i = 0 To n - 1
            Dim we As Double
            we = WeekEnd(CLng(wk(i)))
            If dt > 0# Then
                If CLng(z(Z_WEEK)) = CLng(wk(i)) Then
                    opened(i) = opened(i) + 1#
                    partsA(i) = partsA(i) + CDbl(z(Z_PARTS))
                End If
                If dt < we Then
                    If cl <= 0# Or cl >= we Then
                        hangA(i) = hangA(i) + 1#
                        If dt < we - 14# Then t14(i) = t14(i) + 1#
                    End If
                End If
            End If
            If cl > 0# Then
                If IsoYearWeek(cl) = CLng(wk(i)) Then closedA(i) = closedA(i) + 1#
            End If
            If lv > 0# And ac > 0# And lv >= ac Then
                If IsoYearWeek(lv) = CLng(wk(i)) Then medCol(i).Add (lv - ac) * 24#
            End If
        Next i
    Next k

    Dim hasV As Boolean
    For i = 0 To n - 1
        medA(i) = MedianOf(medCol(i), hasV)
    Next i
    Dim tv() As Variant
    ReDim tv(0 To n - 1)
    For i = 0 To n - 1
        tv(i) = t14(i)
    Next i
    tail14 = tv
End Sub

Private Function ArrOf(ByRef a() As Double) As Variant
    Dim r() As Variant, i As Long
    ReDim r(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        r(i) = a(i)
    Next i
    ArrOf = r
End Function

Public Function BuildKpiOverview() As String
    Dim wk As Variant, tail14 As Variant
    Dim opened() As Double, closedA() As Double, visitsA() As Double, hangA() As Double
    Dim medA() As Double, partsA() As Double, tabPct() As Double
    Slide1Series wk, opened, closedA, visitsA, hangA, medA, tail14, partsA, tabPct

    Dim n As Long
    n = UBound(wk) + 1
    If n < 1 Then BuildKpiOverview = modContentMTO.EmptyNote(): Exit Function
    Dim hasPrev As Boolean
    hasPrev = (n >= 2)
    Dim c As Long, p As Long
    c = n - 1: p = n - 2
    If Not hasPrev Then p = c

    Dim t14 As Variant
    t14 = tail14

    Dim s As String
    s = "<div class=""kpis"">"
    s = s & KpiTileD("Нарядов открыто", modContentMTO.FmtInt(opened(c)), "", _
        DeltaSpan(opened(c), opened(p), 0, False, hasPrev), Spark(ArrOf(opened)))
    s = s & KpiTileD("Закрыто нарядов", modContentMTO.FmtInt(closedA(c)), "", _
        DeltaSpan(closedA(c), closedA(p), 0, True, hasPrev), Spark(ArrOf(closedA)))
    s = s & KpiTileD("Заездов техники", modContentMTO.FmtInt(visitsA(c)), "", _
        DeltaSpan(visitsA(c), visitsA(p), 0, False, hasPrev), Spark(ArrOf(visitsA)))
    s = s & KpiTileD("Висит на конец недели", modContentMTO.FmtInt(hangA(c)), "crit", _
        DeltaSpan(hangA(c), hangA(p), 0, False, hasPrev), Spark(ArrOf(hangA)))
    s = s & KpiTileD("Медиана в ремзоне", Hh(medA(c)), "", _
        DeltaSpan(medA(c), medA(p), 3, False, hasPrev), Spark(ArrOf(medA)))
    s = s & KpiTileD("Висит дольше 14 суток", modContentMTO.FmtInt(CDbl(t14(c))), "crit", _
        DeltaSpan(CDbl(t14(c)), CDbl(t14(p)), 0, False, hasPrev), Spark(t14))
    s = s & KpiTileD("Материалы за неделю", Rub(partsA(c)), "", _
        DeltaSpan(partsA(c), partsA(p), 2, False, hasPrev), Spark(ArrOf(partsA)))
    s = s & KpiTileD("Подписей с планшета", Pc(tabPct(c), 1), "", _
        DeltaSpan(tabPct(c), tabPct(p), 1, True, hasPrev), Spark(ArrOf(tabPct)))
    BuildKpiOverview = s & "</div>"
End Function

' Гистограмма «приёмка -> выбытие». Пара - наряд и дирекция: у каждой дирекции
' своя подпись, поэтому пар вдвое больше, чем нарядов с обеими подписями.
Public Function BuildTimeHist() As String
    EnsureZn
    Dim col As Collection, k As Variant, z As Variant
    Set col = New Collection
    For Each k In mZn.Keys
        z = mZn(k)
        If CDbl(z(Z_ACCG)) > 0# And CDbl(z(Z_LEVG)) >= CDbl(z(Z_ACCG)) Then
            col.Add (CDbl(z(Z_LEVG)) - CDbl(z(Z_ACCG))) * 24#
        End If
        If CDbl(z(Z_ACCD)) > 0# And CDbl(z(Z_LEVD)) >= CDbl(z(Z_ACCD)) Then
            col.Add (CDbl(z(Z_LEVD)) - CDbl(z(Z_ACCD))) * 24#
        End If
    Next k
    If col.Count = 0 Then BuildTimeHist = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, lim As Variant
    labs = Array("до 1 ч", "1" & ChrW$(&H2013) & "4 ч", "4" & ChrW$(&H2013) & "12 ч", _
        "12" & ChrW$(&H2013) & "24 ч", "1" & ChrW$(&H2013) & "3 сут", _
        "3" & ChrW$(&H2013) & "7 сут", "7 сут +")
    lim = Array(1#, 4#, 12#, 24#, 72#, 168#)

    Dim vals(0 To 6) As Double, i As Long, j As Long, sum As Double, mx As Double
    sum = 0#: mx = 0#
    For i = 1 To col.Count
        Dim h As Double, b As Long
        h = CDbl(col(i))
        sum = sum + h
        If h > mx Then mx = h
        b = 6
        For j = 0 To 5
            If h < CDbl(lim(j)) Then
                b = j
                Exit For
            End If
        Next j
        vals(b) = vals(b) + 1#
    Next i
    Dim vv() As Variant
    ReDim vv(0 To 6)
    For i = 0 To 6
        vv(i) = vals(i)
    Next i

    Dim hasV As Boolean, med As Double, p90 As Double, p99 As Double, avg As Double
    med = MedianOf(col, hasV)
    p90 = PctlOf(col, 90#, hasV)
    p99 = PctlOf(col, 99#, hasV)
    avg = sum / col.Count

    Dim ratio As String
    If med > 0# Then ratio = modContentMTO.FmtInt(avg / med) Else ratio = Dash()

    Dim s As String
    s = "<figure>" & HBars(labs, vv, 680, 110, 25)
    s = s & "<figcaption>Медиана " & Hh(med) & ", среднее " & Hh(avg) & ", p90 " & _
        Hh(p90) & ", p99 " & Hh(p99) & ". Всего " & _
        modContentMTO.FmtInt(CDbl(col.Count)) & " пар.</figcaption></figure>"
    s = s & NoteBlk("Хвост длиннее самой метрики: максимум " & ChrW$(&H2014) & " <b>" & _
        FmtF(mx / 24#, 1) & " сут</b>. Ориентир " & ChrW$(&H2014) & " медиана; среднее " & _
        "в " & ratio & " раз выше и для управления непригодно. Интервал меряет всё от " & _
        "приёмки до подписания выбытия, включая очередь и ожидание запчастей: для " & _
        "«плана против факта» он непригоден.")
    BuildTimeHist = s
End Function

Public Function BuildFlowZnType() As String
    EnsureZn
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, t As String, tot As Double
    tot = 0#
    For Each k In mZn.Keys
        t = Trim$(CStr(mZn(k)(Z_TYPE)))
        If t = "" Then t = "(вид не указан)"
        AddCnt d, t, 1#
        tot = tot + 1#
    Next k
    If tot = 0# Then BuildFlowZnType = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant, i As Long
    TopKeys d, 8, labs, vals
    Dim s As String
    s = "<table><thead><tr><th>Вид ремонта</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">%</th></tr></thead><tbody>"
    For i = 0 To UBound(labs)
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(vals(i))) & "</td><td class=""n"">" & _
            FmtF(SafePct(CDbl(vals(i)), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & NoteBlk(modContentMTO.Esc(CStr(labs(0))) & " " & ChrW$(&H2014) & " " & _
        Pc(SafePct(CDbl(vals(0)), tot), 1) & " потока. Планового ТО в ремзоне почти не " & _
        "видно: либо ТО идёт мимо наряда, либо оформляется как «обслуживание при выпуске».")
    BuildFlowZnType = s
End Function

Public Function BuildFlowDefekt() As String
    EnsureZn
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, t As String
    For Each k In mZn.Keys
        t = Trim$(CStr(mZn(k)(Z_DEFEKT)))
        If t = "" Then t = NOSECT
        AddCnt d, t, 1#
    Next k
    If d.Count = 0 Then BuildFlowDefekt = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant
    TopKeys d, 10, labs, vals
    BuildFlowDefekt = HBars(labs, vals, 680, 230, 24) & NoteBlk( _
        "<code>defekt_type</code> " & ChrW$(&H2014) & " готовая группа отказа " & _
        "(двигатель, тормозная система, электрооборудование), " & _
        modContentMTO.FmtInt(CDbl(d.Count)) & " значений на весь массив. Для Парето " & _
        "группировки достаточно, для метрики повторного ремонта " & ChrW$(&H2014) & _
        " слишком крупно (см. слайд 6). Пустое значение " & ChrW$(&H2014) & _
        " норма у планового ТО, поэтому отдельной строкой, а не в прочих.")
End Function

Public Function BuildNoPostWeekly() As String
    EnsureZn
    Dim wk As Variant
    wk = WeekWindow(8)
    Dim n As Long, i As Long
    n = UBound(wk) + 1
    Dim labs() As Variant, bars() As Variant, ln() As Variant
    ReDim labs(0 To n - 1)
    ReDim bars(0 To n - 1)
    ReDim ln(0 To n - 1)

    Dim tot() As Double, np() As Double
    ReDim tot(0 To n - 1)
    ReDim np(0 To n - 1)
    Dim k As Variant, z As Variant
    For Each k In mZn.Keys
        z = mZn(k)
        For i = 0 To n - 1
            If CLng(z(Z_WEEK)) = CLng(wk(i)) Then
                tot(i) = tot(i) + 1#
                If Trim$(CStr(z(Z_POST))) = "" Then np(i) = np(i) + 1#
                Exit For
            End If
        Next i
    Next k

    Dim noSign As Double, noSignNoPost As Double
    noSign = 0#: noSignNoPost = 0#
    For Each k In mZn.Keys
        z = mZn(k)
        If CDbl(z(Z_SIGNED)) = 0# Then
            noSign = noSign + 1#
            If Trim$(CStr(z(Z_POST))) = "" Then noSignNoPost = noSignNoPost + 1#
        End If
    Next k

    For i = 0 To n - 1
        labs(i) = WLab(CLng(wk(i)))
        bars(i) = tot(i)
        ln(i) = SafePct(np(i), tot(i))
    Next i

    Dim s As String
    s = "<figure>" & Collines(labs, bars, ln, 940, 230, 1, True)
    s = s & "<figcaption>Столбики " & ChrW$(&H2014) & " нарядов всего за неделю, линия " & _
        ChrW$(&H2014) & " доля нарядов с пустым <code>post</code>.</figcaption></figure>"
    s = s & "<div class=""legend""><span><i style=""background:var(--s1);opacity:.45""></i>" & _
        "нарядов всего</span><span><i style=""background:var(--s2)""></i>" & _
        "доля без поста, %</span></div>"
    s = s & NoteBlk("«Без поста» и «не подписано» " & ChrW$(&H2014) & _
        " по данным один и тот же дефект: <b>" & _
        Pc(SafePct(noSignNoPost, noSign), 1) & "</b> нарядов без единой подписи имеют " & _
        "пустой <code>post</code>. Отсюда мост к слайду 4.")
    BuildNoPostWeekly = s
End Function

' Диапазон дат недели для подписи: «23–29 мар» либо «30 мар – 5 апр».
Public Function WeekRange(ByVal yw As Long) As String
    Dim a As Date, b As Date, m As Variant
    a = WeekMonday(yw)
    b = a + 6
    m = Array("янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек")
    If Month(a) = Month(b) Then
        WeekRange = CStr(Day(a)) & ChrW$(&H2013) & CStr(Day(b)) & " " & CStr(m(Month(b) - 1))
    Else
        WeekRange = CStr(Day(a)) & " " & CStr(m(Month(a) - 1)) & " " & ChrW$(&H2013) & " " & _
            CStr(Day(b)) & " " & CStr(m(Month(b) - 1))
    End If
End Function

' Подпись отчётной недели: «неделя 2026-13 (23–29 мар)».
Public Function WeekCaption(ByVal yw As Long) As String
    If yw <= 0 Then WeekCaption = "неделя не определена": Exit Function
    WeekCaption = "неделя " & WLab(yw) & " (" & WeekRange(yw) & ")"
End Function

' =====================================================================================
' Числа шапки отчёта. Считаются здесь же, чтобы шапка и слайды не разошлись.
' =====================================================================================
Public Function EventsCount() As Double
    EnsureZn
    EventsCount = CDbl(mRows)
End Function

Public Function OrdersCount() As Double
    EnsureZn
    OrdersCount = CDbl(mZn.Count)
End Function

Public Function FleetCount() As Double
    EnsureVeh
    FleetCount = CDbl(mVeh.Count)
End Function

Public Function NoPostPct() As Double
    EnsureZn
    Dim k As Variant, np As Double
    np = 0#
    For Each k In mZn.Keys
        If Trim$(CStr(mZn(k)(Z_POST))) = "" Then np = np + 1#
    Next k
    NoPostPct = SafePct(np, CDbl(mZn.Count))
End Function

Public Function SnapFrom() As Double
    EnsureZn
    Dim k As Variant, mn As Double
    mn = 0#
    For Each k In mZn.Keys
        Dim d As Double
        d = CDbl(mZn(k)(Z_DATE))
        If d > 0# Then
            If mn = 0# Or d < mn Then mn = d
        End If
    Next k
    SnapFrom = mn
End Function

Public Function SnapTo() As Double
    SnapTo = SnapshotEnd()
End Function
