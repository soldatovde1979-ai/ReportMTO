Attribute VB_Name = "modRollup"
' modRollup - CORE, движок понедельной свёртки истории tbDATA.
'
' Версия 1.0 от 29.09.2026. Постановка и обоснование:
'   docs\plans\tz_svertka_istorii_v1.0.md
'
' ЗАЧЕМ. Ретеншн (DATA/KEEP_WEEKS) удалял старые события безвозвратно: глубина
' отчёта ограничивалась окном хранения, а решение владельца v2.2 требовало
' «хранить всю историю». Свёртка разрешает противоречие: события старше окна
' хранения не удаляются молча, а сначала сворачиваются в таблицу tbARCHIVE -
' одна строка на неделю и набор аналитик. Сырые строки уходят, счётчики остаются.
'
' КАК УСТРОЕНО (порядок важен):
'   1. RollupBeforeLoad вызывается из modMain ДО Refresh Power Query.
'   2. Отсечка считается ровно так же, как ретеншн в Query-ImportJSON (R-1):
'      понедельник недели последней даты tbDATA минус (KEEP_WEEKS - 1) недель;
'      опорная дата строки - max(status_date, date).
'   3. В архив уходят строки с опорной датой в [водяной знак; отсечка).
'      Водяной знак - ключ DATA/ROLLUP_UNTIL на листе Variable: до этой даты
'      история уже свёрнута. После записи архива знак сдвигается на отсечку.
'   4. Power Query удаляет только строки старше водяного знака (qRollupUntil):
'      ни одна строка не удаляется, пока не попала в архив.
'
' ИДЕМПОТЕНТНОСТЬ. Повторный запуск без Refresh ничего не добавит: окно
' [знак; отсечка) пусто, потому что знак уже равен отсечке. Упал Refresh после
' свёртки - строки старше знака остались в tbDATA, но в следующую свёртку не
' попадут (окно начинается от знака), а следующий успешный Refresh их удалит.
' Ошибка на любом шаге свёртки - знак НЕ сдвигается, Power Query ничего лишнего
' не удаляет: худший исход - таблица растёт дольше, чем хотелось.
'
' ВЫКЛЮЧЕНИЕ. DATA/ROLLUP = 0 - свёртка не выполняется. Пока на листе есть
' ключ DATA/ROLLUP_UNTIL, Power Query удаляет только свёрнутое, то есть при
' выключенной свёртке ретеншн тоже стоит. Вернуть прежнее «просто удалять» -
' удалить строку DATA/ROLLUP_UNTIL с листа Variable.
'
' Файл хранится в UTF-8 + CRLF. Символы вне ANSI-1251 запрещены.
Option Explicit

Private Const ARCH_SHEET As String = "tbARCHIVE"
Private Const ARCH_TABLE As String = "tbARCHIVE"
Private Const KEY_UNTIL As String = "DATA/ROLLUP_UNTIL"
Private Const KEY_ON As String = "DATA/ROLLUP"
Private Const KEY_KEEP As String = "DATA/KEEP_WEEKS"

' Состав tbARCHIVE. Меняется только вместе с ArchiveHeaders и записью строк.
Private Const A_WEEK As Long = 1        ' ISO-неделя опорной даты, год*100+неделя
Private Const A_MONTH As Long = 2       ' месяц опорной даты, год*100+месяц
Private Const A_DIR As Long = 3         ' direction
Private Const A_RF As Long = 4          ' ready_for
Private Const A_ARM As Long = 5         ' arm
Private Const A_ZONE As Long = 6        ' postN
Private Const A_DEP As Long = 7         ' emp_dep
Private Const A_EMP As Long = 8         ' employee
Private Const A_TYPE As Long = 9        ' zn_type
Private Const A_EVENTS As Long = 10     ' событий в группе
Private Const A_ORDERS As Long = 11     ' уникальных нарядов в группе (в пределах одной свёртки)
Private Const A_SIGNED As Long = 12     ' событий с arm из {ПК, ПЛАНШЕТ}
Private Const A_TABLET As Long = 13     ' событий с arm = ПЛАНШЕТ
Private Const A_DSUM As Long = 14       ' сумма deltaHours
Private Const A_DCNT As Long = 15       ' число событий с deltaHours
Private Const A_AT As Long = 16         ' когда свёрнуто
Private Const A_COLS As Long = 16

' Итог последнего запуска - для журнала и для подписи в отчёте.
Private mLastRolled As Long

' =====================================================================================
' Точка входа: вызывается modMain перед Refresh. Возвращает число свёрнутых событий,
' -1 при ошибке. Наружу ошибку не бросает: загрузка данных важнее свёртки.
' =====================================================================================
Public Function RollupBeforeLoad() As Long
    Dim keepW As Long, wm As Double, cutoff As Double, maxDt As Double
    Dim lo As ListObject, arr As Variant, hdr As Variant
    Dim cNum As Long, cDt As Long, cSd As Long, cDir As Long, cRf As Long
    Dim cArm As Long, cZone As Long, cDep As Long, cEmp As Long, cType As Long, cDelta As Long
    Dim n As Long, r As Long, sd As Double, dt As Double, anc As Double, basis As Double
    Dim grp As Object, ords As Object, key As String, g As Variant, arm As String
    Dim nm As String, dv As Variant, rolled As Long, t0 As Single

    On Error GoTo ErrHandler
    RollupBeforeLoad = 0
    mLastRolled = 0
    t0 = Timer

    If Trim$(modMain.GetVariableDef(KEY_ON, "1")) = "0" Then
        modLog.WriteDebug 1, "Свёртка", "RollupBeforeLoad", "Выключена ключом " & KEY_ON & " = 0"
        Exit Function
    End If

    keepW = SafeLong(modMain.GetVariableDef(KEY_KEEP, "0"))
    If keepW < 1 Then
        modLog.WriteDebug 1, "Свёртка", "RollupBeforeLoad", _
            "Ретеншн выключен (" & KEY_KEEP & " < 1) - сворачивать нечего, история хранится целиком"
        Exit Function
    End If

    Set lo = FindTable("tbDATA")
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function

    arr = lo.DataBodyRange.Value2
    hdr = lo.HeaderRowRange.Value2
    cNum = ColOf(hdr, "number")
    cDt = ColOf(hdr, "date")
    cSd = ColOf(hdr, "status_date")
    cDir = ColOf(hdr, "direction")
    cRf = ColOf(hdr, "ready_for")
    cArm = ColOf(hdr, "arm")
    cZone = ColOf(hdr, "postN")
    cDep = ColOf(hdr, "emp_dep")
    cEmp = ColOf(hdr, "employee")
    cType = ColOf(hdr, "zn_type")
    cDelta = ColOf(hdr, "deltaHours")
    If cDt = 0 Or cDir = 0 Or cArm = 0 Then
        modLog.WriteLogEntry Now, "Предупреждение", "Свёртка", "RollupBeforeLoad", _
            "В tbDATA нет обязательных столбцов date/direction/arm - свёртка пропущена"
        Exit Function
    End If
    n = UBound(arr, 1)

    ' Отсечка - тем же правилом, что ретеншн Power Query (R-1).
    maxDt = 0#
    For r = 1 To n
        anc = Anchor(arr, r, cSd, cDt)
        If anc > maxDt Then maxDt = anc
    Next r
    If maxDt <= 0# Then Exit Function
    cutoff = Int(maxDt) - (Weekday(CDate(Int(maxDt)), vbMonday) - 1) - 7# * (keepW - 1)

    wm = Watermark()
    If cutoff <= wm Then
        modLog.WriteDebug 1, "Свёртка", "RollupBeforeLoad", _
            "Новых недель для свёртки нет: отсечка " & Format$(CDate(cutoff), "dd.mm.yyyy") & _
            " не позже водяного знака"
        Exit Function
    End If

    Set grp = CreateObject("Scripting.Dictionary")
    Set ords = CreateObject("Scripting.Dictionary")
    rolled = 0
    For r = 1 To n
        anc = Anchor(arr, r, cSd, cDt)
        If anc > 0# Then
            If Int(anc) >= wm And Int(anc) < cutoff Then
                sd = 0#: dt = 0#
                If cSd > 0 Then sd = modContentZone.ToSerial(arr(r, cSd))
                dt = modContentZone.ToSerial(arr(r, cDt))
                ' Неделя и месяц - по дате подписи, как в недельных осях отчёта;
                ' у неподписанного события подписи нет - по дате создания наряда.
                If sd > 0# Then basis = sd Else basis = dt
                arm = Txt(arr, r, cArm)
                key = CStr(modContentZone.IsoYearWeek(basis)) & "|" & _
                      CStr(modContentZone.YearMonth(basis)) & "|" & _
                      Txt(arr, r, cDir) & "|" & Txt(arr, r, cRf) & "|" & arm & "|" & _
                      Txt(arr, r, cZone) & "|" & Txt(arr, r, cDep) & "|" & _
                      Txt(arr, r, cEmp) & "|" & Txt(arr, r, cType)
                If grp.Exists(key) Then
                    g = grp(key)
                Else
                    g = Array(0#, 0#, 0#, 0#, 0#)
                End If
                g(0) = g(0) + 1#
                If arm = "ПК" Or arm = "ПЛАНШЕТ" Then g(1) = g(1) + 1#
                If arm = "ПЛАНШЕТ" Then g(2) = g(2) + 1#
                If cDelta > 0 Then
                    dv = arr(r, cDelta)
                    If Not IsEmpty(dv) And Not IsError(dv) Then
                        If IsNumeric(dv) Then
                            g(3) = g(3) + CDbl(dv)
                            g(4) = g(4) + 1#
                        End If
                    End If
                End If
                grp(key) = g
                If cNum > 0 Then
                    nm = Txt(arr, r, cNum)
                    If Len(nm) > 0 Then ords(key & "|" & nm) = True
                End If
                rolled = rolled + 1
            End If
        End If
    Next r

    If rolled > 0 Then WriteArchive grp, ords

    ' Знак сдвигается ТОЛЬКО после успешной записи архива. Апостроф - чтобы Excel
    ' не превратил текст ISO в дату: Power Query разбирает текст однозначно.
    modMain.SetVariable KEY_UNTIL, "'" & Format$(CDate(cutoff), "yyyy-mm-dd")

    mLastRolled = rolled
    RollupBeforeLoad = rolled
    modLog.WriteMilestone "Свёртка", "RollupBeforeLoad", _
        "Свёрнуто событий: " & CStr(rolled) & " -> строк архива: " & CStr(grp.Count) & _
        "; период [" & IIf(wm > 0#, Format$(CDate(wm), "dd.mm.yyyy"), "начало") & "; " & _
        Format$(CDate(cutoff), "dd.mm.yyyy") & "); " & Round(Timer - t0, 1) & " c"
    Exit Function

ErrHandler:
    RollupBeforeLoad = -1
    modLog.WriteLogEntry Now, "Ошибка", "Свёртка", "RollupBeforeLoad", _
        "Свёртка не выполнена, водяной знак не сдвинут (удаления старых строк не будет): " & _
        Err.Number & " " & Err.Description
End Function

' Ручной запуск с листа (Alt+F8). Сами строки удалит следующая загрузка данных.
Public Sub RollupNow()
    Dim res As Long
    res = RollupBeforeLoad()
    If res >= 0 Then
        modLog.WriteMilestone "Свёртка", "RollupNow", _
            "Ручная свёртка: " & CStr(res) & " событий. Свёрнутые строки уйдут из tbDATA " & _
            "при следующей загрузке данных."
    End If
End Sub

' Водяной знак: до этой даты (не включая) история свёрнута. 0 - свёртки не было.
Public Function Watermark() As Double
    Watermark = modContentZone.ToSerial(Trim$(modMain.GetVariableDef(KEY_UNTIL, "")))
    If Watermark > 0# Then Watermark = Int(Watermark)
End Function

Public Function LastRolled() As Long
    LastRolled = mLastRolled
End Function

' Подписанные события архива по месяцам для дирекции: tot - с известным АРМ,
' tab - с планшета. Ключ словарей - год*100+месяц текстом. Архива нет - словари
' остаются как были. Возвращает число прочитанных строк архива.
Public Function ArchiveTabletByMonth(ByVal dirName As String, ByVal tot As Object, _
                                     ByVal tab1 As Object) As Long
    Dim lo As ListObject, arr As Variant, r As Long, n As Long, mk As String

    ArchiveTabletByMonth = 0
    On Error GoTo Fail
    Set lo = FindTable(ARCH_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then Exit Function
    arr = lo.DataBodyRange.Value2
    n = UBound(arr, 1)
    For r = 1 To n
        If StrComp(Trim$(CStr(arr(r, A_DIR))), dirName, vbTextCompare) = 0 Then
            If IsNumeric(arr(r, A_MONTH)) And IsNumeric(arr(r, A_SIGNED)) Then
                If CDbl(arr(r, A_SIGNED)) > 0# Then
                    mk = CStr(CLng(arr(r, A_MONTH)))
                    modContentZone.AddCnt tot, mk, CDbl(arr(r, A_SIGNED))
                    modContentZone.AddCnt tab1, mk, modContentZone.ToNum(arr(r, A_TABLET))
                End If
            End If
        End If
        ArchiveTabletByMonth = ArchiveTabletByMonth + 1
    Next r
    Exit Function
Fail:
    modLog.WriteLogEntry Now, "Предупреждение", "Свёртка", "ArchiveTabletByMonth", _
        "Архив не прочитан: " & Err.Description
End Function

' =====================================================================================
' Внутреннее
' =====================================================================================
Private Sub WriteArchive(ByVal grp As Object, ByVal ords As Object)
    Dim lo As ListObject, outA() As Variant, i As Long, k As Variant, parts As Variant
    Dim g As Variant, ordCnt As Object, ok As Variant, pk As String, p As Long
    Dim existing As Long, stamp As Date

    ' Уникальные наряды по группе: ключ ords = «ключ группы|номер».
    Set ordCnt = CreateObject("Scripting.Dictionary")
    For Each ok In ords.Keys
        pk = CStr(ok)
        p = InStrRev(pk, "|")
        modContentZone.AddCnt ordCnt, Left$(pk, p - 1), 1#
    Next ok

    stamp = Now
    ReDim outA(1 To grp.Count, 1 To A_COLS)
    i = 0
    For Each k In grp.Keys
        i = i + 1
        parts = Split(CStr(k), "|")
        g = grp(k)
        outA(i, A_WEEK) = CLng(parts(0))
        outA(i, A_MONTH) = CLng(parts(1))
        outA(i, A_DIR) = CStr(parts(2))
        outA(i, A_RF) = CStr(parts(3))
        outA(i, A_ARM) = CStr(parts(4))
        outA(i, A_ZONE) = CStr(parts(5))
        outA(i, A_DEP) = CStr(parts(6))
        outA(i, A_EMP) = CStr(parts(7))
        outA(i, A_TYPE) = CStr(parts(8))
        outA(i, A_EVENTS) = g(0)
        outA(i, A_ORDERS) = modContentZone.DictVal(ordCnt, CStr(k))
        outA(i, A_SIGNED) = g(1)
        outA(i, A_TABLET) = g(2)
        outA(i, A_DSUM) = g(3)
        outA(i, A_DCNT) = g(4)
        outA(i, A_AT) = stamp
    Next k

    Set lo = EnsureArchiveTable()
    existing = lo.ListRows.Count
    If existing = 1 Then
        If Len(Trim$(CStr(lo.DataBodyRange.Cells(1, 1).Value))) = 0 Then existing = 0
    End If
    ' Таблица без строк: Excel держит пустую строку-заготовку, пишем в неё.
    lo.HeaderRowRange.Cells(1, 1).Offset(existing + 1, 0).Resize(grp.Count, A_COLS).Value = outA
    lo.Resize lo.HeaderRowRange.Resize(1 + existing + grp.Count, A_COLS)
End Sub

Private Function EnsureArchiveTable() As ListObject
    Dim ws As Worksheet, lo As ListObject, hdr As Variant

    Set lo = FindTable(ARCH_TABLE)
    If Not lo Is Nothing Then Set EnsureArchiveTable = lo: Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ARCH_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ARCH_SHEET
    End If
    hdr = Array("week", "month", "direction", "ready_for", "arm", "postN", "emp_dep", _
                "employee", "zn_type", "events", "orders", "signed", "tablet", _
                "deltaSum", "deltaCnt", "rolledAt")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, A_COLS)).Value = hdr
    Set lo = ws.ListObjects.Add(1, ws.Range(ws.Cells(1, 1), ws.Cells(2, A_COLS)), , 1)
    lo.Name = ARCH_TABLE
    modLog.WriteMilestone "Свёртка", "EnsureArchiveTable", "Создан лист и таблица " & ARCH_TABLE
    Set EnsureArchiveTable = lo
End Function

Private Function FindTable(ByVal nm As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nm, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindTable = Nothing
End Function

Private Function ColOf(ByVal hdr As Variant, ByVal nm As String) As Long
    Dim j As Long
    ColOf = 0
    For j = LBound(hdr, 2) To UBound(hdr, 2)
        If StrComp(Trim$(CStr(hdr(1, j))), nm, vbTextCompare) = 0 Then ColOf = j: Exit Function
    Next j
End Function

' Опорная дата строки - как в Query-ImportJSON (R-1): max(status_date, date).
Private Function Anchor(ByVal arr As Variant, ByVal r As Long, ByVal cSd As Long, _
                        ByVal cDt As Long) As Double
    Dim s As Double, d As Double
    s = 0#
    If cSd > 0 Then s = modContentZone.ToSerial(arr(r, cSd))
    d = modContentZone.ToSerial(arr(r, cDt))
    If s > d Then Anchor = s Else Anchor = d
End Function

Private Function Txt(ByVal arr As Variant, ByVal r As Long, ByVal c As Long) As String
    Txt = ""
    If c <= 0 Then Exit Function
    If IsError(arr(r, c)) Then Exit Function
    If IsEmpty(arr(r, c)) Then Exit Function
    Txt = Replace$(Trim$(CStr(arr(r, c))), "|", "/")
End Function

Private Function SafeLong(ByVal s As String) As Long
    SafeLong = 0
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    SafeLong = CLng(Int(CDbl(s)))
End Function
