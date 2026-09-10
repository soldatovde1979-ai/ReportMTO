Attribute VB_Name = "modContentDisc"
' modContentDisc - CONTENT-слой части «Дисциплина» (слайды 2-4) отчёта МТО.
'
' Версия 1.0 от 10.09.2026:
'   - первый выпуск: недельная матрица по ремзонам, площадки за отчётную неделю,
'     таблица по сотрудникам, разбор подписей наряда, слайд 4 «не подписано»;
'   - разметка снята с эталона temp/MTO_макет_отчета_v4.0.html;
'   - примитивы разметки и графики берутся из modContentZone (общие для отчёта).
'
' ЕДИНИЦА СЧЁТА. Здесь их две, и путать их нельзя:
'   - недельная матрица и таблица по сотрудникам считают СОБЫТИЯ подписания;
'   - таблица площадок и разбор подписей считают ЗАКАЗ-НАРЯДЫ.
' У каждого блока единица счёта названа в пояснении под ним.
'
' in_bounds НЕ фильтруется: поле означает «последняя отметка чекина раньше начала
' периода отбора» и к качеству данных отношения не имеет (1C_замечания_по_выгрузке).
'
' ФИО выводятся как есть - это внутренний отчёт. Во внешнюю модель ФИО не уходят:
' псевдонимы и обратная замена живут в modContentMTO.
'
' Файл хранится в UTF-8 + CRLF. Символы вне ANSI-1251 запрещены.
Option Explicit

' Поля записи наряда (mE: number -> Variant-массив)
Private Const E_DATE As Long = 0
Private Const E_POST As Long = 1        ' post - родитель поста, «площадка» трека Б
Private Const E_ZONE As Long = 2        ' postN - нормализованная ремзона трека А
Private Const E_OWNER As Long = 3
Private Const E_TYPE As Long = 4
Private Const E_AG As Long = 5          ' АРМ подписи «Готов к приемке», ДГМ
Private Const E_LG As Long = 6          ' АРМ подписи «Готов к выбытию», ДГМ
Private Const E_AD As Long = 7          ' то же, ДЭНТ
Private Const E_LD As Long = 8
Private Const E_TAG As Long = 9         ' отметка времени приёмки ДГМ
Private Const E_TLG As Long = 10
Private Const E_TAD As Long = 11
Private Const E_TLD As Long = 12
Private Const E_FIELDS As Long = 13

Private Const NOPOST As String = "(пост не указан)"

Private mReady As Boolean
Private mE As Object             ' number -> Variant(E_FIELDS)
Private mZoneTot As Object       ' «дирекция|зона|неделя» -> событий с известным АРМ
Private mZoneTab As Object       ' то же, только планшет
Private mWeekTot As Object       ' «дирекция|неделя» -> событий
Private mWeekTab As Object
Private mPTot As Object          ' «дирекция|ФИО|неделя» -> событий
Private mPTab As Object
Private mPAcc As Object          ' «дирекция|ФИО» за отчётную неделю
Private mPLev As Object
Private mPPairs As Object        ' «дирекция|ФИО» -> Collection номеров нарядов (выбытие)
Private mEvRows As Long
Private mEvUnsigned As Long
Private mRw As Long

Public Sub ResetDisc()
    mReady = False
    Set mE = Nothing
    Set mZoneTot = Nothing
    Set mZoneTab = Nothing
    Set mWeekTot = Nothing
    Set mWeekTab = Nothing
    Set mPTot = Nothing
    Set mPTab = Nothing
    Set mPAcc = Nothing
    Set mPLev = Nothing
    Set mPPairs = Nothing
    mRw = 0
End Sub

' Один проход по снимку: события -> наряды + все счётчики слайдов 2-4.
Private Sub EnsureDisc()
    If mReady Then Exit Sub
    If Not modAggregate.IsReady() Then
        Err.Raise vbObjectError + 42, , "modContentDisc: снимок tbDATA не создан"
    End If

    Set mE = CreateObject("Scripting.Dictionary")
    Set mZoneTot = CreateObject("Scripting.Dictionary")
    Set mZoneTab = CreateObject("Scripting.Dictionary")
    Set mWeekTot = CreateObject("Scripting.Dictionary")
    Set mWeekTab = CreateObject("Scripting.Dictionary")
    Set mPTot = CreateObject("Scripting.Dictionary")
    Set mPTab = CreateObject("Scripting.Dictionary")
    Set mPAcc = CreateObject("Scripting.Dictionary")
    Set mPLev = CreateObject("Scripting.Dictionary")
    Set mPPairs = CreateObject("Scripting.Dictionary")

    mRw = modContentZone.ZoneReportWeek()

    Dim hasZone As Boolean, hasOwner As Boolean, hasDep As Boolean, hasEmp As Boolean
    hasZone = modAggregate.HasColumn("postN")
    hasOwner = modAggregate.HasColumn("owner_dep")
    hasDep = modAggregate.HasColumn("emp_dep")
    hasEmp = modAggregate.HasColumn("employee")

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    mEvRows = n
    mEvUnsigned = 0

    For r = 1 To n
        Dim num As String
        num = modAggregate.CellText(r, "number")
        If Len(num) > 0 Then
            Dim e As Variant
            If mE.Exists(num) Then
                e = mE(num)
            Else
                ReDim e(0 To E_FIELDS - 1)
                e(E_DATE) = modContentZone.ToSerial(modAggregate.CellRaw(r, "date"))
                e(E_POST) = modAggregate.CellText(r, "post")
                If hasZone Then e(E_ZONE) = modAggregate.CellText(r, "postN") Else e(E_ZONE) = ""
                If hasOwner Then e(E_OWNER) = modAggregate.CellText(r, "owner_dep") Else e(E_OWNER) = ""
                e(E_TYPE) = modAggregate.CellText(r, "zn_type")
                e(E_AG) = "": e(E_LG) = "": e(E_AD) = "": e(E_LD) = ""
                e(E_TAG) = 0#: e(E_TLG) = 0#: e(E_TAD) = 0#: e(E_TLD) = 0#
            End If

            Dim rf As String, dr As String, arm As String, sd As Double
            rf = modAggregate.CellText(r, "ready_for")
            dr = modAggregate.CellText(r, "direction")
            arm = modAggregate.CellText(r, "arm")
            sd = modContentZone.ToSerial(modAggregate.CellRaw(r, "status_date"))

            Dim isG As Boolean, isAcc As Boolean, signed As Boolean, tab1 As Boolean
            isG = (InStr(1, dr, "ДГМ", vbTextCompare) > 0)
            ' Статусы в выгрузке пишутся через «е»: «Готов к приемке» / «Готов к выбытию».
            isAcc = (InStr(1, rf, "приемке", vbTextCompare) > 0)
            signed = (arm = "ПК" Or arm = "ПЛАНШЕТ")
            tab1 = (arm = "ПЛАНШЕТ")
            If Not signed Then mEvUnsigned = mEvUnsigned + 1

            Dim fa As Long, ft As Long
            If isG Then
                If isAcc Then
                    fa = E_AG: ft = E_TAG
                Else
                    fa = E_LG: ft = E_TLG
                End If
            Else
                If isAcc Then
                    fa = E_AD: ft = E_TAD
                Else
                    fa = E_LD: ft = E_TLD
                End If
            End If
            e(fa) = arm
            e(ft) = sd
            mE(num) = e

            If signed And sd > 0# Then
                Dim wS As String, dk As String, zn As String
                wS = CStr(modContentZone.IsoYearWeek(sd))
                dk = IIf(isG, "ДГМ", "ДЭНТ")
                zn = Trim$(CStr(e(E_ZONE)))
                If zn = "" Then zn = NOPOST
                modContentZone.AddCnt mZoneTot, dk & "|" & zn & "|" & wS, 1#
                modContentZone.AddCnt mWeekTot, dk & "|" & wS, 1#
                If tab1 Then
                    modContentZone.AddCnt mZoneTab, dk & "|" & zn & "|" & wS, 1#
                    modContentZone.AddCnt mWeekTab, dk & "|" & wS, 1#
                End If

                Dim emp As String, dep As String
                If hasEmp Then emp = Trim$(modAggregate.CellText(r, "employee")) Else emp = ""
                If hasDep Then dep = Trim$(modAggregate.CellText(r, "emp_dep")) Else dep = ""
                If Len(emp) > 0 And Len(dep) > 0 Then
                    modContentZone.AddCnt mPTot, dep & "|" & emp & "|" & wS, 1#
                    If tab1 Then modContentZone.AddCnt mPTab, dep & "|" & emp & "|" & wS, 1#
                    If CLng(wS) = mRw Then
                        If isAcc Then
                            modContentZone.AddCnt mPAcc, dep & "|" & emp, 1#
                        Else
                            modContentZone.AddCnt mPLev, dep & "|" & emp, 1#
                            Dim pk As String
                            pk = dep & "|" & emp
                            If Not mPPairs.Exists(pk) Then mPPairs.Add pk, New Collection
                            mPPairs(pk).Add num & Chr$(1) & IIf(isG, "G", "D")
                        End If
                    End If
                End If
            End If
        End If
    Next r

    mReady = True
    modLog.WriteDebug 2, "Дисциплина", "modContentDisc.EnsureDisc", _
        "Событий " & CStr(n) & " -> нарядов " & CStr(mE.Count) & _
        ", отчётная неделя " & CStr(mRw)
End Sub

' =====================================================================================
' СЛАЙДЫ 2 и 3. Дирекция: недели, площадки, люди, разбор подписей.
' =====================================================================================
Private Function ZoneOf(ByVal e As Variant) As String
    Dim z As String
    z = Trim$(CStr(e(E_ZONE)))
    If z = "" Then z = NOPOST
    ZoneOf = z
End Function

Private Function ArmField(ByVal dir As String, ByVal acc As Boolean) As Long
    If dir = "ДГМ" Then
        ArmField = IIf(acc, E_AG, E_LG)
    Else
        ArmField = IIf(acc, E_AD, E_LD)
    End If
End Function

Private Function TimeField(ByVal dir As String, ByVal acc As Boolean) As Long
    If dir = "ДГМ" Then
        TimeField = IIf(acc, E_TAG, E_TLG)
    Else
        TimeField = IIf(acc, E_TAD, E_TLD)
    End If
End Function

' {{BLOCK_WEEKS_*}} - матрица «% планшета» по ремзонам и неделям окна.
' Единица счёта - СОБЫТИЕ подписания с известным АРМ.
Public Function BuildWeeksTable(ByVal dir As String) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(8)
    Dim n As Long, i As Long
    n = UBound(wk) + 1

    ' Набор ремзон дирекции и их объём за окно.
    Dim zt As Object
    Set zt = CreateObject("Scripting.Dictionary")
    Dim k As Variant, parts As Variant
    For Each k In mZoneTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dir Then
            For i = 0 To n - 1
                If CStr(parts(2)) = CStr(wk(i)) Then
                    modContentZone.AddCnt zt, CStr(parts(1)), modContentZone.DictVal(mZoneTot, CStr(k))
                    Exit For
                End If
            Next i
        End If
    Next k
    If zt.Count = 0 Then BuildWeeksTable = modContentMTO.EmptyNote(): Exit Function

    Dim zl As Variant, zv As Variant
    modContentZone.TopKeys zt, 0, zl, zv

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Ремзона</th>"
    For i = 0 To n - 1
        s = s & "<th class=""n"">" & modContentZone.WLab(CLng(wk(i))) & "</th>"
    Next i
    s = s & "</tr></thead><tbody>"

    s = s & "<tr class=""total""><td class=""head"">Все ремзоны</td>"
    For i = 0 To n - 1
        Dim tw As Double, bw As Double
        tw = modContentZone.DictVal(mWeekTot, dir & "|" & CStr(wk(i)))
        bw = modContentZone.DictVal(mWeekTab, dir & "|" & CStr(wk(i)))
        s = s & modContentZone.PctTd(modContentZone.SafePct(bw, tw), tw > 0#)
    Next i
    s = s & "</tr>"

    Dim j As Long
    For j = 0 To UBound(zl)
        Dim zn As String
        zn = CStr(zl(j))
        If zn = NOPOST Then
            s = s & "<tr><td style=""color:var(--crit)"">" & modContentMTO.Esc(zn) & "</td>"
        Else
            s = s & "<tr><td>" & modContentMTO.Esc(zn) & "</td>"
        End If
        For i = 0 To n - 1
            Dim key As String, t2 As Double, b2 As Double
            key = dir & "|" & zn & "|" & CStr(wk(i))
            t2 = modContentZone.DictVal(mZoneTot, key)
            b2 = modContentZone.DictVal(mZoneTab, key)
            s = s & modContentZone.PctTd(modContentZone.SafePct(b2, t2), t2 > 0#)
        Next i
        s = s & "</tr>"
    Next j

    s = s & "<tr><td class=""head"">Событий всего</td>"
    For i = 0 To n - 1
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mWeekTot, dir & "|" & CStr(wk(i)))) & "</td>"
    Next i
    s = s & "</tr></tbody></table></div>"

    s = s & modContentZone.NoteBlk("В ячейках " & ChrW$(&H2014) & " <b>% планшета</b> от " & _
        "подписаний с известным АРМ (<code>arm " & ChrW$(&H2208) & _
        " {ПК, ПЛАНШЕТ}</code>). Цвет уходит в рамку, а не в заливку. Строка «Событий " & _
        "всего» " & ChrW$(&H2014) & " знаменатель, чтобы процент нельзя было читать в " & _
        "отрыве от объёма.")
    BuildWeeksTable = s
End Function

' {{BLOCK_POSTS_*}} - площадки за отчётную неделю. Единица счёта - ЗАКАЗ-НАРЯД:
' наряд считается «с планшета», только если ВСЕ подписи дирекции по нему за неделю
' поставлены с планшета. Смешивать с соседними блоками (там события) нельзя.
Public Function BuildPostsTable(ByVal dir As String) As String
    EnsureDisc
    Dim zn As Object, zt As Object
    Set zn = CreateObject("Scripting.Dictionary")
    Set zt = CreateObject("Scripting.Dictionary")

    Dim k As Variant, e As Variant
    For Each k In mE.Keys
        e = mE(k)
        Dim cntS As Long, cntT As Long, p As Long
        cntS = 0: cntT = 0
        For p = 0 To 1
            Dim a As String, t As Double
            a = CStr(e(ArmField(dir, p = 0)))
            t = CDbl(e(TimeField(dir, p = 0)))
            If (a = "ПК" Or a = "ПЛАНШЕТ") And t > 0# Then
                If modContentZone.IsoYearWeek(t) = mRw Then
                    cntS = cntS + 1
                    If a = "ПЛАНШЕТ" Then cntT = cntT + 1
                End If
            End If
        Next p
        If cntS > 0 Then
            Dim z As String
            z = ZoneOf(e)
            modContentZone.AddCnt zn, z, 1#
            If cntT = cntS Then modContentZone.AddCnt zt, z, 1#
        End If
    Next k
    If zn.Count = 0 Then BuildPostsTable = modContentMTO.EmptyNote(): Exit Function

    Dim zl As Variant, zv As Variant
    modContentZone.TopKeys zn, 0, zl, zv

    Dim s As String, i As Long
    s = "<table><thead><tr><th>Площадка</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">С планшета</th><th class=""n"">%</th><th>Оценка</th>" & _
        "</tr></thead><tbody>"
    For i = 0 To UBound(zl)
        Dim tot As Double, tab1 As Double, pct As Double
        tot = CDbl(zv(i))
        tab1 = modContentZone.DictVal(zt, CStr(zl(i)))
        pct = modContentZone.SafePct(tab1, tot)
        If CStr(zl(i)) = NOPOST Then
            s = s & "<tr><td style=""color:var(--crit)"">" & modContentMTO.Esc(CStr(zl(i))) & "</td>"
        Else
            s = s & "<tr><td>" & modContentMTO.Esc(CStr(zl(i))) & "</td>"
        End If
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(tot) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(tab1) & "</td>"
        s = s & modContentZone.PctTd(pct, tot > 0#)
        s = s & "<td>" & modContentZone.Grade(pct, tot, tot > 0#) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & modContentZone.NoteBlk("<b>Здесь единица счёта " & ChrW$(&H2014) & _
        " наряд</b>, а не событие: соседние блоки считают события, путать нельзя. " & _
        "<code>post</code> " & ChrW$(&H2014) & " родитель поста, поэтому строки " & _
        "читаются как площадки. «Пост не указан» " & ChrW$(&H2014) & " всегда отдельной " & _
        "строкой, норма " & ChrW$(&H2265) & " 90 %, провал &lt; 50 %.")
    BuildPostsTable = s
End Function

' {{BLOCK_PEOPLE_*}} - сотрудники дирекции. Слева тренд процента за четыре недели,
' справа объём и время ЗА ОТЧЁТНУЮ НЕДЕЛЮ: тянуть пять колонок на четыре недели
' незачем, таблица от этого становится нечитаемой.
Public Function BuildPeople(ByVal dir As String) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(4)
    Dim minRec As Double
    minRec = modContentZone.ToNum(modMain.GetVariableDef("REPORT/MIN_RECORDS", "10"))
    If minRec <= 0# Then minRec = 10#

    ' Отбор: не меньше minRec событий за отчётную неделю.
    Dim pick As Object
    Set pick = CreateObject("Scripting.Dictionary")
    Dim k As Variant, parts As Variant
    For Each k In mPTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dir And CLng(parts(2)) = mRw Then
            If modContentZone.DictVal(mPTot, CStr(k)) >= minRec Then
                pick(CStr(parts(1))) = modContentZone.SafePct( _
                    modContentZone.DictVal(mPTab, CStr(k)), _
                    modContentZone.DictVal(mPTot, CStr(k)))
            End If
        End If
    Next k
    If pick.Count = 0 Then BuildPeople = modContentMTO.EmptyNote(): Exit Function

    Dim pl As Variant, pv As Variant
    modContentZone.TopKeys pick, 0, pl, pv

    Dim s As String, i As Long
    s = "<table><thead><tr><th rowspan=""2"">Сотрудник</th>" & _
        "<th class=""grp"" colspan=""4"">% планшета по неделям</th>" & _
        "<th class=""grp sep-l"" colspan=""5"">Отчётная неделя " & _
        modContentZone.WLab(mRw) & " (" & modContentZone.WeekRange(mRw) & ")</th></tr><tr>"
    For i = 0 To 3
        Dim cap As String
        If i = 3 Then cap = "ПН" Else cap = "ПН-" & CStr(3 - i)
        s = s & "<th class=""n"">" & cap & " " & ChrW$(&HB7) & " " & _
            modContentZone.WLab(CLng(wk(i))) & "</th>"
    Next i
    s = s & "<th class=""n sep-l"">Всего подписей</th><th class=""n"">Из них планшет</th>" & _
        "<th class=""n"">Приёмка</th><th class=""n"">Выбытие</th>" & _
        "<th class=""n"">Ср. время</th></tr></thead><tbody>"

    For i = 0 To UBound(pl)
        Dim emp As String
        emp = CStr(pl(i))
        s = s & "<tr><td class=""head"">" & modContentMTO.Esc(emp) & "</td>"
        Dim j As Long
        For j = 0 To 3
            Dim key As String, t As Double, b As Double
            key = dir & "|" & emp & "|" & CStr(wk(j))
            t = modContentZone.DictVal(mPTot, key)
            b = modContentZone.DictVal(mPTab, key)
            s = s & modContentZone.PctTd(modContentZone.SafePct(b, t), t > 0#)
        Next j
        Dim rk As String, tot As Double, tab1 As Double
        rk = dir & "|" & emp & "|" & CStr(mRw)
        tot = modContentZone.DictVal(mPTot, rk)
        tab1 = modContentZone.DictVal(mPTab, rk)
        s = s & "<td class=""n sep-l"">" & modContentMTO.FmtInt(tot) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(tab1) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mPAcc, dir & "|" & emp)) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mPLev, dir & "|" & emp)) & "</td>"
        s = s & "<td class=""n"">" & AvgSpan(dir & "|" & emp) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"

    s = s & modContentZone.NoteBlk("Процент показан за четыре недели, чтобы был виден " & _
        "тренд по человеку: ПН-3 " & ChrW$(&H2026) & " ПН. Справа " & ChrW$(&H2014) & _
        " объём и время за отчётную неделю; тянуть эти пять колонок на четыре недели " & _
        "незачем, таблица от этого становится нечитаемой. ФИО выводятся как есть " & _
        ChrW$(&H2014) & " это внутренний отчёт; во внешнюю модель ФИО <b>не уходят</b>, " & _
        "в промпт подставляются псевдонимы «Сотрудник N», обратная замена делается уже " & _
        "над готовым текстом. Порог включения " & ChrW$(&H2014) & " не менее " & _
        modContentMTO.FmtInt(minRec) & " событий за отчётную неделю; в списке " & _
        modContentMTO.FmtInt(CDbl(pick.Count)) & " человек.")
    BuildPeople = s
End Function

' Среднее «приёмка -> выбытие» по нарядам, где сотрудник подписал выбытие.
Private Function AvgSpan(ByVal pk As String) As String
    AvgSpan = ChrW$(&H2014)
    If Not mPPairs.Exists(pk) Then Exit Function
    Dim col As Collection, i As Long, sum As Double, cnt As Long
    Set col = mPPairs(pk)
    sum = 0#: cnt = 0
    For i = 1 To col.Count
        Dim pr As Variant, e As Variant, a As Double, l As Double
        pr = Split(CStr(col(i)), Chr$(1))
        If mE.Exists(CStr(pr(0))) Then
            e = mE(CStr(pr(0)))
            If CStr(pr(1)) = "G" Then
                a = CDbl(e(E_TAG)): l = CDbl(e(E_TLG))
            Else
                a = CDbl(e(E_TAD)): l = CDbl(e(E_TLD))
            End If
            If a > 0# And l >= a Then
                sum = sum + (l - a) * 24#
                cnt = cnt + 1
            End If
        End If
    Next i
    If cnt > 0 Then AvgSpan = modContentZone.Hh(sum / cnt)
End Function

' {{BLOCK_SIGNSTAT_*}} - разбор всех нарядов дирекции по подписанным статусам.
' Единица счёта - ЗАКАЗ-НАРЯД.
Public Function BuildSignStat(ByVal dir As String) As String
    EnsureDisc
    Dim tot As Double, full As Double, onlyA As Double, onlyL As Double, none As Double
    Dim bothT As Double, mixT As Double, bothP As Double
    Dim aUn(0 To 3) As Double, lUn(0 To 3) As Double

    Dim k As Variant, e As Variant
    For Each k In mE.Keys
        e = mE(k)
        tot = tot + 1#
        Dim aa As String, ll As String, accS As Boolean, levS As Boolean
        aa = CStr(e(ArmField(dir, True)))
        ll = CStr(e(ArmField(dir, False)))
        accS = (aa = "ПК" Or aa = "ПЛАНШЕТ")
        levS = (ll = "ПК" Or ll = "ПЛАНШЕТ")

        If accS And levS Then
            full = full + 1#
            Dim nt As Long
            nt = 0
            If aa = "ПЛАНШЕТ" Then nt = nt + 1
            If ll = "ПЛАНШЕТ" Then nt = nt + 1
            If nt = 2 Then
                bothT = bothT + 1#
            ElseIf nt = 1 Then
                mixT = mixT + 1#
            Else
                bothP = bothP + 1#
            End If
        ElseIf accS Then
            onlyA = onlyA + 1#
        ElseIf levS Then
            onlyL = onlyL + 1#
        Else
            none = none + 1#
        End If

        Dim b As Long
        b = AgeBucket4(CDbl(e(E_DATE)))
        If b >= 0 Then
            If Not accS Then aUn(b) = aUn(b) + 1#
            If Not levS Then lUn(b) = lUn(b) + 1#
        End If
    Next k
    If tot = 0# Then BuildSignStat = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<div class=""two-col wide-l""><div><table><thead><tr>" & _
        "<th>Заказ-наряды дирекции</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">Доля</th></tr></thead><tbody>"
    s = s & "<tr class=""total""><td class=""head"">Всего заказ-нарядов</td><td class=""n"">" & _
        modContentMTO.FmtInt(tot) & "</td><td class=""n"">" & _
        modContentZone.Pc(100#, 1) & "</td></tr>"
    s = s & SsRow("Из них подписаны полностью", full, tot, False)
    s = s & SsRow("оба статуса с планшета", bothT, full, True)
    s = s & SsRow("один с планшета, другой с ПК", mixT, full, True)
    s = s & SsRow("оба статуса с ПК", bothP, full, True)
    s = s & SsRow("Подписана только «Готов к приемке»", onlyA, tot, False)
    s = s & SsRow("Подписана только «Готов к выбытию»", onlyL, tot, False)
    s = s & "<tr><td><b style=""color:var(--crit)"">Не подписан ни один статус</b></td>" & _
        "<td class=""n"">" & modContentMTO.FmtInt(none) & "</td><td class=""n"">" & _
        modContentZone.Pc(modContentZone.SafePct(none, tot), 1) & "</td></tr>"
    s = s & "</tbody></table></div><div>"
    s = s & modContentZone.MockLabel("Неподписанные статусы по возрасту наряда")

    Dim labs As Variant, av() As Variant, lv() As Variant, i As Long
    labs = Array("до суток", "сутки " & ChrW$(&H2013) & " неделя", _
        "неделя " & ChrW$(&H2013) & " месяц", "больше месяца")
    ReDim av(0 To 3)
    ReDim lv(0 To 3)
    For i = 0 To 3
        av(i) = aUn(i)
        lv(i) = lUn(i)
    Next i
    s = s & modContentZone.GroupBars(labs, av, lv, 620, 250)
    s = s & "<div class=""legend""><span><i style=""background:var(--s2)""></i>" & _
        "не подписан «Готов к приемке»</span><span><i style=""background:var(--s7)""></i>" & _
        "не подписан «Готов к выбытию»</span></div></div></div>"

    s = s & modContentZone.NoteBlk("Четыре строки вверху " & ChrW$(&H2014) & " разбор " & _
        "всех нарядов дирекции по тому, какие статусы подписаны. <b>Эти четыре числа у " & _
        "ДЭНТ и ДГМ совпадают до единицы</b>, и это не ошибка: признак «подписано» " & _
        "берётся из даты статуса, а она в заказ-наряде одна на обе дирекции. Дирекции " & _
        "различаются только тем, <b>чем</b> подписывали (" & dir & ": " & _
        modContentZone.Pc(modContentZone.SafePct(bothT, full), 1) & " нарядов целиком с " & _
        "планшета) и кто подписывал. График справа " & ChrW$(&H2014) & " те же наряды в " & _
        "разрезе возраста: чем правее столбик, тем безнадёжнее, подпись задним числом " & _
        "за прошлый месяц уже не поставить.")
    BuildSignStat = s
End Function

Private Function SsRow(ByVal lab As String, ByVal v As Double, ByVal den As Double, _
                       ByVal indent As Boolean) As String
    Dim cell As String
    If indent Then
        cell = "<td>" & modContentZone.Indent() & modContentMTO.Esc(lab) & "</td>"
    Else
        cell = "<td>" & modContentMTO.Esc(lab) & "</td>"
    End If
    SsRow = "<tr>" & cell & "<td class=""n"">" & modContentMTO.FmtInt(v) & _
        "</td><td class=""n"">" & modContentZone.Pc(modContentZone.SafePct(v, den), 1) & _
        "</td></tr>"
End Function

' Возраст наряда, четыре корзины: сутки / неделя / месяц / больше.
Private Function AgeBucket4(ByVal dser As Double) As Long
    AgeBucket4 = -1
    If dser <= 0# Then Exit Function
    Dim a As Double
    a = modContentZone.SnapshotEnd() - dser
    If a < 1# Then
        AgeBucket4 = 0
    ElseIf a < 7# Then
        AgeBucket4 = 1
    ElseIf a < 30# Then
        AgeBucket4 = 2
    Else
        AgeBucket4 = 3
    End If
End Function

' =====================================================================================
' СЛАЙД 4. Не подписано.
' =====================================================================================
Private Function SignedCount(ByVal e As Variant) As Long
    Dim c As Long, f As Variant, i As Long
    f = Array(E_AG, E_LG, E_AD, E_LD)
    c = 0
    For i = 0 To 3
        Dim a As String
        a = CStr(e(CLng(f(i))))
        If a = "ПК" Or a = "ПЛАНШЕТ" Then c = c + 1
    Next i
    SignedCount = c
End Function

Public Function BuildKpiUnsigned() As String
    EnsureDisc
    Dim k As Variant, e As Variant, tot As Double, none As Double, part As Double
    Dim oldest As Double
    tot = 0#: none = 0#: part = 0#: oldest = 0#
    For Each k In mE.Keys
        e = mE(k)
        tot = tot + 1#
        Dim c As Long
        c = SignedCount(e)
        If c = 0 Then
            none = none + 1#
            If CDbl(e(E_DATE)) > 0# Then
                Dim a As Double
                a = modContentZone.SnapshotEnd() - CDbl(e(E_DATE))
                If a > oldest Then oldest = a
            End If
        ElseIf c < 4 Then
            part = part + 1#
        End If
    Next k
    If tot = 0# Then BuildKpiUnsigned = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<div class=""kpis"">"
    s = s & modContentZone.KpiTile("Нарядов без единой подписи", _
        modContentMTO.FmtInt(none), "crit", _
        modContentZone.Pc(modContentZone.SafePct(none, tot), 1) & " от всех нарядов")
    s = s & modContentZone.KpiTile("Событий «НЕ ПОДПИСАНО»", _
        modContentMTO.FmtInt(CDbl(mEvUnsigned)), "crit", _
        modContentZone.Pc(modContentZone.SafePct(CDbl(mEvUnsigned), CDbl(mEvRows)), 1) & _
        " от всех событий")
    s = s & modContentZone.KpiTile("Подписан частично", modContentMTO.FmtInt(part), "", _
        "часть из 4 строк наряда")
    s = s & modContentZone.KpiTile("Самый старый", modContentMTO.FmtInt(Int(oldest)) & _
        " <small>сут</small>", "crit", "от даты создания наряда")
    BuildKpiUnsigned = s & "</div>"
End Function

' Наряды без единой подписи: счётчик по значению поля.
Private Function UnsignedBy(ByVal fld As Long, ByVal emptyLab As String) As Object
    EnsureDisc
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, e As Variant
    For Each k In mE.Keys
        e = mE(k)
        If SignedCount(e) = 0 Then
            Dim v As String
            v = Trim$(CStr(e(fld)))
            If v = "" Then v = emptyLab
            modContentZone.AddCnt d, v, 1#
        End If
    Next k
    Set UnsignedBy = d
End Function

Public Function BuildUnsignedAge() As String
    EnsureDisc
    Dim labs As Variant
    labs = Array("0" & ChrW$(&H2013) & "3 сут", "3" & ChrW$(&H2013) & "7 сут", _
        "7" & ChrW$(&H2013) & "14 сут", "14" & ChrW$(&H2013) & "30 сут", "30 сут +")
    Dim vals(0 To 4) As Double
    Dim k As Variant, e As Variant
    For Each k In mE.Keys
        e = mE(k)
        If SignedCount(e) = 0 And CDbl(e(E_DATE)) > 0# Then
            Dim a As Double
            a = modContentZone.SnapshotEnd() - CDbl(e(E_DATE))
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
        End If
    Next k
    Dim vv() As Variant, i As Long
    ReDim vv(0 To 4)
    For i = 0 To 4
        vv(i) = vals(i)
    Next i

    Dim m As Variant
    m = Array("янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек")
    Dim se As Date
    se = CDate(modContentZone.SnapshotEnd())

    BuildUnsignedAge = modContentZone.Cols(labs, vv, 680, 200) & modContentZone.NoteBlk( _
        "Возраст считается до конца выгрузки, а не до сегодня: снимок обрезан " & _
        CStr(Day(se)) & " " & CStr(m(Month(se) - 1)) & " " & CStr(Year(se)) & _
        ". Старше 30 суток " & ChrW$(&H2014) & " " & modContentMTO.FmtInt(vals(4)) & _
        " нарядов; это уже не «не успели», а брошенные.")
End Function

Public Function BuildUnsignedPost() As String
    Dim d As Object
    Set d = UnsignedBy(E_POST, NOPOST)
    If d.Count = 0 Then BuildUnsignedPost = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedPost = modContentZone.HBars(labs, vals, 660, 230, 24) & _
        modContentZone.NoteBlk("Почти вся масса " & ChrW$(&H2014) & _
        " строки с пустым <code>post</code>. Это тот же дефект, что на слайде 1: наряд " & _
        "не привязан к площадке и не попадает ни в один разрез.")
End Function

Public Function BuildUnsignedOwner() As String
    Dim d As Object
    Set d = UnsignedBy(E_OWNER, "(владелец не указан)")
    If d.Count = 0 Then BuildUnsignedOwner = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedOwner = modContentZone.HBars(labs, vals, 660, 280, 24) & _
        modContentZone.NoteBlk("Разрез по <code>owner_dep</code> отвечает на вопрос " & _
        "«чью технику не подписывают».")
End Function

Public Function BuildUnsignedZnType() As String
    Dim d As Object
    Set d = UnsignedBy(E_TYPE, "(вид не указан)")
    If d.Count = 0 Then BuildUnsignedZnType = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedZnType = modContentZone.HBars(labs, vals, 680, 230, 24) & _
        modContentZone.NoteBlk("Если бы неподписанные концентрировались в плановом ТО, " & _
        "это была бы техническая особенность оформления. Они распределены как и весь " & _
        "поток " & ChrW$(&H2014) & " значит дело в дисциплине.")
End Function

' =====================================================================================
' Точка входа: заполнение плейсхолдеров слайдов 2-4.
' =====================================================================================
Public Sub FillDiscPlaceholders(ByVal d As Object)
    Dim t0 As Single
    t0 = Timer
    EnsureDisc

    d("BLOCK_WEEKS_DENT") = BuildWeeksTable("ДЭНТ")
    d("BLOCK_POSTS_DENT") = BuildPostsTable("ДЭНТ")
    d("BLOCK_PEOPLE_DENT") = BuildPeople("ДЭНТ")
    d("BLOCK_SIGNSTAT_DENT") = BuildSignStat("ДЭНТ")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 2 готов: " & Round(Timer - t0, 2) & " c"

    d("BLOCK_WEEKS_DGM") = BuildWeeksTable("ДГМ")
    d("BLOCK_POSTS_DGM") = BuildPostsTable("ДГМ")
    d("BLOCK_PEOPLE_DGM") = BuildPeople("ДГМ")
    d("BLOCK_SIGNSTAT_DGM") = BuildSignStat("ДГМ")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 3 готов: " & Round(Timer - t0, 2) & " c"

    d("KPI_UNSIGNED") = BuildKpiUnsigned()
    d("BLOCK_UNSIGNED_AGE") = BuildUnsignedAge()
    d("BLOCK_UNSIGNED_POST") = BuildUnsignedPost()
    d("BLOCK_UNSIGNED_OWNER") = BuildUnsignedOwner()
    d("BLOCK_UNSIGNED_ZNTYPE") = BuildUnsignedZnType()
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 4 готов: " & Round(Timer - t0, 2) & " c"
End Sub
