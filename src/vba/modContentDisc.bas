Attribute VB_Name = "modContentDisc"
' modContentDisc - CONTENT-слой части «Дисциплина» (слайды 2-4) отчёта МТО.
'
' Версия 1.1 от 10.09.2026 Модульная переменная mE переименована в mOrd:
'   VBA не различает регистр, и `mE` - это зарезервированное слово `Me`,
'   объявление не компилировалось. Линтер дополнен до v1.1, чтобы ловить это.
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

' Поля записи наряда (mOrd: number -> Variant-массив)
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
Private Const E_TEK As Long = 13        ' TekStatusPoDoc (таблицы 5/6 «НЕ ПОДПИСАН»)
Private Const E_FIELDS As Long = 14

Private Const NOPOST As String = "(пост не указан)"

Private mReady As Boolean
Private mOrd As Object             ' number -> Variant(E_FIELDS)
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
    Set mOrd = Nothing
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

    Set mOrd = CreateObject("Scripting.Dictionary")
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
    Dim hasTek As Boolean
    hasZone = modAggregate.HasColumn("postN")
    hasOwner = modAggregate.HasColumn("owner_dep")
    hasDep = modAggregate.HasColumn("emp_dep")
    hasEmp = modAggregate.HasColumn("employee")
    hasTek = modAggregate.HasColumn("TekStatusPoDoc")

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    mEvRows = n
    mEvUnsigned = 0

    For r = 1 To n
        Dim num As String
        num = modAggregate.CellText(r, "number")
        If Len(num) > 0 Then
            Dim e As Variant
            If mOrd.Exists(num) Then
                e = mOrd(num)
            Else
                ReDim e(0 To E_FIELDS - 1)
                e(E_DATE) = modContentZone.ToSerial(modAggregate.CellRaw(r, "date"))
                e(E_POST) = modAggregate.CellText(r, "post")
                If hasZone Then e(E_ZONE) = modAggregate.CellText(r, "postN") Else e(E_ZONE) = ""
                If hasOwner Then e(E_OWNER) = modAggregate.CellText(r, "owner_dep") Else e(E_OWNER) = ""
                e(E_TYPE) = modAggregate.CellText(r, "zn_type")
                e(E_AG) = "": e(E_LG) = "": e(E_AD) = "": e(E_LD) = ""
                e(E_TAG) = 0#: e(E_TLG) = 0#: e(E_TAD) = 0#: e(E_TLD) = 0#
                If hasTek Then e(E_TEK) = modAggregate.CellText(r, "TekStatusPoDoc") Else e(E_TEK) = ""
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
            mOrd(num) = e

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
        "Событий " & CStr(n) & " -> нарядов " & CStr(mOrd.Count) & _
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

    s = s & modContentZone.NoteBlk("В ячейках " & ChrW$(&H2014) & " % подписаний с " & _
        "планшета от событий подписания с известным АРМ (<code>arm " & ChrW$(&H2208) & _
        " {ПК, ПЛАНШЕТ}</code>) дирекции <b>" & modContentMTO.Esc(dir) & "</b> за окно 8 " & _
        "недель по дате статуса. Строка «Событий всего» " & ChrW$(&H2014) & _
        " знаменатель процента.")
    BuildWeeksTable = s
End Function

' {{BLOCK_TABLE1_*}} - Таблица 1 (Б5): события подписания с arm ПК/ПЛАНШЕТ по неделям
' окна 8 (неделя даты статуса), ремзоны набора REPORT/SLIDE_ZONES (по умолчанию
' «СТК+ПРК», разделитель «+» нормализуется в «;»). Строки: Всего / ПЛАНШЕТ / ПК /
' % планшет. Подсветка строки «% планшет» - шкала #ef4444 -> #f59e0b -> #10b981.
Public Function BuildBlock1(ByVal sDir As String) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(8)
    Dim n As Long
    n = UBound(wk) + 1

    Dim raw As String, zones As Object, pz As Variant
    raw = modMain.GetVariableDef("REPORT/SLIDE_ZONES", "СТК+ПРК")
    raw = Replace$(Trim$(raw), "+", ";")
    Set zones = CreateObject("Scripting.Dictionary")
    For Each pz In Split(raw, ";")
        If Len(Trim$(CStr(pz))) > 0 Then zones(Trim$(CStr(pz))) = True
    Next pz

    Dim tot As Object, pc As Object, dictTab As Object
    Set tot = CreateObject("Scripting.Dictionary")
    Set pc = CreateObject("Scripting.Dictionary")
    Set dictTab = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To modAggregate.RowCount()
        Dim dr As String
        dr = modAggregate.CellText(r, "direction")
        If IIf(InStr(1, dr, "ДГМ", vbTextCompare) > 0, "ДГМ", "ДЭНТ") = sDir Then
            Dim zn As String
            zn = Trim$(modAggregate.CellText(r, "postN"))
            If zones.Exists(zn) Then
                Dim arm As String, sd As Double
                arm = modAggregate.CellText(r, "arm")
                If arm = "ПК" Or arm = "ПЛАНШЕТ" Then
                    sd = modContentZone.ToSerial(modAggregate.CellRaw(r, "status_date"))
                    If sd > 0# Then
                        Dim wS As String
                        wS = CStr(modContentZone.IsoYearWeek(sd))
                        modContentZone.AddCnt tot, wS, 1#
                        If arm = "ПК" Then
                            modContentZone.AddCnt pc, wS, 1#
                        Else
                            modContentZone.AddCnt dictTab, wS, 1#
                        End If
                    End If
                End If
            End If
        End If
    Next r

    Dim hasAny As Boolean, i As Long
    hasAny = False
    For i = 0 To n - 1
        If modContentZone.DictVal(tot, CStr(wk(i))) > 0# Then hasAny = True
    Next i
    If Not hasAny Then BuildBlock1 = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Дирекция</th><th>АРМ</th>"
    For i = n - 1 To 0 Step -1
        s = s & "<th class=""n"">" & modContentZone.WLab(CLng(wk(i))) & "</th>"
    Next i
    s = s & "<th class=""n"">Всего</th></tr></thead><tbody>"

    Dim sumTot As Double, sumPc As Double, sumTab As Double
    sumTot = 0#: sumPc = 0#: sumTab = 0#
    For i = 0 To n - 1
        sumTot = sumTot + modContentZone.DictVal(tot, CStr(wk(i)))
        sumPc = sumPc + modContentZone.DictVal(pc, CStr(wk(i)))
        sumTab = sumTab + modContentZone.DictVal(dictTab, CStr(wk(i)))
    Next i

    Dim i2 As Long
    For i2 = 0 To 3
        If i2 = 0 Then
            s = s & "<tr><th rowspan=""4"">" & modContentMTO.Esc(sDir) & "</th>"
        Else
            s = s & "<tr>"
        End If
        Select Case i2
            Case 0
                s = s & "<td class=""head"">Всего</td>"
            Case 1
                s = s & "<td>ПЛАНШЕТ</td>"
            Case 2
                s = s & "<td>ПК</td>"
            Case 3
                s = s & "<td>% планшет</td>"
        End Select
        For i = n - 1 To 0 Step -1
            Dim tw As Double, pw As Double, bw As Double
            tw = modContentZone.DictVal(tot, CStr(wk(i)))
            pw = modContentZone.DictVal(pc, CStr(wk(i)))
            bw = modContentZone.DictVal(dictTab, CStr(wk(i)))
            Select Case i2
                Case 0
                    s = s & "<td class=""n"">" & modContentMTO.FmtInt(tw) & "</td>"
                Case 1
                    s = s & "<td class=""n"">" & modContentMTO.FmtInt(bw) & "</td>"
                Case 2
                    s = s & "<td class=""n"">" & modContentMTO.FmtInt(pw) & "</td>"
                Case 3
                    If tw > 0# Then
                        s = s & "<td class=""n"" style=""background:" & PctHeat( _
                            modContentZone.SafePct(bw, tw)) & ";color:#fff"">" & _
                            modContentZone.FmtF(modContentZone.SafePct(bw, tw), 0) & "</td>"
                    Else
                        s = s & "<td class=""n"" style=""color:var(--muted)"">" & _
                            modContentZone.Dash() & "</td>"
                    End If
            End Select
        Next i
        Select Case i2
            Case 0
                s = s & "<td class=""n"">" & modContentMTO.FmtInt(sumTot) & "</td></tr>"
            Case 1
                s = s & "<td class=""n"">" & modContentMTO.FmtInt(sumTab) & "</td></tr>"
            Case 2
                s = s & "<td class=""n"">" & modContentMTO.FmtInt(sumPc) & "</td></tr>"
            Case 3
                If sumTot > 0# Then
                    s = s & "<td class=""n"" style=""background:" & PctHeat( _
                        modContentZone.SafePct(sumTab, sumTot)) & ";color:#fff"">" & _
                        modContentZone.FmtF(modContentZone.SafePct(sumTab, sumTot), 0) & "</td></tr>"
                Else
                    s = s & "<td class=""n"" style=""color:var(--muted)"">" & _
                        modContentZone.Dash() & "</td></tr>"
                End If
        End Select
    Next i2
    s = s & "</tbody></table></div>"
    s = s & modContentZone.NoteBlk("События подписания с <code>arm</code> из {ПК, " & _
        "ПЛАНШЕТ} дирекции <b>" & modContentMTO.Esc(sDir) & "</b>, ремзоны набора " & _
        "<code>REPORT/SLIDE_ZONES</code> (по умолчанию СТК + ПРК). Недели - последние 8 " & _
        "по дате статуса (<code>status_date</code>), от свежей к старой. % планшет = " & _
        "ПЛАНШЕТ / (ПЛАНШЕТ + ПК) за неделю; цвет строки - шкала #ef4444 " & _
        ChrW$(&H2192) & " #f59e0b " & ChrW$(&H2192) & " #10b981. «НЕ ПОДПИСАНО» " & _
        "исключено.")
    BuildBlock1 = s
End Function

' Подсветка строки «% планшет»: 0-50 % - красный к жёлтому, 50-100 % - жёлтый к зелёному.
Private Function PctHeat(ByVal p As Double) As String
    PctHeat = modColor.PercentToColor(p / 100#, "#ef4444", "#f59e0b", "#10b981")
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
    For Each k In mOrd.Keys
        e = mOrd(k)
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
    s = "<table><thead><tr><th>РемЗона</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">С планшета</th><th class=""n"">% планшетов</th>" & _
        "</tr></thead><tbody>"

    Dim allTot As Double, allTab As Double
    allTot = 0#: allTab = 0#
    For i = 0 To UBound(zl)
        allTot = allTot + CDbl(zv(i))
        allTab = allTab + modContentZone.DictVal(zt, CStr(zl(i)))
    Next i
    s = s & "<tr class=""total""><td class=""head"">Всего</td><td class=""n"">" & _
        modContentMTO.FmtInt(allTot) & "</td><td class=""n"">" & _
        modContentMTO.FmtInt(allTab) & "</td>" & _
        modContentZone.PctTd(modContentZone.SafePct(allTab, allTot), allTot > 0#) & "</tr>"
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
        s = s & modContentZone.PctTd(pct, tot > 0#) & "</tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & modContentZone.NoteBlk("Наряды дирекции <b>" & modContentMTO.Esc(dir) & "</b> с " & _
        "подписью на отчётной неделе " & modContentZone.WLab(mRw) & " (неделя по дате " & _
        "статуса). «С планшета» " & ChrW$(&H2014) & " наряд, у которого все подписи " & _
        "дирекции за эту неделю поставлены с планшета. Разрез " & ChrW$(&H2014) & _
        " нормализованная ремзона <code>postN</code>; «пост не указан» " & _
        ChrW$(&H2014) & " отдельной строкой.")
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
    s = "<table><thead><tr><th rowspan=""4"">Сотрудник</th>" & _
        "<th class=""grp"" colspan=""4"">% планшета по неделям</th>" & _
        "<th class=""grp sep-l"" colspan=""5"">Отчётная неделя " & _
        modContentZone.WLab(mRw) & " (" & modContentZone.WeekRange(mRw) & ")</th></tr>"
    ' Шапка недель лесенкой: ПН-1 в первой строке справа, ПН слева и ПН-2 во второй,
    ' ПН-3 в третьей (колонки слева направо: ПН, ПН-1, ПН-2, ПН-3).
    s = s & "<tr><th></th><th class=""n"">ПН-1 " & ChrW$(&HB7) & " " & _
        modContentZone.WLab(CLng(wk(2))) & "</th><th></th><th></th>" & _
        "<th class=""n sep-l"" rowspan=""3"">Всего подписей</th>" & _
        "<th class=""n"" rowspan=""3"">Из них планшет</th>" & _
        "<th class=""n"" rowspan=""3"">Приёмка</th>" & _
        "<th class=""n"" rowspan=""3"">Выбытие</th>" & _
        "<th class=""n"" rowspan=""3"">Ср. время</th></tr>"
    s = s & "<tr><th class=""n"">ПН " & ChrW$(&HB7) & " " & _
        modContentZone.WLab(CLng(wk(3))) & "</th><th></th><th class=""n"">ПН-2 " & _
        ChrW$(&HB7) & " " & modContentZone.WLab(CLng(wk(1))) & "</th><th></th></tr>"
    s = s & "<tr><th></th><th></th><th></th><th class=""n"">ПН-3 " & _
        ChrW$(&HB7) & " " & modContentZone.WLab(CLng(wk(0))) & "</th></tr></thead><tbody>"

    For i = 0 To UBound(pl)
        Dim emp As String
        emp = CStr(pl(i))
        s = s & "<tr><td class=""head"">" & modContentMTO.Esc(emp) & "</td>"
        Dim j As Long
        Dim pj(0 To 3) As Double, hj(0 To 3) As Boolean
        For j = 0 To 3
            Dim key As String, t As Double, b As Double
            key = dir & "|" & emp & "|" & CStr(wk(j))
            t = modContentZone.DictVal(mPTot, key)
            b = modContentZone.DictVal(mPTab, key)
            pj(j) = modContentZone.SafePct(b, t)
            hj(j) = (t > 0#)
        Next j
        ' Колонки слева направо: ПН (без стрелки), ПН-1, ПН-2, ПН-3.
        s = s & modContentZone.PctTd(pj(3), hj(3))
        s = s & PctArrowTd(pj(2), hj(2), pj(1), hj(1))
        s = s & PctArrowTd(pj(1), hj(1), pj(0), hj(0))
        Dim keyP As String, tPrev As Double
        keyP = dir & "|" & emp & "|" & CStr(modContentZone.PrevWeek(CLng(wk(0))))
        tPrev = modContentZone.DictVal(mPTot, keyP)
        s = s & PctArrowTd(pj(0), hj(0), _
            modContentZone.SafePct(modContentZone.DictVal(mPTab, keyP), tPrev), tPrev > 0#)
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

    s = s & modContentZone.NoteBlk("% планшета за четыре недели (ПН-3 " & _
        ChrW$(&H2026) & " ПН) по дате статуса; стрелка " & ChrW$(&H2191) & " / " & _
        ChrW$(&H2193) & " - рост / падение процента к предыдущей неделе сотрудника. " & _
        "Справа - объём и время за отчётную неделю " & modContentZone.WLab(mRw) & ". " & _
        "Порог включения - не менее " & modContentMTO.FmtInt(minRec) & " событий за " & _
        "отчётную неделю; в списке " & modContentMTO.FmtInt(CDbl(pick.Count)) & _
        " человек. ФИО во внешнюю модель не уходят.")
    BuildPeople = s
End Function

' Ячейка процента со стрелкой к предыдущей неделе (Б7): вверх - рост, вниз - падение.
' Равенство и отсутствие предыдущей недели - без стрелки.
Private Function PctArrowTd(ByVal p As Double, ByVal hasValue As Boolean, _
                            ByVal prev As Double, ByVal hasPrev As Boolean) As String
    Dim s As String
    s = modContentZone.PctTd(p, hasValue)
    If Not hasPrev Then PctArrowTd = s: Exit Function
    If Not hasValue Then PctArrowTd = s: Exit Function
    If Abs(p - prev) < 0.000000001 Then PctArrowTd = s: Exit Function
    Dim arrow As String, cls As String
    If p > prev Then
        arrow = ChrW$(&H2191): cls = "up"
    Else
        arrow = ChrW$(&H2193): cls = "dn"
    End If
    PctArrowTd = Replace$(s, "</span></td>", _
        "</span> <span class=""delta " & cls & """>" & arrow & "</span></td>")
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
        If mOrd.Exists(CStr(pr(0))) Then
            e = mOrd(CStr(pr(0)))
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
    For Each k In mOrd.Keys
        e = mOrd(k)
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

    s = s & modContentZone.NoteBlk("Разбор нарядов по подписанным статусам дирекции <b>" & _
        modContentMTO.Esc(dir) & "</b>: подписаны оба, один из двух, ни одного. Дата " & _
        "статуса одна на обе дирекции, поэтому числа у ДЭНТ и ДГМ совпадают; дирекции " & _
        "различаются АРМ подписей (" & dir & ": " & _
        modContentZone.Pc(modContentZone.SafePct(bothT, full), 1) & " нарядов целиком с " & _
        "планшета). График справа " & ChrW$(&H2014) & " неподписанные статусы в разрезе " & _
        "возраста наряда от даты создания. Период - весь снимок.")
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

' {{BLOCK_NOSIGN_*}} - Таблицы 5/6 (Б9/Б10): наряды без единой подписи
' (SignedCount = 0, как на слайде 4) в разрезе TekStatusPoDoc / zn_type.
' Двухколоночная разметка по образцу BuildSignStat: таблица + график справа.
Public Function BuildNoSignSplit(ByVal sDir As String, ByVal fld As Long, _
                                 ByVal emptyLab As String, ByVal colLab As String) As String
    EnsureDisc
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, e As Variant, tot As Double
    tot = 0#
    For Each k In mOrd.Keys
        e = mOrd(k)
        If SignedCount(e) = 0 Then
            Dim v As String
            v = Trim$(CStr(e(fld)))
            If v = "" Then v = emptyLab
            modContentZone.AddCnt d, v, 1#
            tot = tot + 1#
        End If
    Next k
    If tot = 0# Then BuildNoSignSplit = modContentMTO.EmptyNote(): Exit Function

    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 0, labs, vals

    Dim s As String, i As Long
    s = "<div class=""two-col wide-l""><div><table><thead><tr>" & _
        "<th>" & modContentMTO.Esc(colLab) & "</th><th class=""n"">Количество</th>" & _
        "<th class=""n"">Процент</th></tr></thead><tbody>"
    For i = 0 To UBound(labs)
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(vals(i))) & "</td><td class=""n"">" & _
            modContentZone.Pc(modContentZone.SafePct(CDbl(vals(i)), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table></div><div>"
    s = s & modContentZone.MockLabel("Количество нарядов")
    s = s & modContentZone.HBars(labs, vals, 620, 170, 24) & "</div></div>"
    s = s & modContentZone.NoteBlk("Наряды без единой подписи (0 из 4 подписей с АРМ " & _
        "ПК/ПЛАНШЕТ), разрез " & modContentMTO.Esc(colLab) & "; процент от числа таких " & _
        "нарядов (" & modContentMTO.FmtInt(tot) & "). Отбор не зависит от дирекции: у " & _
        "наряда без единой подписи нет подписей ни одной дирекции. Период - весь снимок.")
    BuildNoSignSplit = s
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
    For Each k In mOrd.Keys
        e = mOrd(k)
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
    For Each k In mOrd.Keys
        e = mOrd(k)
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
    For Each k In mOrd.Keys
        e = mOrd(k)
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
        "Возраст нарядов без единой подписи от даты создания, до конца выгрузки: " & _
        "снимок обрезан " & CStr(Day(se)) & " " & CStr(m(Month(se) - 1)) & " " & _
        CStr(Year(se)) & ". Старше 30 суток " & ChrW$(&H2014) & " " & _
        modContentMTO.FmtInt(vals(4)) & " нарядов.")
End Function

Public Function BuildUnsignedPost() As String
    Dim d As Object
    Set d = UnsignedBy(E_ZONE, NOPOST)
    If d.Count = 0 Then BuildUnsignedPost = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedPost = modContentZone.HBars(labs, vals, 660, 230, 24) & _
        modContentZone.NoteBlk("Наряды без единой подписи в разрезе нормализованной " & _
        "ремзоны <code>postN</code>; пустой postN " & ChrW$(&H2014) & _
        " отдельной строкой «(пост не указан)». Период - весь снимок.")
End Function

Public Function BuildUnsignedOwner() As String
    Dim d As Object
    Set d = UnsignedBy(E_OWNER, "(владелец не указан)")
    If d.Count = 0 Then BuildUnsignedOwner = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedOwner = modContentZone.HBars(labs, vals, 660, 280, 24) & _
        modContentZone.NoteBlk("Наряды без единой подписи в разрезе <code>owner_dep</code>; " & _
        "пустой владелец показан отдельной строкой «(владелец не указан)». Период - весь снимок.")
End Function

Public Function BuildUnsignedZnType() As String
    Dim d As Object
    Set d = UnsignedBy(E_TYPE, "(вид не указан)")
    If d.Count = 0 Then BuildUnsignedZnType = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedZnType = modContentZone.HBars(labs, vals, 680, 230, 24) & _
        modContentZone.NoteBlk("Наряды без единой подписи в разрезе <code>zn_type</code>; " & _
        "пустой вид показан отдельной строкой «(вид не указан)». Период - весь снимок.")
End Function

' =====================================================================================
' Точка входа: заполнение плейсхолдеров слайдов 2-4.
' =====================================================================================
Public Sub FillDiscPlaceholders(ByVal d As Object)
    Dim t0 As Single
    t0 = Timer
    EnsureDisc

    d("BLOCK_WEEKS_DENT") = BuildWeeksTable("ДЭНТ")
    d("BLOCK_TABLE1_DENT") = BuildBlock1("ДЭНТ")
    d("BLOCK_POSTS_DENT") = BuildPostsTable("ДЭНТ")
    d("BLOCK_PEOPLE_DENT") = BuildPeople("ДЭНТ")
    d("BLOCK_SIGNSTAT_DENT") = BuildSignStat("ДЭНТ")
    d("BLOCK_NOSIGN_STATUS_DENT") = BuildNoSignSplit("ДЭНТ", E_TEK, "(статус не указан)", "Текущий статус заказ-наряда")
    d("BLOCK_NOSIGN_TYPE_DENT") = BuildNoSignSplit("ДЭНТ", E_TYPE, "(вид не указан)", "Вид ремонта")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 2 готов: " & Round(Timer - t0, 2) & " c"

    d("BLOCK_WEEKS_DGM") = BuildWeeksTable("ДГМ")
    d("BLOCK_TABLE1_DGM") = BuildBlock1("ДГМ")
    d("BLOCK_POSTS_DGM") = BuildPostsTable("ДГМ")
    d("BLOCK_PEOPLE_DGM") = BuildPeople("ДГМ")
    d("BLOCK_SIGNSTAT_DGM") = BuildSignStat("ДГМ")
    d("BLOCK_NOSIGN_STATUS_DGM") = BuildNoSignSplit("ДГМ", E_TEK, "(статус не указан)", "Текущий статус заказ-наряда")
    d("BLOCK_NOSIGN_TYPE_DGM") = BuildNoSignSplit("ДГМ", E_TYPE, "(вид не указан)", "Вид ремонта")
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
