Attribute VB_Name = "modContentDisc"
' modContentDisc - CONTENT-слой части «Дисциплина» (слайды 2-4) отчёта МТО.
'
' Версия 1.9 от 01.10.2026 (владелец, 01.10.2026): слайды 2 и 3 приведены к постановке.
'   - ВОЗВРАЩЕНА «Таблица 1» (BuildPivot, {{BLOCK_TABLE1_*}}): Всего / ПЛАНШЕТ / ПК /
'     % планшет по 8 неделям, одна дирекция, ТОЛЬКО СТК и ПРК (REPORT/SLIDE_ZONES). Её
'     снимали в v1.2 как «дубль матрицы по ремзонам» - матрица другая (строки - ремзоны,
'     только проценты), и требование «только СТК и ПРК» терялось совсем.
'   - НОВАЯ «Таблица 2» ({{BLOCK_TABLE2_*}}): тот же вид, но все ремзоны.
'   - Порядок слайдов 2-3: Таблица 1 -> Таблица 2 -> ремзоны по неделям -> ремзоны за
'     неделю -> люди -> подразделения -> подписи. Матрица по ремзонам - расшифровка.
'   - BuildPeople: «% тек. нед.» + три прошлые недели в ОДНОЙ ячейке со стрелкой к
'     соседней более свежей неделе (вверх - неделя лучше), сортировка от лучшего к
'     худшему, у трёх лучших и трёх худших пометка, яркая шапка. Отдельных колонок
'     ПН-1..ПН-3, «Из них планшет», «Приёмка», «Выбытие» больше нет.
' Версия 1.8 от 30.09.2026: база слайда 4 и статистики подписания - без видов, которым
'   подпись не положена (Обслуживание при выпуске, Omnicomm): InBase, NoSignNote.
' Версия 1.7 от 29.09.2026: {{BLOCK_HISTORY_DENT}} / {{BLOCK_HISTORY_DGM}} -
'   помесячный «% планшета» на всю глубину (+ AutoLines - автовыводы слайдов 2-4): архив свёртки tbARCHIVE (modRollup)
'   плюс живые события. Живые события старше водяного знака свёртки не считаются,
'   иначе после неудачного Refresh месяц посчитался бы дважды.
' Версия 1.6 от 17.09.2026: BLOCK_ACCLEV_* снят со слайдов 2 и 3.
'   Владелец: «у меня на слайде уже перебор такой инфы». Тот же вопрос -
'   влияет ли способ подписи на результат - разобран на слайде 6 блоками
'   «возвраты по подписи ДЭНТ на выбытии», и там он доведён до ответа,
'   а не до процента. BuildAccLev оставлена в коде невызываемой.
' Версия 1.5 от 17.09.2026: конкретные даты в подписях периода слайдов 2-4,
'   как в modContentZone v2.12.
' Версия 1.3 от 17.09.2026: приёмка против выдачи.
'   Новое: BuildAccLev / {{BLOCK_ACCLEV_DENT}} и {{BLOCK_ACCLEV_DGM}} - события
'   подписания раздельно по статусу ready_for: ПЛАНШЕТ / ПК / % планшета по неделям
'   окна. Данные для этого в снимке были всегда (arm + ready_for), но нигде не
'   сводились: все блоки слайда считали «% планшета» по всем подписям сразу, и разрыв
'   между приёмкой и выдачей был неотличим от общего отставания дирекции.
'   Счётчики mALTot/mALTab наполняются в том же проходе EnsureDisc, лишнего прохода
'   по снимку не добавилось.
' Версия 1.2 от 16.09.2026: раскладка слайдов 2-4 по аудиту
'   docs\plans\audit_kod_i_slaidy_v1.0.md.
'   - ИСПРАВЛЕНО (P0-1 аудита): BuildNoSignSplit не использовал параметр дирекции,
'     и слайды ДЭНТ и ДГМ получали байт в байт одинаковый HTML. Параметр убран,
'     блок переехал на слайд 4; заодно снята гистограмма, повторявшая таблицу слева.
'   - СНЯТО со слайдов 2 и 3: {{BLOCK_TABLE1_*}} (BuildBlock1) - «Таблица 1»
'     повторяла матрицу BLOCK_WEEKS_*, стоявшую прямо над ней. Функция осталась
'     в коде невызываемой.
'   - СНЯТО со слайда 4: {{BLOCK_UNSIGNED_ZNTYPE}} (BuildUnsignedZnType) - та же
'     выборка и тот же разрез, что у нового {{BLOCK_NOSIGN_TYPE}}, но без процентов.
'     Функция осталась в коде невызываемой.
'   - ПОРЯДОК на слайдах 2 и 3: недели -> ремзоны -> люди -> подразделения ->
'     разбор подписей. Один показатель раскручивается вглубь, без возвратов назад.
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
Private mDTot As Object          ' «дирекция|подразделение|неделя» -> событий (R9)
Private mDTab As Object
Private mDAcc As Object          ' «дирекция|подразделение» за отчётную неделю
Private mDLev As Object
Private mDPairs As Object        ' «дирекция|подразделение» -> Collection «номер наряда|G/D»
Private mALTot As Object         ' «дирекция|A/L|неделя» -> событий с известным АРМ
Private mALTab As Object         ' то же, только планшет. A - «Готов к приемке», L - «Готов к выбытию»
Private mMonTot As Object         ' «дирекция|год*100+месяц» -> событий с известным АРМ (живые данные)
Private mMonTab As Object         ' то же, только планшет
Private mWm As Double             ' водяной знак свёртки: живые события старше него не считаются
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
    Set mDTot = Nothing
    Set mDTab = Nothing
    Set mDAcc = Nothing
    Set mDLev = Nothing
    Set mDPairs = Nothing
    Set mALTot = Nothing
    Set mALTab = Nothing
    Set mMonTot = Nothing
    Set mMonTab = Nothing
    mWm = 0#
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
    Set mDTot = CreateObject("Scripting.Dictionary")
    Set mDTab = CreateObject("Scripting.Dictionary")
    Set mDAcc = CreateObject("Scripting.Dictionary")
    Set mDLev = CreateObject("Scripting.Dictionary")
    Set mDPairs = CreateObject("Scripting.Dictionary")
    Set mALTot = CreateObject("Scripting.Dictionary")
    Set mALTab = CreateObject("Scripting.Dictionary")
    Set mMonTot = CreateObject("Scripting.Dictionary")
    Set mMonTab = CreateObject("Scripting.Dictionary")

    mRw = modContentZone.ZoneReportWeek()
    mWm = modRollup.Watermark()

    Dim hasZone As Boolean, hasOwner As Boolean, hasDep As Boolean, hasEmp As Boolean
    Dim hasTek As Boolean
    hasZone = modAggregate.HasColumn("postN")
    hasOwner = modAggregate.HasColumn("owner_dep")
    hasDep = modAggregate.HasColumn("emp_dep")
    hasEmp = modAggregate.HasColumn("employee")
    hasTek = modAggregate.HasColumn("TekStatusPoDoc")

    Dim n As Long, r As Long
    n = modAggregate.RowCount()
    mEvRows = 0
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
            ' v1.8: события видов без подписи (Обслуживание при выпуске, Omnicomm) в долю
            ' «НЕ ПОДПИСАНО» не идут - им подпись не положена (решение владельца 30.09.2026).
            If Not modContentZone.IsNoSignType(CStr(e(E_TYPE))) Then
                mEvRows = mEvRows + 1
                If Not signed Then mEvUnsigned = mEvUnsigned + 1
            End If

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
            ' v1.7 (задача 1.6): при нескольких записях одного статуса берётся САМАЯ
            ' РАННЯЯ подпись - как в modContentZone (минимальная дата). Раньше побеждала
            ' последняя строка таблицы, и «приёмка -> выбытие» на слайдах 2-3 и 7
            ' считалась от разных отметок. Неподписанная строка не затирает подписанную.
            If CDbl(e(ft)) = 0# Or (sd > 0# And sd < CDbl(e(ft))) Then
                If sd > 0# Or Len(CStr(e(fa))) = 0 Then
                    e(fa) = arm
                    e(ft) = sd
                End If
            End If
            mOrd(num) = e

            If signed And sd > 0# Then
                Dim wS As String, dk As String, zn As String
                wS = CStr(modContentZone.IsoYearWeek(sd))
                dk = IIf(isG, "ДГМ", "ДЭНТ")
                zn = Trim$(CStr(e(E_ZONE)))
                If zn = "" Then zn = NOPOST
                modContentZone.AddCnt mZoneTot, dk & "|" & zn & "|" & wS, 1#
                modContentZone.AddCnt mWeekTot, dk & "|" & wS, 1#
                ' Приёмка против выдачи: тот же % планшета, но раздельно по статусам.
                ' Ключ A/L, а не текст статуса - статус в выгрузке пишется через «е».
                Dim alK As String
                alK = dk & "|" & IIf(isAcc, "A", "L") & "|" & wS
                modContentZone.AddCnt mALTot, alK, 1#
                If tab1 Then modContentZone.AddCnt mALTab, alK, 1#
                ' История по месяцам (v1.7): событие старше водяного знака свёртки уже
                ' лежит в tbARCHIVE - его считает архив, иначе оно попадёт дважды.
                If mWm <= 0# Or sd >= mWm Or CDbl(e(E_DATE)) >= mWm Then
                    Dim mnK As String
                    mnK = dk & "|" & CStr(modContentZone.YearMonth(sd))
                    modContentZone.AddCnt mMonTot, mnK, 1#
                    If tab1 Then modContentZone.AddCnt mMonTab, mnK, 1#
                End If
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

                ' Агрегаты по подразделениям (R9: ключ с дирекцией, без ФИО).
                Dim dk2 As String
                dk2 = dk & "|" & dep
                modContentZone.AddCnt mDTot, dk2 & "|" & wS, 1#
                If tab1 Then modContentZone.AddCnt mDTab, dk2 & "|" & wS, 1#
                If CLng(wS) = mRw Then
                    If isAcc Then
                        modContentZone.AddCnt mDAcc, dk2, 1#
                    Else
                        modContentZone.AddCnt mDLev, dk2, 1#
                        If Not mDPairs.Exists(dk2) Then mDPairs.Add dk2, New Collection
                        mDPairs(dk2).Add num & Chr$(1) & IIf(isG, "G", "D")
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
    ' v1.8: неразрывный пробел, табуляция и переводы строки - тоже «пусто». Trim$ их
    ' не снимает, и в таблице постов появлялась строка без имени с нулями.
    z = Replace$(Replace$(Replace$(Replace$(CStr(e(E_ZONE)), ChrW$(&HA0), " "), vbTab, " "), vbCr, " "), vbLf, " ")
    z = Trim$(z)
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

' {{BLOCK_TABLE1_*}} / {{BLOCK_TABLE2_*}} - сводная «АРМ по неделям» (постановка владельца:
' «Таблица 1 только по СТК и ПРК; следующая такая же, но по всем ремзонам»). Строки:
' Всего / ПЛАНШЕТ / ПК / % планшет; столбцы - последние 8 недель (от старой к свежей) и
' «Всего»; ОДНА дирекция. Единица счёта - СОБЫТИЕ подписания с известным АРМ.
' zonesOnly = True  - только ремзоны REPORT/SLIDE_ZONES (по умолчанию СТК+ПРК): Таблица 1;
' zonesOnly = False - все ремзоны, включая «(пост не указан)»: Таблица 2.
' Считается из готовых mZoneTot/mZoneTab/mWeekTot/mWeekTab, лишнего прохода по снимку нет.
Public Function BuildPivot(ByVal dir As String, ByVal zonesOnly As Boolean) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(8)
    Dim n As Long, i As Long
    n = UBound(wk) + 1

    Dim zones As Object, pz As Variant
    Set zones = CreateObject("Scripting.Dictionary")
    For Each pz In Split(Replace$(Trim$(modMain.GetVariableDef("REPORT/SLIDE_ZONES", "СТК+ПРК")), "+", ";"), ";")
        If Len(Trim$(CStr(pz))) > 0 Then zones(Trim$(CStr(pz))) = True
    Next pz

    Dim tw() As Double, bw() As Double
    ReDim tw(0 To n - 1)
    ReDim bw(0 To n - 1)
    Dim k As Variant, parts As Variant
    If zonesOnly Then
        For Each k In mZoneTot.Keys
            parts = Split(CStr(k), "|")
            If CStr(parts(0)) = dir And zones.Exists(CStr(parts(1))) Then
                For i = 0 To n - 1
                    If CStr(parts(2)) = CStr(wk(i)) Then
                        tw(i) = tw(i) + modContentZone.DictVal(mZoneTot, CStr(k))
                        bw(i) = bw(i) + modContentZone.DictVal(mZoneTab, CStr(k))
                        Exit For
                    End If
                Next i
            End If
        Next k
    Else
        For i = 0 To n - 1
            tw(i) = modContentZone.DictVal(mWeekTot, dir & "|" & CStr(wk(i)))
            bw(i) = modContentZone.DictVal(mWeekTab, dir & "|" & CStr(wk(i)))
        Next i
    End If

    Dim hasAny As Boolean, sumT As Double, sumB As Double
    hasAny = False
    For i = 0 To n - 1
        If tw(i) > 0# Then hasAny = True
        sumT = sumT + tw(i)
        sumB = sumB + bw(i)
    Next i
    If Not hasAny Then BuildPivot = modContentMTO.EmptyNote(): Exit Function

    Dim s As String, r As Long
    s = "<div class=""scroll""><table><thead><tr><th>Дирекция / АРМ</th>"
    For i = 0 To n - 1
        s = s & "<th class=""n"">" & modContentZone.WLab(CLng(wk(i))) & "</th>"
    Next i
    s = s & "<th class=""n"">Всего</th></tr></thead><tbody>"
    For r = 0 To 3
        Select Case r
            Case 0
                s = s & "<tr class=""total""><td class=""head"">" & modContentMTO.Esc(dir) & _
                    " " & ChrW$(&HB7) & " всего</td>"
            Case 1
                s = s & "<tr><td>ПЛАНШЕТ</td>"
            Case 2
                s = s & "<tr><td>ПК</td>"
            Case Else
                s = s & "<tr><td>% планшет</td>"
        End Select
        For i = 0 To n - 1
            s = s & PivotCell(r, tw(i), bw(i))
        Next i
        s = s & PivotCell(r, sumT, sumB) & "</tr>"
    Next r
    s = s & "</tbody></table></div>"

    Dim scope As String
    If zonesOnly Then
        scope = "только ремзоны <code>REPORT/SLIDE_ZONES</code>: " & _
            modContentMTO.Esc(Join(zones.Keys, " + "))
    Else
        scope = "все ремзоны, включая «" & NOPOST & "»"
    End If
    s = s & modContentZone.NoteBlk("События подписания с <code>arm</code> из {ПК, ПЛАНШЕТ} " & _
        "дирекции <b>" & modContentMTO.Esc(dir) & "</b>, " & scope & ". Недели - последние 8 " & _
        "по дате статуса (<code>status_date</code>), слева старые, справа свежие. " & _
        "% планшет = ПЛАНШЕТ / (ПЛАНШЕТ + ПК); «НЕ ПОДПИСАНО» не входит.")
    BuildPivot = s
End Function

' Ячейка сводной: r = 0 всего, 1 планшет, 2 ПК, 3 % планшет.
Private Function PivotCell(ByVal r As Long, ByVal t As Double, ByVal b As Double) As String
    Select Case r
        Case 0
            PivotCell = "<td class=""n"">" & modContentMTO.FmtInt(t) & "</td>"
        Case 1
            PivotCell = "<td class=""n"">" & modContentMTO.FmtInt(b) & "</td>"
        Case 2
            PivotCell = "<td class=""n"">" & modContentMTO.FmtInt(t - b) & "</td>"
        Case Else
            PivotCell = modContentZone.PctTd(modContentZone.SafePct(b, t), t > 0#)
    End Select
End Function

' УСТАРЕЛО (v1.9): заменена BuildPivot, не вызывается. Прежнее описание:
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

' {{BLOCK_ACCLEV_*}} - приёмка против выдачи: чем подписывают. Единица счёта -
' СОБЫТИЕ подписания с известным АРМ (arm из {ПК, ПЛАНШЕТ}), неделя - по дате статуса.
'
' Зачем отдельный блок: соседние блоки слайда считают «% планшета» по всем подписям
' сразу, и разрыв между двумя статусами в них не виден. А он и есть рабочая гипотеза:
' приёмку инженер подписывает на площадке с планшета, а выдачу - вернувшись за ПК.
' Пока приёмка и выдача смешаны в одном проценте, этот перекос не отличить от общего
' отставания дирекции, и меры принимаются не те.
Public Function BuildAccLev(ByVal dir As String) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(8)
    Dim n As Long, i As Long
    n = UBound(wk) + 1

    Dim hasAny As Boolean
    hasAny = False
    For i = 0 To n - 1
        If modContentZone.DictVal(mALTot, dir & "|A|" & CStr(wk(i))) > 0# Then hasAny = True
        If modContentZone.DictVal(mALTot, dir & "|L|" & CStr(wk(i))) > 0# Then hasAny = True
    Next i
    If Not hasAny Then BuildAccLev = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<div class=""scroll""><table><thead><tr><th>Статус</th><th>АРМ</th>"
    For i = 0 To n - 1
        s = s & "<th class=""n"">" & modContentZone.WLab(CLng(wk(i))) & "</th>"
    Next i
    s = s & "<th class=""n"">Всего за окно</th></tr></thead><tbody>"

    Dim g As Long
    For g = 0 To 1
        Dim code As String, lab As String
        If g = 0 Then
            code = "A": lab = "Готов к приемке"
        Else
            code = "L": lab = "Готов к выбытию"
        End If

        Dim sTot As Double, sTab As Double
        sTot = 0#: sTab = 0#
        For i = 0 To n - 1
            sTot = sTot + modContentZone.DictVal(mALTot, dir & "|" & code & "|" & CStr(wk(i)))
            sTab = sTab + modContentZone.DictVal(mALTab, dir & "|" & code & "|" & CStr(wk(i)))
        Next i

        Dim r As Long
        For r = 0 To 2
            If r = 0 Then
                s = s & "<tr" & IIf(g = 1, " class=""total""", "") & _
                    "><th rowspan=""3"">" & modContentMTO.Esc(lab) & "</th>"
            Else
                s = s & "<tr>"
            End If
            Select Case r
                Case 0: s = s & "<td>ПЛАНШЕТ</td>"
                Case 1: s = s & "<td>ПК</td>"
                Case 2: s = s & "<td class=""head"">% планшета</td>"
            End Select

            For i = 0 To n - 1
                Dim t As Double, b As Double
                t = modContentZone.DictVal(mALTot, dir & "|" & code & "|" & CStr(wk(i)))
                b = modContentZone.DictVal(mALTab, dir & "|" & code & "|" & CStr(wk(i)))
                Select Case r
                    Case 0: s = s & "<td class=""n"">" & modContentMTO.FmtInt(b) & "</td>"
                    Case 1: s = s & "<td class=""n"">" & modContentMTO.FmtInt(t - b) & "</td>"
                    Case 2: s = s & modContentZone.PctTd(modContentZone.SafePct(b, t), t > 0#)
                End Select
            Next i

            Select Case r
                Case 0: s = s & "<td class=""n"">" & modContentMTO.FmtInt(sTab) & "</td>"
                Case 1: s = s & "<td class=""n"">" & modContentMTO.FmtInt(sTot - sTab) & "</td>"
                Case 2: s = s & modContentZone.PctTd( _
                    modContentZone.SafePct(sTab, sTot), sTot > 0#)
            End Select
            s = s & "</tr>"
        Next r
    Next g
    s = s & "</tbody></table></div>"

    ' Разрыв за окно: положительный - выдачу подписывают с планшета реже приёмки.
    Dim aT As Double, aB As Double, lT As Double, lB As Double
    aT = 0#: aB = 0#: lT = 0#: lB = 0#
    For i = 0 To n - 1
        aT = aT + modContentZone.DictVal(mALTot, dir & "|A|" & CStr(wk(i)))
        aB = aB + modContentZone.DictVal(mALTab, dir & "|A|" & CStr(wk(i)))
        lT = lT + modContentZone.DictVal(mALTot, dir & "|L|" & CStr(wk(i)))
        lB = lB + modContentZone.DictVal(mALTab, dir & "|L|" & CStr(wk(i)))
    Next i
    Dim gap As Double, verdict As String
    gap = modContentZone.SafePct(aB, aT) - modContentZone.SafePct(lB, lT)
    If aT = 0# Or lT = 0# Then
        verdict = "Один из статусов за окно не встречался " & ChrW$(&H2014) & _
            " сравнивать не с чем."
    ElseIf Abs(gap) < 1# Then
        verdict = "Приёмку и выдачу подписывают одинаково, разрыв меньше процентного пункта."
    ElseIf gap > 0# Then
        verdict = "Выдачу подписывают с планшета <b>реже</b> приёмки на <b>" & _
            modContentZone.FmtF(gap, 1) & " п.п.</b> " & ChrW$(&H2014) & " приёмка " & _
            modContentZone.Pc(modContentZone.SafePct(aB, aT), 1) & ", выдача " & _
            modContentZone.Pc(modContentZone.SafePct(lB, lT), 1) & _
            ". Похоже на возврат за ПК после работ; проверять стоит выдачу, не приёмку."
    Else
        verdict = "Выдачу подписывают с планшета <b>чаще</b> приёмки на <b>" & _
            modContentZone.FmtF(-gap, 1) & " п.п.</b> " & ChrW$(&H2014) & " приёмка " & _
            modContentZone.Pc(modContentZone.SafePct(aB, aT), 1) & ", выдача " & _
            modContentZone.Pc(modContentZone.SafePct(lB, lT), 1) & "."
    End If

    s = s & modContentZone.NoteBlk("События подписания дирекции <b>" & _
        modContentMTO.Esc(dir) & "</b> с <code>arm</code> из {ПК, ПЛАНШЕТ}, разрез по " & _
        "статусу <code>ready_for</code>; недели " & ChrW$(&H2014) & " по дате статуса " & _
        "(<code>status_date</code>). ПК = все события минус планшет. «НЕ ПОДПИСАНО» " & _
        "исключено, поэтому суммы не сходятся с числом нарядов. " & verdict)
    BuildAccLev = s
End Function

' {{BLOCK_POSTS_*}} - площадки за отчётную неделю. Единица счёта - ЗАКАЗ-НАРЯД:
' наряд считается «с планшета», только если ВСЕ подписи дирекции по нему за неделю
' поставлены с планшета. Смешивать с соседними блоками (там события) нельзя.
Public Function BuildPostsTable(ByVal dir As String) As String
    EnsureDisc
    Dim zn As Object, zt As Object
    Set zn = CreateObject("Scripting.Dictionary")
    Set zt = CreateObject("Scripting.Dictionary")

    ' И1: все ремзоны. Набор строк - REPORT/ZONES_ALL (разделитель «+»/«;»);
    ' по умолчанию - объединение REPORT/SLIDE_ZONES и всех postN снимка.
    Dim zonesAll As Object
    Set zonesAll = CreateObject("Scripting.Dictionary")
    Dim raw As String, pz As Variant, k As Variant, e As Variant
    raw = modMain.GetVariableDef("REPORT/ZONES_ALL", "")
    raw = Replace$(Trim$(raw), "+", ";")
    For Each pz In Split(raw, ";")
        If Len(Trim$(CStr(pz))) > 0 Then zonesAll(Trim$(CStr(pz))) = True
    Next pz
    If zonesAll.Count = 0 Then
        raw = Replace$(Trim$(modMain.GetVariableDef("REPORT/SLIDE_ZONES", "СТК+ПРК")), "+", ";")
        For Each pz In Split(raw, ";")
            If Len(Trim$(CStr(pz))) > 0 Then zonesAll(Trim$(CStr(pz))) = True
        Next pz
        For Each k In mOrd.Keys
            e = mOrd(k)
            zonesAll(ZoneOf(e)) = True
        Next k
    End If

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

    ' И1: ремзоны набора без нарядов за неделю - строками с нулями.
    Dim ext As Long
    ext = 0
    For Each pz In zonesAll.Keys
        If Not zn.Exists(CStr(pz)) Then ext = ext + 1
    Next pz
    If ext > 0 Then
        ReDim Preserve zl(0 To UBound(zl) + ext)
        ReDim Preserve zv(0 To UBound(zv) + ext)
        Dim gi As Long
        gi = UBound(zl) - ext
        For Each pz In zonesAll.Keys
            If Not zn.Exists(CStr(pz)) Then
                zl(gi) = CStr(pz)
                zv(gi) = 0#
                gi = gi + 1
            End If
        Next pz
    End If

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
        ChrW$(&H2014) & " отдельной строкой. Строки - все ремзоны набора " & _
        "<code>REPORT/ZONES_ALL</code> (по умолчанию SLIDE_ZONES + все postN " & _
        "снимка); ремзоны без нарядов за неделю показаны нулями.")
    BuildPostsTable = s
End Function

' {{BLOCK_PEOPLE_*}} - сотрудники дирекции (вид по постановке владельца, 01.10.2026):
' «% тек. нед.» и ТРИ прошлые недели в одной ячейке столбиком. Стрелка у прошлой недели -
' сравнение с СОСЕДНЕЙ БОЛЕЕ СВЕЖЕЙ неделей: ПН-1 с ПН, ПН-2 с ПН-1, ПН-3 с ПН-2;
' вверх и зелёная - неделя была лучше, вниз и красная - хуже, «=» - так же; у ПН
' (текущей) стрелки нет. Сортировка от лучшего к худшему по текущей неделе.
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

    ' TopKeys отдаёт ключи по убыванию значения: первый - лучший, последний - худший.
    Dim pl As Variant, pv As Variant
    modContentZone.TopKeys pick, 0, pl, pv
    Dim nP As Long, kMark As Long
    nP = UBound(pl) + 1
    kMark = 3
    If nP < 6 Then kMark = nP \ 2

    Dim s As String, i As Long
    s = "<table class=""ppl""><thead><tr><th>Сотрудник</th>" & _
        "<th class=""n"">% тек. нед.<br>" & modContentZone.WLab(mRw) & "</th>" & _
        "<th>Прошлые недели<br>ПН-1 " & ChrW$(&HB7) & " ПН-2 " & ChrW$(&HB7) & " ПН-3</th>" & _
        "<th class=""n"">Подписей за неделю</th>" & _
        "<th class=""n"">Приёмка " & ChrW$(&H2192) & " выбытие, медиана</th></tr></thead><tbody>"

    For i = 0 To UBound(pl)
        Dim emp As String
        emp = CStr(pl(i))
        s = s & "<tr><td class=""head"">" & modContentMTO.Esc(emp)
        If i < kMark Then s = s & "<span class=""rk"">" & modContentZone.Pill("good", "лучший") & "</span>"
        If i >= nP - kMark Then s = s & "<span class=""rk"">" & modContentZone.Pill("crit", "худший") & "</span>"
        s = s & "</td>"
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
        ' Колонка «% тек. нед.» - последняя неделя окна (индекс 3), без стрелки.
        s = s & modContentZone.PctTd(pj(3), hj(3))
        ' Прошлые недели столбиком: ПН-1 (индекс 2), ПН-2 (1), ПН-3 (0).
        s = s & "<td class=""stk"">"
        For j = 2 To 0 Step -1
            s = s & PplLine("ПН-" & CStr(3 - j), pj(j), hj(j), pj(j + 1), hj(j + 1), CLng(wk(j)))
        Next j
        s = s & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mPTot, dir & "|" & emp & "|" & CStr(mRw))) & "</td>"
        s = s & "<td class=""n"">" & AvgSpan(dir & "|" & emp) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"

    s = s & modContentZone.NoteBlk("«% тек. нед.» - % подписей с планшета на отчётной неделе " & _
        modContentZone.WLab(mRw) & " (" & modContentZone.WeekRange(mRw) & "). Прошлые недели: " & _
        "ПН-1 = " & modContentZone.WLab(CLng(wk(2))) & ", ПН-2 = " & modContentZone.WLab(CLng(wk(1))) & _
        ", ПН-3 = " & modContentZone.WLab(CLng(wk(0))) & ". Стрелка у прошлой недели - сравнение с " & _
        "соседней более свежей: ПН-1 с текущей, ПН-2 с ПН-1, ПН-3 с ПН-2; " & ChrW$(&H25B2) & _
        " - та неделя была лучше, " & ChrW$(&H25BC) & " - хуже, = - так же.")
    s = s & modContentZone.NoteBlk("Сортировка - от лучшего к худшему по текущей неделе; пометки " & _
        "«лучший»/«худший» - у трёх первых и трёх последних. Порог включения - не менее " & _
        modContentMTO.FmtInt(minRec) & " событий за отчётную неделю; в списке " & _
        modContentMTO.FmtInt(CDbl(pick.Count)) & " человек. ФИО во внешнюю модель не уходят. " & _
        "«Приёмка " & ChrW$(&H2192) & " выбытие» - медиана времени между подписями «Готов к " & _
        "приемке» и «Готов к выбытию» по нарядам, выбытие которых сотрудник подписал на " & _
        "отчётной неделе.")
    BuildPeople = s
End Function

' Строка «прошлая неделя» таблицы сотрудников: метка, процент и стрелка к соседней более
' свежей неделе (q). Вверх - эта неделя лучше, вниз - хуже, «=» - так же. Нет данных за
' неделю - прочерк без стрелки; нет данных у соседней - процент без стрелки.
Private Function PplLine(ByVal lab As String, ByVal p As Double, ByVal hasP As Boolean, _
                         ByVal q As Double, ByVal hasQ As Boolean, ByVal weekNo As Long) As String
    Dim v As String, a As String
    If hasP Then v = modContentZone.Pc(p, 0) Else v = modContentZone.Dash()
    a = ""
    If hasP And hasQ Then
        If Abs(p - q) < 0.000000001 Then
            a = " <span class=""delta flat"">=</span>"
        ElseIf p > q Then
            a = " <span class=""delta up"">" & ChrW$(&H25B2) & "</span>"
        Else
            a = " <span class=""delta dn"">" & ChrW$(&H25BC) & "</span>"
        End If
    End If
    PplLine = "<div class=""wk"" title=""" & modContentZone.WLab(weekNo) & """><span class=""wl"">" & _
        lab & "</span><span class=""wv"">" & v & "</span>" & a & "</div>"
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
    Dim col As Collection, i As Long, sum As Double, cnt As Long, spans As Collection, hasM As Boolean
    Set col = mPPairs(pk)
    sum = 0#: cnt = 0
    Set spans = New Collection
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
                spans.Add (l - a) * 24#
                cnt = cnt + 1
            End If
        End If
    Next i
    ' v1.8: медиана вместо среднего - одна стоянка на месяц раздувала «среднее» до суток.
    If cnt > 0 Then AvgSpan = modContentZone.Hh(modContentZone.MedianOf(spans, hasM))
End Function

' {{BLOCK_DEPTS_*}} - подразделения дирекции (по образцу BuildPeople, без сотрудников).
' Слева тренд % планшета за четыре недели, справа объём и время за отчётную неделю.
' Ключи агрегатов содержат дирекцию (решение R9): слайды 2 и 3 различаются.
Public Function BuildDepts(ByVal dir As String) As String
    EnsureDisc
    Dim wk As Variant
    wk = modContentZone.WeekWindow(4)

    ' Отбор (R7): подразделения с хотя бы одним событием за отчётную неделю.
    Dim pick As Object
    Set pick = CreateObject("Scripting.Dictionary")
    Dim k As Variant, parts As Variant
    For Each k In mDTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dir And CLng(parts(2)) = mRw Then
            pick(CStr(parts(1))) = modContentZone.SafePct( _
                modContentZone.DictVal(mDTab, CStr(k)), _
                modContentZone.DictVal(mDTot, CStr(k)))
        End If
    Next k
    If pick.Count = 0 Then BuildDepts = modContentMTO.EmptyNote(): Exit Function

    Dim dl As Variant, dv As Variant
    modContentZone.TopKeys pick, 0, dl, dv

    Dim s As String, i As Long
    s = "<table><thead><tr><th rowspan=""4"">Подразделение</th>" & _
        "<th class=""grp"" colspan=""4"">% планшета по неделям</th>" & _
        "<th class=""grp sep-l"" colspan=""5"">Отчётная неделя " & _
        modContentZone.WLab(mRw) & " (" & modContentZone.WeekRange(mRw) & ")</th></tr>"
    s = s & "<tr><th></th><th class=""n"">ПН-1 " & ChrW$(&HB7) & " " & _
        modContentZone.WLab(CLng(wk(2))) & "</th><th></th><th></th>" & _
        "<th class=""n sep-l"" rowspan=""3"">Всего подписей</th>" & _
        "<th class=""n"" rowspan=""3"">Из них планшет</th>" & _
        "<th class=""n"" rowspan=""3"">Приёмка</th>" & _
        "<th class=""n"" rowspan=""3"">Выбытие</th>" & _
        "<th class=""n"" rowspan=""3"">Приёмка " & ChrW$(&H2192) & " выбытие, медиана</th></tr>"
    s = s & "<tr><th class=""n"">ПН " & ChrW$(&HB7) & " " & _
        modContentZone.WLab(CLng(wk(3))) & "</th><th></th><th class=""n"">ПН-2 " & _
        ChrW$(&HB7) & " " & modContentZone.WLab(CLng(wk(1))) & "</th><th></th></tr>"
    s = s & "<tr><th></th><th></th><th></th><th class=""n"">ПН-3 " & _
        ChrW$(&HB7) & " " & modContentZone.WLab(CLng(wk(0))) & "</th></tr></thead><tbody>"

    For i = 0 To UBound(dl)
        Dim dep As String, depLab As String
        dep = CStr(dl(i))
        If dep = "" Then depLab = "(подразделение не указано)" Else depLab = dep
        s = s & "<tr><td class=""head"">" & modContentMTO.Esc(depLab) & "</td>"
        Dim j As Long
        Dim pj(0 To 3) As Double, hj(0 To 3) As Boolean
        For j = 0 To 3
            Dim key As String, t As Double, b As Double
            key = dir & "|" & dep & "|" & CStr(wk(j))
            t = modContentZone.DictVal(mDTot, key)
            b = modContentZone.DictVal(mDTab, key)
            pj(j) = modContentZone.SafePct(b, t)
            hj(j) = (t > 0#)
        Next j
        s = s & modContentZone.PctTd(pj(3), hj(3))
        s = s & PctArrowTd(pj(2), hj(2), pj(1), hj(1))
        s = s & PctArrowTd(pj(1), hj(1), pj(0), hj(0))
        Dim keyP As String, tPrev As Double
        keyP = dir & "|" & dep & "|" & CStr(modContentZone.PrevWeek(CLng(wk(0))))
        tPrev = modContentZone.DictVal(mDTot, keyP)
        s = s & PctArrowTd(pj(0), hj(0), _
            modContentZone.SafePct(modContentZone.DictVal(mDTab, keyP), tPrev), tPrev > 0#)
        Dim rk As String, tot As Double, tab1 As Double
        rk = dir & "|" & dep & "|" & CStr(mRw)
        tot = modContentZone.DictVal(mDTot, rk)
        tab1 = modContentZone.DictVal(mDTab, rk)
        s = s & "<td class=""n sep-l"">" & modContentMTO.FmtInt(tot) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt(tab1) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mDAcc, dir & "|" & dep)) & "</td>"
        s = s & "<td class=""n"">" & modContentMTO.FmtInt( _
            modContentZone.DictVal(mDLev, dir & "|" & dep)) & "</td>"
        s = s & "<td class=""n"">" & AvgSpanD(dir & "|" & dep) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"

    s = s & modContentZone.NoteBlk("% планшета за четыре недели (ПН-3 " & _
        ChrW$(&H2026) & " ПН) по дате статуса; стрелка " & ChrW$(&H2191) & " / " & _
        ChrW$(&H2193) & " - рост / падение процента к предыдущей неделе. " & _
        "Справа - объём и время за отчётную неделю " & modContentZone.WLab(mRw) & ". " & _
        "Порог включения - хотя бы одно событие подписания за отчётную неделю (R7); " & _
        "в списке " & modContentMTO.FmtInt(CDbl(pick.Count)) & _
        " подразделений. Подразделение - <code>emp_dep</code> сотрудника.")
    BuildDepts = s
End Function

' Среднее «приёмка -> выбытие» по нарядам подразделения (аналог AvgSpan на mDPairs).
Private Function AvgSpanD(ByVal dk2 As String) As String
    AvgSpanD = ChrW$(&H2014)
    If Not mDPairs.Exists(dk2) Then Exit Function
    Dim col As Collection, i As Long, sum As Double, cnt As Long, spans As Collection, hasM As Boolean
    Set col = mDPairs(dk2)
    sum = 0#: cnt = 0
    Set spans = New Collection
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
                spans.Add (l - a) * 24#
                cnt = cnt + 1
            End If
        End If
    Next i
    ' v1.8: медиана вместо среднего - одна стоянка на месяц раздувала «среднее» до суток.
    If cnt > 0 Then AvgSpanD = modContentZone.Hh(modContentZone.MedianOf(spans, hasM))
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
        If InBase(e) Then
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
        End If
    Next k
    If tot = 0# Then BuildSignStat = modContentMTO.EmptyNote(): Exit Function

    Dim s As String
    s = "<div class=""two-col wide-l""><div><table><thead><tr>" & _
        "<th>Заказ-наряды дирекции</th><th class=""n"">Нарядов</th>" & _
        "<th class=""n"">Доля, % от всех</th></tr></thead><tbody>"
    s = s & "<tr class=""total""><td class=""head"">Всего заказ-нарядов</td><td class=""n"">" & _
        modContentMTO.FmtInt(tot) & "</td><td class=""n"">" & _
        modContentZone.Pc(100#, 1) & "</td></tr>"
    s = s & SsRow("Из них подписаны полностью", full, tot, False)
    ' v1.8: у вложенных строк база другая - подписанные полностью; пишем это прямо.
    s = s & SsRow("оба статуса с планшета (% от подписанных полностью)", bothT, full, True)
    s = s & SsRow("один с планшета, другой с ПК (% от подписанных полностью)", mixT, full, True)
    s = s & SsRow("оба статуса с ПК (% от подписанных полностью)", bothP, full, True)
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
        "возраста наряда от даты создания. Период - с начала года (01.01.2026)." & NoSignNote())
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
' (SignedCount = 0) в разрезе TekStatusPoDoc / zn_type.
'
' v1.2 от 16.09.2026, две правки по аудиту раскладки:
'   - убран параметр дирекции. Он НЕ ИСПОЛЬЗОВАЛСЯ в теле функции, и слайды
'     ДЭНТ и ДГМ получали байт в байт одинаковый HTML под разными заголовками.
'     Отбор и не может зависеть от дирекции: у наряда без единой подписи нет
'     подписей ни одной. Блок переехал на слайд 4 «Не подписано вообще».
'   - снята гистограмма справа: она строилась по тем же labs/vals, что таблица
'     слева, то есть повторяла её числа один в один. Осталась таблица - в ней
'     есть проценты, которых на полосах не было.
Public Function BuildNoSignSplit(ByVal fld As Long, _
                                 ByVal emptyLab As String, ByVal colLab As String) As String
    EnsureDisc
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim k As Variant, e As Variant, tot As Double
    tot = 0#
    For Each k In mOrd.Keys
        e = mOrd(k)
        If SignedCount(e) = 0 And InBase(e) Then
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
    s = "<table><thead><tr>" & _
        "<th>" & modContentMTO.Esc(colLab) & "</th><th class=""n"">Количество</th>" & _
        "<th class=""n"">Процент</th></tr></thead><tbody>"
    For i = 0 To UBound(labs)
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(vals(i))) & "</td><td class=""n"">" & _
            modContentZone.Pc(modContentZone.SafePct(CDbl(vals(i)), tot), 1) & "</td></tr>"
    Next i
    s = s & "</tbody></table>"
    s = s & modContentZone.NoteBlk("Наряды без единой подписи (0 из 4 подписей с АРМ " & _
        "ПК/ПЛАНШЕТ), разрез " & modContentMTO.Esc(colLab) & "; процент от числа таких " & _
        "нарядов (" & modContentMTO.FmtInt(tot) & "). Отбор не зависит от дирекции: у " & _
        "наряда без единой подписи нет подписей ни одной дирекции " & ChrW$(&H2014) & _
        " поэтому блок стоит здесь, а не на слайдах ДЭНТ и ДГМ. Период - " & _
        "с начала года (01.01.2026)." & NoSignNote())
    BuildNoSignSplit = s
End Function

' =====================================================================================
' СЛАЙД 4. Не подписано.
' =====================================================================================
' База слайда 4 и статистики подписания (v1.8): наряд с начала года, которому подпись
' положена. «Обслуживание при выпуске» и Omnicomm планшетом не подписываются
' (владелец 30.09.2026) - раньше они давали 64 % «не подписанных вообще».
Private Function InBase(ByVal e As Variant) As Boolean
    InBase = False
    If Not modContentZone.InYtdOrd(e) Then Exit Function
    If modContentZone.IsNoSignType(CStr(e(E_TYPE))) Then Exit Function
    InBase = True
End Function

' Сколько нарядов с начала года исключено из базы как «подпись не положена».
Private Function NoSignExcluded() As Double
    Dim k As Variant, e As Variant, c As Double
    c = 0#
    For Each k In mOrd.Keys
        e = mOrd(k)
        If modContentZone.InYtdOrd(e) Then
            If modContentZone.IsNoSignType(CStr(e(E_TYPE))) Then c = c + 1#
        End If
    Next k
    NoSignExcluded = c
End Function

Private Function NoSignNote() As String
    NoSignNote = " Не входят в базу: " & modContentMTO.FmtInt(NoSignExcluded()) & _
        " нарядов «Обслуживание при выпуске» и Omnicomm " & ChrW$(&H2014) & _
        " планшетом они не подписываются по регламенту."
End Function

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
        If InBase(e) Then
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
    s = s & modContentZone.NoteBlk("Плитки слайда 4: нарядов без единой подписи - " & _
        "наряды без подписей с АРМ ПК/ПЛАНШЕТ (0 из 4), % от всех нарядов; событий " & _
        "«НЕ ПОДПИСАНО» - строки снимка с пустым АРМ, % от всех событий; подписан " & _
        "частично - наряды с 1-3 подписями из 4; самый старый - возраст в сутках " & _
        "самого старого неподписанного наряда от даты создания до конца снимка. " & _
        "Период - с начала года (01.01.2026)." & NoSignNote())
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
        If SignedCount(e) = 0 And InBase(e) Then
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
        If SignedCount(e) = 0 And CDbl(e(E_DATE)) > 0# And InBase(e) Then
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
        " отдельной строкой «(пост не указан)». Период - с начала года (01.01.2026).")
End Function

Public Function BuildUnsignedOwner() As String
    Dim d As Object
    Set d = UnsignedBy(E_OWNER, "(владелец не указан)")
    If d.Count = 0 Then BuildUnsignedOwner = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedOwner = modContentZone.HBars(labs, vals, 660, 280, 24) & _
        modContentZone.NoteBlk("Наряды без единой подписи в разрезе <code>owner_dep</code>; " & _
        "пустой владелец показан отдельной строкой «(владелец не указан)». Период - " & _
        "с начала года (01.01.2026).")
End Function

Public Function BuildUnsignedZnType() As String
    Dim d As Object
    Set d = UnsignedBy(E_TYPE, "(вид не указан)")
    If d.Count = 0 Then BuildUnsignedZnType = modContentMTO.EmptyNote(): Exit Function
    Dim labs As Variant, vals As Variant
    modContentZone.TopKeys d, 6, labs, vals
    BuildUnsignedZnType = modContentZone.HBars(labs, vals, 680, 230, 24) & _
        modContentZone.NoteBlk("Наряды без единой подписи в разрезе <code>zn_type</code>; " & _
        "пустой вид показан отдельной строкой «(вид не указан)». Период - с начала " & _
        "года (01.01.2026).")
End Function

' =====================================================================================
' Точка входа: заполнение плейсхолдеров слайдов 2-4.
' =====================================================================================
' {{BLOCK_HISTORY_*}} (v1.7) - «% планшета» дирекции по месяцам на всю глубину
' хранения: свёрнутый архив tbARCHIVE (modRollup) + живые события tbDATA.
' Зачем: матрица недель показывает 8 недель и не отвечает на вопрос «дисциплина
' растёт или это колебание». Помесячный ряд за всю историю отвечает.
' Единица счёта - СОБЫТИЕ подписания с известным АРМ, месяц - по дате подписи.
Public Function BuildHistory(ByVal dirName As String) As String
    Dim tot As Object, tb As Object, k As Variant, parts As Variant, arcRows As Long
    Dim mk As Object, ml As Variant, mv As Variant, n As Long, i As Long, first As Long
    Dim labs() As Variant, bars() As Variant, pcts() As Variant, s As String
    Dim p As Double, pPrev As Double, best As Double, worst As Double
    Dim bestM As String, worstM As String, hasArc As Boolean
    Const MAX_M As Long = 24

    EnsureDisc
    Set tot = CreateObject("Scripting.Dictionary")
    Set tb = CreateObject("Scripting.Dictionary")
    For Each k In mMonTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dirName Then
            modContentZone.AddCnt tot, CStr(parts(1)), modContentZone.DictVal(mMonTot, CStr(k))
            modContentZone.AddCnt tb, CStr(parts(1)), modContentZone.DictVal(mMonTab, CStr(k))
        End If
    Next k
    arcRows = modRollup.ArchiveTabletByMonth(dirName, tot, tb)
    hasArc = (arcRows > 0)

    ' Только месяцы с данными; ось - по возрастанию, последние MAX_M.
    Set mk = CreateObject("Scripting.Dictionary")
    For Each k In tot.Keys
        If CDbl(tot(k)) > 0# And CLng(k) > 190000 Then mk(CStr(k)) = CDbl(tot(k))
    Next k
    If mk.Count = 0 Then BuildHistory = modContentMTO.EmptyNote(): Exit Function
    ml = modAggregate.SortKeys(mk, True)
    n = UBound(ml) - LBound(ml) + 1
    first = LBound(ml)
    If n > MAX_M Then first = UBound(ml) - MAX_M + 1: n = MAX_M

    ReDim labs(0 To n - 1)
    ReDim bars(0 To n - 1)
    ReDim pcts(0 To n - 1)
    best = -1#: worst = 101#
    For i = 0 To n - 1
        labs(i) = modContentZone.MLab(CLng(ml(first + i)))
        bars(i) = modContentZone.DictVal(tot, CStr(ml(first + i)))
        pcts(i) = modContentZone.SafePct(modContentZone.DictVal(tb, CStr(ml(first + i))), CDbl(bars(i)))
        If CDbl(pcts(i)) > best Then best = CDbl(pcts(i)): bestM = CStr(labs(i))
        If CDbl(pcts(i)) < worst Then worst = CDbl(pcts(i)): worstM = CStr(labs(i))
    Next i

    s = ""
    If n >= 2 Then
        s = s & modContentZone.Collines(labs, bars, pcts, 1080, 260, 0, True)
        s = s & "<div class=""legend""><span><i style=""background:var(--s1)""></i>" & _
            "событий подписания за месяц</span><span><i style=""background:var(--s2)""></i>" & _
            "% с планшета</span></div>"
    End If

    s = s & "<div class=""scroll""><table><thead><tr><th>Месяц</th>" & _
        "<th class=""n"">Событий</th><th class=""n"">С планшета</th><th class=""n"">% планшета</th>" & _
        "<th class=""n"">К пред. месяцу</th></tr></thead><tbody>"
    pPrev = 0#
    For i = n - 1 To 0 Step -1
        p = CDbl(pcts(i))
        s = s & "<tr><td>" & modContentMTO.Esc(CStr(labs(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(CDbl(bars(i))) & "</td><td class=""n"">" & _
            modContentMTO.FmtInt(modContentZone.DictVal(tb, CStr(ml(first + i)))) & "</td>" & _
            modContentZone.PctTd(p, True)
        If i > 0 Then
            pPrev = CDbl(pcts(i - 1))
            s = s & "<td class=""n"">" & HistDelta(p - pPrev) & "</td></tr>"
        Else
            s = s & "<td class=""n"">" & modContentZone.Dash() & "</td></tr>"
        End If
    Next i
    s = s & "</tbody></table></div>"

    s = s & modContentZone.NoteBlk("Помесячный «% планшета» дирекции <b>" & _
        modContentMTO.Esc(dirName) & "</b>: подписания с планшета от событий подписания " & _
        "с известным АРМ (ПК, ПЛАНШЕТ), месяц - по дате подписи. Лучший месяц - " & _
        modContentMTO.Esc(bestM) & " (" & modContentZone.Pc(best, 1) & "), худший - " & _
        modContentMTO.Esc(worstM) & " (" & modContentZone.Pc(worst, 1) & "). " & _
        IIf(hasArc, "Месяцы старше окна хранения взяты из свёрнутого архива tbARCHIVE " & _
        "(сырые события удалены, счётчики сохранены).", _
        "Архив свёртки пуст - ряд собран по живым данным tbDATA; глубина ряда растёт " & _
        "с каждой свёрткой, а не обрывается ретеншном."))
    BuildHistory = s
End Function

Private Function HistDelta(ByVal dp As Double) As String
    If Abs(dp) < 0.05 Then
        HistDelta = "<span class=""delta flat"">= 0 п.п.</span>"
    ElseIf dp > 0# Then
        HistDelta = "<span class=""delta up"">" & ChrW$(&H25B2) & " " & _
            modContentZone.FmtF(dp, 1) & " п.п.</span>"
    Else
        HistDelta = "<span class=""delta dn"">" & ChrW$(&H25BC) & " " & _
            modContentZone.FmtF(-dp, 1) & " п.п.</span>"
    End If
End Function

Public Sub FillDiscPlaceholders(ByVal d As Object)
    Dim t0 As Single
    t0 = Timer
    EnsureDisc

    ' Подписи периода (PeriodCap) - централизованно, как в FillZonePlaceholders.
    ' Периоды - с конкретными датами, как на слайдах 1 и 5-8.
    Dim wl As String, w8 As String, w4 As String, ytd As String
    Dim wkD As Variant, w4D As Variant, weekLast As Date
    wkD = modContentZone.WeekWindow(8)
    w4D = modContentZone.WeekWindow(4)
    weekLast = modContentZone.WeekMonday(mRw) + 6
    wl = "неделя " & modContentZone.WLab(mRw) & " (" & _
        modContentZone.WeekRange(mRw) & "), по " & Format$(weekLast, "dd.mm.yyyy")
    w8 = "8 недель: " & modContentZone.WLab(CLng(wkD(0))) & " " & ChrW$(&H2192) & _
        " " & modContentZone.WLab(mRw) & ", по " & Format$(weekLast, "dd.mm.yyyy") & _
        ", по дате статуса"
    w4 = "4 недели: " & modContentZone.WLab(CLng(w4D(0))) & " " & ChrW$(&H2192) & _
        " " & modContentZone.WLab(mRw) & ", по дате статуса; объём и время - " & wl
    ytd = "с начала года: " & Format$(CDate(modContentZone.YtdStart()), "dd.mm.yyyy") & _
        " " & ChrW$(&H2192) & " " & _
        Format$(CDate(modContentZone.SnapshotEnd()), "dd.mm.yyyy")

    ' Слайды 2 и 3 раскручивают ОДИН показатель - «% подписаний с планшета» - вглубь:
    ' недели -> приёмка против выдачи -> ремзоны -> люди -> подразделения -> подписи.
    ' v1.9: «Таблица 1» (СТК и ПРК) и «Таблица 2» (все ремзоны) возвращены первыми на
    ' слайдах 2 и 3 (BuildPivot); матрица по ремзонам BLOCK_WEEKS_* - их расшифровка.
    ' BuildBlock1 заменена BuildPivot и оставлена в коде невызываемой.
    modMain.ShowProgress 78, "слайд 2 из 8: планшет ДЭНТ"
    d("BLOCK_TABLE1_DENT") = modContentZone.PeriodCap(w8) & BuildPivot("ДЭНТ", True)
    d("BLOCK_TABLE2_DENT") = modContentZone.PeriodCap(w8) & BuildPivot("ДЭНТ", False)
    d("BLOCK_WEEKS_DENT") = modContentZone.PeriodCap(w8) & BuildWeeksTable("ДЭНТ")
    d("BLOCK_POSTS_DENT") = modContentZone.PeriodCap(wl) & BuildPostsTable("ДЭНТ")
    d("BLOCK_PEOPLE_DENT") = modContentZone.PeriodCap(w4) & BuildPeople("ДЭНТ")
    d("BLOCK_DEPTS_DENT") = modContentZone.PeriodCap(w4) & BuildDepts("ДЭНТ")
    d("BLOCK_SIGNSTAT_DENT") = modContentZone.PeriodCap(ytd) & BuildSignStat("ДЭНТ")
    d("BLOCK_HISTORY_DENT") = modContentZone.PeriodCap(HistCap()) & BuildHistory("ДЭНТ")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 2 готов: " & Round(Timer - t0, 2) & " c"

    modMain.ShowProgress 84, "слайд 3 из 8: планшет ДГМ"
    d("BLOCK_TABLE1_DGM") = modContentZone.PeriodCap(w8) & BuildPivot("ДГМ", True)
    d("BLOCK_TABLE2_DGM") = modContentZone.PeriodCap(w8) & BuildPivot("ДГМ", False)
    d("BLOCK_WEEKS_DGM") = modContentZone.PeriodCap(w8) & BuildWeeksTable("ДГМ")
    d("BLOCK_POSTS_DGM") = modContentZone.PeriodCap(wl) & BuildPostsTable("ДГМ")
    d("BLOCK_PEOPLE_DGM") = modContentZone.PeriodCap(w4) & BuildPeople("ДГМ")
    d("BLOCK_DEPTS_DGM") = modContentZone.PeriodCap(w4) & BuildDepts("ДГМ")
    d("BLOCK_SIGNSTAT_DGM") = modContentZone.PeriodCap(ytd) & BuildSignStat("ДГМ")
    d("BLOCK_HISTORY_DGM") = modContentZone.PeriodCap(HistCap()) & BuildHistory("ДГМ")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 3 готов: " & Round(Timer - t0, 2) & " c"

    ' Слайд 4. Блоки «НЕ ПОДПИСАН» переехали сюда со слайдов 2 и 3: их выборка
    ' (наряды без единой подписи) от дирекции не зависит, и на тех слайдах они
    ' давали две одинаковые копии. Разрез по виду ремонта заменил прежний
    ' BLOCK_UNSIGNED_ZNTYPE - у них была одна и та же выборка и один и тот же
    ' разрез, но таблица даёт ещё и проценты. BuildUnsignedZnType остаётся в коде.
    modMain.ShowProgress 90, "слайд 4 из 8: не подписано"
    d("KPI_UNSIGNED") = modContentZone.PeriodCap(ytd) & BuildKpiUnsigned()
    d("BLOCK_UNSIGNED_AGE") = modContentZone.PeriodCap(ytd) & BuildUnsignedAge()
    d("BLOCK_UNSIGNED_POST") = modContentZone.PeriodCap(ytd) & BuildUnsignedPost()
    d("BLOCK_UNSIGNED_OWNER") = modContentZone.PeriodCap(ytd) & BuildUnsignedOwner()
    d("BLOCK_NOSIGN_TYPE") = modContentZone.PeriodCap(ytd) & _
        BuildNoSignSplit(E_TYPE, "(вид не указан)", "Вид ремонта")
    d("BLOCK_NOSIGN_STATUS") = modContentZone.PeriodCap(ytd) & _
        BuildNoSignSplit(E_TEK, "(статус не указан)", "Текущий статус заказ-наряда")
    modLog.WriteDebug 1, "Дисциплина", "FillDiscPlaceholders", _
        "Слайд 4 готов: " & Round(Timer - t0, 2) & " c"
End Sub

' Подпись периода блока истории: вся глубина хранения + архив свёртки.
Private Function HistCap() As String
    Dim wm As Double
    wm = modRollup.Watermark()
    If wm > 0# Then
        HistCap = "вся история по месяцам (до 24 последних): архив свёртки до " & _
            Format$(CDate(wm), "dd.mm.yyyy") & " + события tbDATA по " & _
            Format$(CDate(modContentZone.SnapshotEnd()), "dd.mm.yyyy")
    Else
        HistCap = "вся глубина tbDATA по месяцам (до 24 последних), по " & _
            Format$(CDate(modContentZone.SnapshotEnd()), "dd.mm.yyyy")
    End If
End Function

' =====================================================================================
' Автовыводы по правилам (v1.7): мнение по слайдам 2-4 без внешнего ИИ и сырьё для
' «Главное за неделю». Формат как у modContentZone.AutoLines: «уровень<TAB>текст»
' через vbLf; уровень crit / warn / good / info.
' =====================================================================================
Public Function AutoLines(ByVal slideNo As Long) As String
    On Error GoTo Fail
    Select Case slideNo
        Case 2: AutoLines = AutoDir("ДЭНТ")
        Case 3: AutoLines = AutoDir("ДГМ")
        Case 4: AutoLines = AutoUnsigned()
        Case Else: AutoLines = ""
    End Select
    Exit Function
Fail:
    modLog.WriteLogEntry Now, "Предупреждение", "Автовыводы", "modContentDisc.AutoLines", _
        "Слайд " & CStr(slideNo) & ": " & Err.Description
    AutoLines = ""
End Function

Private Function DL(ByVal sev As String, ByVal txt As String) As String
    DL = sev & vbTab & txt & vbLf
End Function

Private Function AutoDir(ByVal dirName As String) As String
    Dim s As String, pw As Long, t As Double, b As Double, tp As Double, bp As Double
    Dim p As Double, pPrev As Double, norma As Double, proval As Double, minRec As Double
    Dim k As Variant, parts As Variant, worstZ As String, worstP As Double, zt As Double
    Dim own As Double, alien As Double, w4 As Variant, i As Long, inW As Boolean
    Dim lowCnt As Long, pt As Double, pb As Double

    EnsureDisc
    ' Val, а не CDbl: не зависит от десятичного разделителя локали (как в modContentMTO).
    norma = Val(Replace(modMain.GetVariableDef("NormaForPlanshet", "90"), ",", "."))
    proval = Val(Replace(modMain.GetVariableDef("ProvalForPlanshet", "50"), ",", "."))
    If norma <= 0# Then norma = 90#
    If proval <= 0# Then proval = 50#
    minRec = modContentZone.ToNum(modMain.GetVariableDef("REPORT/MIN_POST_RECORDS", "10"))

    pw = modContentZone.PrevWeek(mRw)
    t = modContentZone.DictVal(mWeekTot, dirName & "|" & CStr(mRw))
    b = modContentZone.DictVal(mWeekTab, dirName & "|" & CStr(mRw))
    tp = modContentZone.DictVal(mWeekTot, dirName & "|" & CStr(pw))
    bp = modContentZone.DictVal(mWeekTab, dirName & "|" & CStr(pw))
    If t > 0# Then
        p = modContentZone.SafePct(b, t)
        If p >= norma Then
            s = s & DL("good", dirName & ": " & modContentZone.Pc(p, 1) & _
                " подписаний с планшета за отчётную неделю - в норме (порог " & _
                modContentZone.Pc(norma, 0) & ").")
        ElseIf p < proval Then
            s = s & DL("crit", dirName & ": только " & modContentZone.Pc(p, 1) & _
                " подписаний с планшета за отчётную неделю - ниже порога провала " & _
                modContentZone.Pc(proval, 0) & ", подпись идёт с ПК.")
        Else
            s = s & DL("warn", dirName & ": " & modContentZone.Pc(p, 1) & _
                " подписаний с планшета - между порогами " & modContentZone.Pc(proval, 0) & _
                " и " & modContentZone.Pc(norma, 0) & ".")
        End If
        If tp > 0# Then
            pPrev = modContentZone.SafePct(bp, tp)
            If p - pPrev <= -5# Then
                s = s & DL("crit", "Доля планшета у " & dirName & " упала на " & _
                    modContentZone.FmtF(pPrev - p, 1) & " п.п. к прошлой неделе (" & _
                    modContentZone.Pc(pPrev, 1) & " -> " & modContentZone.Pc(p, 1) & ").")
            ElseIf p - pPrev >= 5# Then
                s = s & DL("good", "Доля планшета у " & dirName & " выросла на " & _
                    modContentZone.FmtF(p - pPrev, 1) & " п.п. к прошлой неделе.")
            End If
        End If
    End If

    ' Худшая ремзона недели среди достаточно нагруженных.
    worstP = 101#
    For Each k In mZoneTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dirName And CStr(parts(2)) = CStr(mRw) Then
            zt = modContentZone.DictVal(mZoneTot, CStr(k))
            If zt >= minRec Then
                pt = modContentZone.SafePct(modContentZone.DictVal(mZoneTab, CStr(k)), zt)
                If pt < worstP Then worstP = pt: worstZ = CStr(parts(1))
            End If
        End If
    Next k
    If worstP < proval Then
        s = s & DL("warn", "Слабее всего ремзона «" & worstZ & "»: " & _
            modContentZone.Pc(worstP, 0) & " с планшета за неделю - адрес для разбора.")
    End If

    ' Подпись за чужую дирекцию: сотрудники другой дирекции за 4 недели.
    w4 = modContentZone.WeekWindow(4)
    own = 0#: alien = 0#
    For Each k In mDTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dirName And Len(CStr(parts(1))) > 0 Then
            inW = False
            For i = 0 To UBound(w4)
                If CStr(parts(2)) = CStr(w4(i)) Then inW = True: Exit For
            Next i
            If inW Then
                If CStr(parts(1)) = dirName Then
                    own = own + modContentZone.DictVal(mDTot, CStr(k))
                Else
                    alien = alien + modContentZone.DictVal(mDTot, CStr(k))
                End If
            End If
        End If
    Next k
    If alien > 0# Then
        s = s & DL(IIf(modContentZone.SafePct(alien, own + alien) >= 5#, "warn", "info"), _
            "За " & dirName & " за 4 недели подписали сотрудники другой дирекции: " & _
            modContentMTO.FmtInt(alien) & " событий (" & _
            modContentZone.Pc(modContentZone.SafePct(alien, own + alien), 1) & ").")
    End If

    ' Сотрудники дирекции ниже порога провала за отчётную неделю.
    lowCnt = 0
    For Each k In mPTot.Keys
        parts = Split(CStr(k), "|")
        If CStr(parts(0)) = dirName And CStr(parts(2)) = CStr(mRw) Then
            pt = modContentZone.DictVal(mPTot, CStr(k))
            If pt >= 5# Then
                pb = modContentZone.DictVal(mPTab, CStr(k))
                If modContentZone.SafePct(pb, pt) < proval Then lowCnt = lowCnt + 1
            End If
        End If
    Next k
    If lowCnt > 0 Then
        s = s & DL("warn", CStr(lowCnt) & " сотрудн. " & dirName & _
            " подписывали за неделю в основном с ПК (ниже " & modContentZone.Pc(proval, 0) & _
            " с планшета) - проверить, работает ли у них планшет.")
    End If
    AutoDir = s
End Function

Private Function AutoUnsigned() As String
    Dim k As Variant, e As Variant, tot As Double, none As Double, old30 As Double
    Dim s As String
    EnsureDisc
    For Each k In mOrd.Keys
        e = mOrd(k)
        If InBase(e) Then
            tot = tot + 1#
            If SignedCount(e) = 0 Then
                none = none + 1#
                If CDbl(e(E_DATE)) > 0# Then
                    If modContentZone.SnapshotEnd() - CDbl(e(E_DATE)) > 30# Then old30 = old30 + 1#
                End If
            End If
        End If
    Next k
    If tot = 0# Then AutoUnsigned = "": Exit Function
    s = DL(IIf(modContentZone.SafePct(none, tot) >= 10#, "crit", "warn"), _
        modContentMTO.FmtInt(none) & " нарядов с начала года (" & _
        modContentZone.Pc(modContentZone.SafePct(none, tot), 1) & ") не подписаны ни одной " & _
        "дирекцией - по ним дисциплину не проверить вообще.")
    If old30 > 0# Then
        s = s & DL("warn", modContentMTO.FmtInt(old30) & _
            " из них старше 30 суток - подписать задним числом уже нельзя, это потеря контроля.")
    End If
    AutoUnsigned = s
End Function
