Attribute VB_Name = "modSelfTest"
Option Explicit

' modSelfTest - штатный тестовый модуль сквозного тестирования (run-e2e-tests-v1.ps1).
'
' Версия 1.0 от 07.09.2026:
'   - перенесён из временного tools/_tmp_selftest.bas (отладка v7.0) в штатное место tests\;
'   - добавлена процедура SelfTestUpsert: проверка замены строк после загрузки
'     tests/test_upsert_same_keys_v1.json (25 строк, у обеих строк ЗН-002 arm = ПК).
' Версия 1.1 от 07.09.2026:
'   - числовые CHECK пересчитаны под правила v7 (tests/expected.md, раздел
'     «Контроль отчёта v7.0»): REPORT/WEEK=0 (авто) -> отчётная неделя 202535,
'     окно Блока 1 = {101, 202535}. CHECK:BLOCK1_DGM_000, CHECK:BLOCK1_DENT_EMPTY,
'     CHECK:DASH_YTD_476, CHECK:BLOCK6_DGM_IVANOV; прежние BLOCK1_DGM_421 /
'     BLOCK1_DENT_000 / HTML_IVANOV_75 / HTML_IVANOV_8 удалены (эталон v6.1).
' Файл хранится в UTF-8; перед импортом в VBE перекодируется в ANSI 1251
' скриптом run-e2e-tests-v1.ps1 (редактор VBA читает .bas только в ANSI).

Private outPath As String

Private Sub WLine(s As String)
    Open outPath For Append As #1
    Print #1, s
    Close #1
End Sub

Private Function B2S(b As Boolean) As String
    If b Then B2S = "1" Else B2S = "0"
End Function

Public Sub SelfTest()
    On Error GoTo Fail
    outPath = Environ$("TEMP") & "\selftest_result.txt"
    On Error Resume Next
    Kill outPath
    On Error GoTo Fail

    WLine "START"

    ' 1. Число строк tbDATA
    WLine "CHECK:ROWS=" & CStr(ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").ListRows.Count)

    ' 2. Промпт: белый список, маркеры, 7 слайдов, без ФИО и defect_desc
    modContentMTO.BuildPivots
    Dim prompt As String
    prompt = modContentMTO.BuildPrompt()
    WLine "CHECK:PROMPT_HAS_BLOCK6=" & B2S(InStr(prompt, "block6_people") > 0)
    WLine "CHECK:PROMPT_HAS_EMP1=" & B2S(InStr(prompt, "[EMP_1]") > 0)
    WLine "CHECK:PROMPT_NO_IVANOV=" & B2S(InStr(prompt, "Иванов") = 0)
    WLine "CHECK:PROMPT_NO_DEFECTDESC=" & B2S(InStr(prompt, "не горит фара") = 0)
    WLine "CHECK:PROMPT_7SLIDES=" & B2S(InStr(prompt, "slide1_conclusions") > 0 And InStr(prompt, "slide7_conclusions") > 0)

    ' 3. ParseAIResponse с фиктивным ответом на 7 слайдов (маркеры [EMP_1], [EMP_2])
    Dim fake As String
    fake = "{""slide1_conclusions"":""Обзор [EMP_1].""," & _
           """slide2_conclusions"":""ДЭНТ [EMP_2].""," & _
           """slide3_conclusions"":""ДГМ [EMP_1] и [EMP_2].""," & _
           """slide4_conclusions"":""Синхронность [EMP_2].""," & _
           """slide5_conclusions"":""Дефекты [EMP_1].""," & _
           """slide6_conclusions"":""Динамика дирекций.""," & _
           """slide7_conclusions"":""Динамика постов.""}"
    Dim s3 As String, s4 As String, s5 As String
    Dim ok As Boolean
    ok = modContentMTO.ParseAIResponse(fake, s3, s4, s5)
    WLine "CHECK:PARSE_OK=" & B2S(ok)

    ' 4. Пошаговая диагностика составных частей BuildPlaceholders
    Dim d As Object
    Set d = modContentMTO.BuildPlaceholders("X3", "X4", "X5")
    WLine "AFTER_PLACEHOLDERS_1"
    WLine "CHECK:KPIS_DASH=" & B2S(InStr(CStr(d("DASH_1_YTD")), "Создали заказ-нарядов") > 0)
    WLine "CHECK:KPIS_DASH_PREV=" & B2S(InStr(CStr(d("DASH_1_PREV")), "Медиана") > 0)
    WLine "CHECK:BLOCK2_PREV=" & B2S(InStr(CStr(d("BLOCK_2_PREV")), "Нарядов") > 0)
    WLine "CHECK:BLOCK2_CHART=" & B2S(InStr(CStr(d("BLOCK_2_CHART")), "<svg") > 0)
    WLine "CHECK:B1_DENT_ALL=" & B2S(InStr(CStr(d("BLOCK_1_DENT_ALL")), "Дирекция: ДЭНТ") > 0)
    WLine "CHECK:B1_DGM_ZONES=" & B2S(InStr(CStr(d("BLOCK_1_DGM_ZONES")), "ремзоны: СТК") > 0)
    WLine "CHECK:B6_DENT=" & B2S(InStr(CStr(d("BLOCK_6_DENT")), "ВСЕГО ПОДПИСЕЙ") > 0)
    WLine "CHECK:B6_DGM=" & B2S(InStr(CStr(d("BLOCK_6_DGM")), "ВСЕГО ПОДПИСЕЙ") > 0)
    WLine "CHECK:B7=" & B2S(InStr(CStr(d("BLOCK_7_TABLE")), "10 аномалий") > 0)
    WLine "CHECK:B8=" & B2S(InStr(CStr(d("BLOCK_8_TABLE")), "Пар нарядов") > 0)
    WLine "CHECK:B9=" & B2S(InStr(CStr(d("BLOCK_9_TABLE")), "Вид ремонта") > 0)
    WLine "CHECK:B9A=" & B2S(InStr(CStr(d("BLOCK_9A_TABLE")), "% планшет по видам ремонта") > 0)
    WLine "CHECK:B6_RATING=" & B2S(InStr(CStr(d("BLOCK_6_RATING")), "Топ по использованию") > 0)
    WLine "CHECK:B4=" & B2S(InStr(CStr(d("BLOCK_4_TABLE")), "Динамика % планшет по дирекциям") > 0)
    WLine "CHECK:B5=" & B2S(InStr(CStr(d("BLOCK_5_TABLE")), "Динамика % планшет по постам") > 0)
    WLine "CHECK:DUMP=" & B2S(InStr(CStr(d("DATA_DUMP")), "drill-dump") > 0)

    ' 4а. Числовые контрольные значения v7 (tests/expected.md, раздел «Контроль отчёта v7.0»).
    ' REPORT/WEEK=0 (авто) -> отчётная неделя 202535; окно Блока 1 = {101, 202535}.
    ' ДГМ: 2 события ПК в окне, планшетов 0 -> «Итого» 0,0%. ДЭНТ: событий в окне нет -> «Нет данных».
    WLine "CHECK:BLOCK1_DGM_000=" & B2S(InStr(CStr(d("BLOCK_1_DGM_ALL")), "0,0%") > 0)
    WLine "CHECK:BLOCK1_DENT_EMPTY=" & B2S(InStr(CStr(d("BLOCK_1_DENT_ALL")), "Нет данных") > 0)
    WLine "CHECK:DASH_YTD_476=" & B2S(InStr(CStr(d("DASH_1_YTD")), "47,6%") > 0)
    WLine "CHECK:BLOCK6_DGM_IVANOV=" & B2S(InStr(CStr(d("BLOCK_6_DGM")), "Иванов И.И.") > 0 And InStr(CStr(d("BLOCK_6_DGM")), ">1<") > 0)

    ' 6. Замена маркеров на ФИО из кэша
    Dim t2 As String
    t2 = CStr(d("AI_INSIGHT_SLIDE_2"))
    WLine "CHECK:S2_MARKER_REPLACED=" & B2S(InStr(t2, "[EMP_2]") = 0 And InStr(t2, "Кузнецов") > 0)
    Dim t3 As String
    t3 = CStr(d("AI_INSIGHT_SLIDE_3"))
    WLine "CHECK:S3_FROM_CACHE=" & B2S(InStr(t3, "Иванов") > 0)

    ' 7. Сброс кэша + STUB-путь: протухших выводов предыдущего разбора быть не должно
    modContentMTO.BuildPivots
    Dim d2 As Object
    Set d2 = modContentMTO.BuildPlaceholders("[offline]", "[offline]", "[offline]")
    Dim s1b As String
    s1b = CStr(d2("AI_INSIGHT_SLIDE_1"))
    Dim s2b As String
    s2b = CStr(d2("AI_INSIGHT_SLIDE_2"))
    WLine "CHECK:STUB_NO_STALE=" & B2S(InStr(s1b, "Обзор") = 0 And InStr(s2b, "Кузнецов") = 0)
    WLine "CHECK:STUB_FALLBACK=" & B2S(InStr(s1b, "Внешний ИИ недоступен") > 0)

    ' 8. Полная сборка HTML и контроль содержимого
    modContentMTO.BuildPivots
    Dim d3 As Object
    Set d3 = modContentMTO.BuildPlaceholders("[offline]", "[offline]", "[offline]")
    Dim html As String
    html = modHTMLEngine.RenderTemplate(ThisWorkbook.Path & "\tmp_index.html", d3)
    WLine "CHECK:HTML_NO_PLACEHOLDER=" & B2S(InStr(html, "{{") = 0)
    WLine "CHECK:HTML_HAS_DUMP=" & B2S(InStr(html, "drill-dump") > 0)
    WLine "CHECK:HTML_HAS_ZN013=" & B2S(InStr(html, "ЗН-013") > 0)
    WLine "CHECK:HTML_HAS_IVANOV=" & B2S(InStr(html, "Иванов") > 0)
    WLine "CHECK:HTML_HAS_ANOMALIES=" & B2S(InStr(html, "10 аномалий") > 0)
    WLine "CHECK:HTML_HAS_RECIPE=" & B2S(InStr(html, "class='recipe'") > 0)
    WLine "CHECK:HTML_HAS_CHART=" & B2S(InStr(html, "<svg") > 0)
    WLine "CHECK:HTML_BLOCK7_PAIR=" & B2S(InStr(html, "2:00") > 0)
    WLine "CHECK:HTML_WEEKLY=" & B2S(InStr(html, "ВСЕГО ПОДПИСЕЙ") > 0)
    WLine "CHECK:HTML_GRADE=" & B2S(InStr(html, "мало данных") > 0 Or InStr(html, "провал") > 0)
    WLine "CHECK:HTML_7_TABS=" & B2S(InStr(html, "07 · Динамика по постам") > 0)

    modHTMLEngine.WriteUtf8 Environ$("TEMP") & "\selftest_out.html", html
    WLine "DONE"
    Exit Sub
Fail:
    On Error Resume Next
    WLine "CHECK:VBA_ERROR=" & Err.Description & "|" & CStr(Err.Number) & "|" & Erl
End Sub

' Проверка замены строк после загрузки tests/test_upsert_same_keys_v1.json (expected.md,
' Прогон 3): строк по-прежнему 25, у обеих строк ЗН-002 поле arm = ПК.
Public Sub SelfTestUpsert()
    On Error GoTo Fail
    outPath = Environ$("TEMP") & "\selftest_result.txt"
    WLine "UPSERT_START"

    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA")
    WLine "CHECK:UPSERT_ROWS=" & CStr(lo.ListRows.Count)

    Dim hdr As Range
    Set hdr = lo.HeaderRowRange
    Dim colNumber As Long, colArm As Long
    colNumber = 0: colArm = 0
    Dim cell As Range
    For Each cell In hdr.Cells
        If LCase$(CStr(cell.Value)) = "number" Then colNumber = cell.Column - hdr.Column + 1
        If LCase$(CStr(cell.Value)) = "arm" Then colArm = cell.Column - hdr.Column + 1
    Next cell

    Dim r As ListRow, okArm As Boolean, zn002 As Long
    okArm = True: zn002 = 0
    For Each r In lo.ListRows
        If CStr(r.Range.Cells(1, colNumber).Value) = "ЗН-002" Then
            zn002 = zn002 + 1
            If CStr(r.Range.Cells(1, colArm).Value) <> "ПК" Then okArm = False
        End If
    Next r

    WLine "CHECK:UPSERT_ZN002_COUNT=" & CStr(zn002)
    WLine "CHECK:UPSERT_ARM_PK=" & B2S(okArm And zn002 = 2)
    WLine "DONE_UPSERT"
    Exit Sub
Fail:
    On Error Resume Next
    WLine "CHECK:VBA_ERROR=" & Err.Description & "|" & CStr(Err.Number) & "|" & Erl
End Sub
