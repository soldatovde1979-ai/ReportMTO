Attribute VB_Name = "modSelfTest"
Option Explicit

' modSelfTest - штатный тестовый модуль сквозного тестирования (run-e2e-tests-v1.ps1).
'
' Версия 1.0 от 07.09.2026:
'   - перенесён из временного tools/_tmp_selftest.bas (отладка v7.0) в штатное место tests\;
'   - добавлена процедура SelfTestUpsert: проверка замены строк после загрузки
'     tests/test_upsert_same_keys_v1.json.
' Версия 1.1 от 07.09.2026:
'   - числовые CHECK пересчитаны под правила v7 (tests/expected.md).
' Версия 1.2 от 09.09.2026 (ТЗ v1.2, 4 слайда + ретеншн):
'   - эталоны переписаны под схему 4 слайдов (slide1_conclusions..slide4_conclusions,
'     плейсхолдеры KPI_OVERVIEW/BLOCK_WEEKS_*/BLOCK_PEOPLE_*/KPI_UNSIGNED и т.д.);
'   - ожидание строк tbDATA: 22 (тестовый JSON 25 строк; ретеншн KEEP_WEEKS=52
'     отсекает единственную строку вне окна 2025-08-25);
'   - проверка ключевых цифр отчёта по дампу от 09.09.2026: FACTS 22 события/12
'     нарядов, медиана 10:00 / p90 11,2 ч, ДГМ: всего ЗН 11 / подписаны полностью 7,
'     неподписанные: 1 ЗН (8,3 %), событий без подписи 4,5 %, «0 из 1» без поста;
'   - ParseAIResponse: fake-ответ на 4 слайда; контракт параметров наружу slide3/4/5
'     = слайды 2/3/4 (слайд 1 - только кэш mInsights);
'   - HTML: drill-dump в шаблоне v3.0 больше нет - проверки на mto-report/slide-nav.
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

    ' 1. Число строк tbDATA (эталон 22: 25 строк JSON - 1 строка, отсечённая ретеншном)
    Dim rowCnt As Long
    rowCnt = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").ListRows.Count
    WLine "CHECK:ROWS=" & CStr(rowCnt)
    WLine "CHECK:ROWS_22=" & B2S(rowCnt = 22)

    ' 2. Промпт под 4 слайда: маркеры слайдов, блоки, без defect_desc
    modContentMTO.BuildPivots
    Dim prompt As String
    prompt = modContentMTO.BuildPrompt()
    WLine "CHECK:PROMPT_4SLIDES=" & B2S(InStr(prompt, "slide1_conclusions") > 0 And InStr(prompt, "slide4_conclusions") > 0)
    WLine "CHECK:PROMPT_4BLOCKS=" & B2S(InStr(prompt, "slide1_overview") > 0 And InStr(prompt, "slide4_unsigned") > 0)
    WLine "CHECK:PROMPT_NO_DEFECTDESC=" & B2S(InStr(prompt, "не горит фара") = 0)

    ' 3. ParseAIResponse с фиктивным ответом на 4 слайда (контракт: slide3/4/5 = слайды 2/3/4)
    Dim fake As String
    fake = "{""slide1_conclusions"":""Обзор: аномалия у сотрудника.""," & _
           """slide2_conclusions"":""ДЭНТ: стабильно.""," & _
           """slide3_conclusions"":""ДГМ: рост.""," & _
           """slide4_conclusions"":""Неподписанные: рост.""}"
    Dim s3 As String, s4 As String, s5 As String
    Dim ok As Boolean
    ok = modContentMTO.ParseAIResponse(fake, s3, s4, s5)
    WLine "CHECK:PARSE_OK=" & B2S(ok)
    WLine "CHECK:PARSE_SPLIT=" & B2S(s3 = "ДЭНТ: стабильно." And s4 = "ДГМ: рост." And s5 = "Неподписанные: рост.")

    ' 4. BuildPlaceholders: ключевые цифры отчёта по эталону (дамп 09.09.2026)
    Dim d As Object
    Set d = modContentMTO.BuildPlaceholders("X3", "X4", "X5")
    WLine "AFTER_PLACEHOLDERS_1"
    WLine "CHECK:KPI_OPENED=" & B2S(InStr(CStr(d("KPI_OVERVIEW")), "Нарядов открыто") > 0)
    WLine "CHECK:KPI_SPARK=" & B2S(InStr(CStr(d("KPI_OVERVIEW")), "<svg") > 0)
    WLine "CHECK:FACTS_22_12=" & B2S(InStr(CStr(d("FACTS")), "22</dd>") > 0 And InStr(CStr(d("FACTS")), "12</dd>") > 0)
    WLine "CHECK:TIME_MED_P90=" & B2S(InStr(CStr(d("BLOCK_TIME_STATS")), "медиана 10:00") > 0 And InStr(CStr(d("BLOCK_TIME_STATS")), "p90 11,2 ч") > 0)
    WLine "CHECK:FLOW_ZNTYPE=" & B2S(InStr(CStr(d("BLOCK_FLOW_ZNTYPE")), "Внеплановый ремонт") > 0)
    WLine "CHECK:WEEKS_DGM=" & B2S(InStr(CStr(d("BLOCK_WEEKS_DGM")), "ДГМ · все ремзоны") > 0)
    WLine "CHECK:PEOPLE_DGM=" & B2S(InStr(CStr(d("BLOCK_PEOPLE_DGM")), "Кузнецов К.К.") > 0 And InStr(CStr(d("BLOCK_PEOPLE_DGM")), "Иванов И.И.") > 0)
    WLine "CHECK:PEOPLE_DENT=" & B2S(InStr(CStr(d("BLOCK_PEOPLE_DENT")), "Петров П.П.") > 0)
    WLine "CHECK:SIGNSTAT_DGM=" & B2S(InStr(CStr(d("BLOCK_SIGNSTAT_DGM")), "Подписаны полностью") > 0 And InStr(CStr(d("BLOCK_SIGNSTAT_DGM")), ">7<") > 0)
    WLine "CHECK:KPI_UNSIGNED=" & B2S(InStr(CStr(d("KPI_UNSIGNED")), "ЗН без единой подписи") > 0 And InStr(CStr(d("KPI_UNSIGNED")), "8,3 %") > 0)
    WLine "CHECK:UNSIGNED_SOURCE=" & B2S(InStr(CStr(d("BLOCK_UNSIGNED_SOURCE")), "0 из 1") > 0)
    WLine "CHECK:POSTS_DGM=" & B2S(InStr(CStr(d("BLOCK_POSTS_DGM")), "СТК") > 0)
    WLine "CHECK:POSTS_DENT_EMPTY=" & B2S(InStr(CStr(d("BLOCK_POSTS_DENT")), "Нет данных") > 0)

    ' 4а. Выводы ИИ из кэша (после ParseAIResponse fake выше)
    WLine "CHECK:AI_S1_CACHE=" & B2S(InStr(CStr(d("AI_INSIGHT_SLIDE_1")), "Обзор:") > 0)
    WLine "CHECK:AI_S3_CACHE=" & B2S(InStr(CStr(d("AI_INSIGHT_SLIDE_3")), "ДГМ:") > 0)

    ' 5. Сброс кэша + STUB-путь: протухших выводов предыдущего разбора быть не должно
    modContentMTO.BuildPivots
    Dim d2 As Object
    Set d2 = modContentMTO.BuildPlaceholders("[offline]", "[offline]", "[offline]")
    Dim s1b As String
    s1b = CStr(d2("AI_INSIGHT_SLIDE_1"))
    Dim s2b As String
    s2b = CStr(d2("AI_INSIGHT_SLIDE_2"))
    WLine "CHECK:STUB_NO_STALE=" & B2S(InStr(s1b, "Обзор:") = 0)
    WLine "CHECK:STUB_FALLBACK=" & B2S(InStr(s1b, "Внешний ИИ недоступен") > 0)
    WLine "CHECK:STUB_OFFLINE=" & B2S(InStr(s2b, "[offline]") > 0)

    ' 6. Полная сборка HTML и контроль содержимого (шаблон v3.0, 4 слайда)
    modContentMTO.BuildPivots
    Dim d3 As Object
    Set d3 = modContentMTO.BuildPlaceholders("[offline]", "[offline]", "[offline]")
    Dim html As String
    html = modHTMLEngine.RenderTemplate(ThisWorkbook.Path & "\tmp_index.html", d3)
    WLine "CHECK:HTML_NO_PLACEHOLDER=" & B2S(InStr(html, "{{") = 0)
    WLine "CHECK:HTML_ROOT=" & B2S(InStr(html, "mto-report") > 0)
    WLine "CHECK:HTML_NAV=" & B2S(InStr(html, "slide-nav") > 0)
    WLine "CHECK:HTML_SVG=" & B2S(InStr(html, "<svg") > 0)
    WLine "CHECK:HTML_TITLE=" & B2S(InStr(html, "Отчёт МТО") > 0)
    WLine "CHECK:HTML_OFFLINE=" & B2S(InStr(html, "[offline]") > 0)
    WLine "CHECK:HTML_PEOPLE=" & B2S(InStr(html, "Кузнецов К.К.") > 0)
    WLine "CHECK:HTML_ZNTYPE=" & B2S(InStr(html, "Внеплановый ремонт") > 0)

    modHTMLEngine.WriteUtf8 Environ$("TEMP") & "\selftest_out.html", html
    WLine "DONE"
    Exit Sub
Fail:
    On Error Resume Next
    WLine "CHECK:VBA_ERROR=" & Err.Description & "|" & CStr(Err.Number) & "|" & Erl
End Sub

' Проверка замены строк после загрузки tests/test_upsert_same_keys_v1.json
' (ожидание: 22 строки; у обеих строк ЗН-002 поле arm = ПК).
Public Sub SelfTestUpsert()
    On Error GoTo Fail
    outPath = Environ$("TEMP") & "\selftest_result.txt"
    WLine "UPSERT_START"

    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA")
    Dim rowCnt As Long
    rowCnt = lo.ListRows.Count
    WLine "CHECK:UPSERT_ROWS=" & CStr(rowCnt)
    WLine "CHECK:UPSERT_ROWS_22=" & B2S(rowCnt = 22)

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
    WLine "CHECK:UPSERT_ZN002_2=" & B2S(zn002 = 2)
    WLine "CHECK:UPSERT_ARM_PK=" & B2S(okArm And zn002 = 2)
    WLine "DONE_UPSERT"
    Exit Sub
Fail:
    On Error Resume Next
    WLine "CHECK:VBA_ERROR=" & Err.Description & "|" & CStr(Err.Number) & "|" & Erl
End Sub
