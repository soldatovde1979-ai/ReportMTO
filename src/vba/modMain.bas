Attribute VB_Name = "modMain"
' modMain - CORE, оркестрация (Architecture Core, сценарии §11). Обработчики кнопок «Загрузить» /
' «Сформировать отчёт» на листе Main. Вызывает контрактные функции по фиксированным именам -
' их реализация для МТО находится в modContentMTO.bas:
'   modContentMTO.BuildPivots()                                   - подготовка аналитических блоков
'   modContentMTO.BuildPrompt() As String                          - тело запроса к внешнему ИИ (JSON)
'   modContentMTO.ParseAIResponse(...) As Boolean                  - разбор ответа ИИ
'   modContentMTO.BuildPlaceholders(s3, s4, s5) As Object          - словарь {{ИМЯ}} -> значение
'
' v3.1 (правки по ревью 24.08.2026):
'   P0-2  - выводы ИИ прокидываются в BuildPlaceholders (раньше результат ParseAIResponse
'           никуда не попадал и отчёт ВСЕГДА собирался с заглушками).
'   P1-10 - GenerateReport целиком под On Error: любая ошибка пишется в Logs и показывается
'           пользователю, а не роняет процедуру системным диалогом VBA.
'   P1-13 - проверка непустой tbDATA до начала расчётов.
'   P1-11 - путь result\ резолвится в modHTMLEngine.ResolveOutputFolder.
'   + GetVariableDef - чтение ключа Variable со значением по умолчанию (для необязательных ключей).
'
' v8.1 (11.09.2026) - ResolveAiApiKey: ключ чистится от BOM/кавычек/пробелов перед
'   заголовком Authorization (провайдер отвечал 401 "auth header format should be
'   Bearer sk-..."), источник ключа и его длина пишутся в журнал при DEBUG>=1.
'   Сам ключ в журнал не попадает.
'
' v8.2 (13.09.2026) - ТЗ v1.0: записи «было/стало» переведены в тип «Веха»,
'   манифест очереди файлов перед пакетной загрузкой, трассировка DEBUG=2
'   (вехи/тайминги) в LoadSourceFile и LoadPackage, сортировка ListJsonFiles
'   по дате/времени из имени файла.
'
' v9.1 (29.09.2026) - прогресс сборки отчёта: ShowProgress / EndProgress (строка
'   состояния + полоса на листе Main), защита от повторного запуска во время сборки.
' v9.0 (29.09.2026) - свёртка истории: modRollup.RollupBeforeLoad перед Refresh
'   в LoadSourceFile и LoadPackage; в GenerateReport ошибка слоя ИИ проверяется
'   после каждого шага (задача 9.5 docs/task-rep-v3.md).
'
' v7.1 (08.09.2026) - ключ ИИ вынесен за пределы книги:
'   ResolveAiApiKey: AI_API_KEY (переменная окружения) -> %APPDATA%\ReportMTO\deepseek.key
'   (UTF-8) -> лист Variable (legacy, с предупреждением в лог). Ключ в лог и HTML не пишется.
Option Explicit

' Прогресс сборки отчёта (v9.1): см. ShowProgress / EndProgress в конце модуля.
Private mPrgT0 As Single
Private mPrgStage As String
Private mBusy As Boolean

Public Sub LoadSourceFile()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("JSON files (*.json), *.json", , "Выберите файл выгрузки")
    If filePath = False Then Exit Sub

    Dim rowsBefore As Long
    rowsBefore = SafeRowCount()

    On Error GoTo ErrHandler

    ' Трассировка DEBUG=2 (ТЗ v1.0, п.2.3). Граница возможностей: строки внутри
    ' Power Query VBA не видит, поэтому «каждая строка» заменена вехами/таймингами -
    ' максимум достижимого при одном проходе.
    Dim t0 As Single
    t0 = Timer
    modLog.WriteDebug 2, "Загрузка данных", "LoadSourceFile", "Старт (строк до: " & rowsBefore & ")"

    ' v9.0: свёртка истории ДО Refresh - ретеншн Power Query удалит только то,
    ' что уже лежит в tbARCHIVE (modRollup, qRollupUntil). Ошибка свёртки загрузку
    ' не останавливает: водяной знак не сдвигается, и ничего лишнего не удаляется.
    modRollup.RollupBeforeLoad

    SetSourcePathParameter CStr(filePath)
    ' Итог этапа 1/3 (вариант А - логирование в оркестраторе): полный путь нужен для
    ' возобновления загрузки после обрыва на следующих этапах.
    modLog.WriteLogEntry Now, "Инфо", "Параметр PQ", "prmSourcePath", _
        "Путь выбранного файла: " & CStr(filePath)

    modLog.WriteDebug 2, "Загрузка данных", "LoadSourceFile", "До Refresh"
    modPQSync.RefreshImportQuery
    modLog.WriteDebug 2, "Загрузка данных", "LoadSourceFile", "После Refresh (" & Round(Timer - t0, 1) & " c)"

    Dim rowsAfter As Long
    rowsAfter = SafeRowCount()

    ' Итог этапа 2/3: строки до/после - вместо прежней общей записи «Загрузка данных»,
    ' чтобы не дублировать одни и те же цифры в двух строках лога.
    ' Тип «Веха» (ТЗ v1.0, п.2.1): попадает в лист Logs и в файл при DEBUG>=1.
    modLog.WriteMilestone "Обновление импорта", "Query-ImportJSON", _
        "Строк до: " & rowsBefore & "; строк после: " & rowsAfter

    modContentMTO.BuildPivots
    modLog.WriteDebug 2, "Загрузка данных", "LoadSourceFile", "BuildPivots (" & Round(Timer - t0, 1) & " c)"
    ' Итог этапа 3/3: BuildPivots не возвращает результат, поэтому фиксируем факт
    ' завершения и текущий объём tbDATA.
    modLog.WriteLogEntry Now, "Инфо", "Сводки", "modContentMTO.BuildPivots", _
        "Сводки построены; строк в tbDATA: " & SafeRowCount()

    modLog.WriteMilestone "Загрузка данных", "LoadSourceFile", _
        "Загрузка завершена. Строк в tbDATA: " & rowsAfter
    Exit Sub

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Загрузка данных", Dir(CStr(filePath)), Err.Description
    ' v5 (M-2): прежний текст утверждал, что виноват файл. Большинство отказов на этом пути
    ' к файлу отношения не имеют (подключение не найдено, Formula Firewall, обрыв на типизации),
    ' и формулировка уводила от причины.
    modLog.WriteLogEntry Now, "Ошибка", "Загрузка данных", "LoadSourceFile", _
        "Загрузка не выполнена: " & Err.Description & _
        ". Если упоминается Formula.Firewall - разово выключите уровни конфиденциальности " & _
        "(Данные -> Получить данные -> Параметры запроса -> Конфиденциальность)."
End Sub

' «Загрузить пакет»: пересобирает tbDATA из всех *.json папки DATA/SOURCE_FOLDER
' одним проходом Power Query (режим папки R-2 в Query-ImportJSON).
'
' Порядок файлов - по имени (время выгрузки зашито в имя sppr_tablet_YYYYMMDD_HHMMSS_...).
' Вложенные папки НЕ читаются (только верхний уровень).
'
' v1.1 (12.09.2026): пофайловый upsert заменён одним проходом по папке - при 255 тыс.
'   строк в tbDATA upsert одного файла занимал 5 ч 12 мин (замер R-2). Перед загрузкой
'   tbDATA и лог очищаются: в Logs остаётся только итоговая запись «Записано N из M»,
'   где N - строк в tbDATA после загрузки, M - суммарное число записей во всех файлах
'   (число берётся из имени файла _NNNNNrec.json).
Public Sub LoadPackage()
    Dim folder As String
    folder = Trim$(GetVariableDef("DATA/SOURCE_FOLDER", ""))
    If folder = "" Then
        modLog.WriteLogEntry Now, "Ошибка", "Загрузка пакета", "LoadPackage", _
            "Ключ DATA/SOURCE_FOLDER пуст на листе Variable - укажите папку с выгрузками."
        Exit Sub
    End If

    Dim files As Variant
    files = ListJsonFiles(folder)
    If IsEmpty(files) Then
        modLog.WriteLogEntry Now, "Ошибка", "Загрузка пакета", "LoadPackage", _
            "В папке нет *.json файлов: " & folder
        Exit Sub
    End If

    Dim totalRecs As Long
    totalRecs = SumRecsFromNames(files)

    On Error GoTo ErrHandler

    ' Трассировка DEBUG=2 (ТЗ v1.0, п.2.3). Граница возможностей: строки внутри
    ' Power Query VBA не видит, поэтому «каждая строка» заменена вехами/таймингами -
    ' максимум достижимого при одном проходе.
    Dim t0 As Single
    t0 = Timer
    Dim rowsBeforePkg As Long
    rowsBeforePkg = SafeRowCount()
    modLog.WriteDebug 2, "Загрузка пакета", "LoadPackage", "Старт (строк до: " & rowsBeforePkg & ")"

    ClearLogs
    ' v9.0: свёртка - до пересборки, пока в tbDATA лежат прежние строки.
    modRollup.RollupBeforeLoad
    ClearTbData

    ' Манифест очереди (ТЗ v1.0, п.2.2): одна запись «Веха» ДО Refresh - перечень
    ' файлов в порядке очереди: номер, имя, дата/время из имени, rec-счётчик.
    ' Это реализация «записей между файлами» при однопроходной загрузке.
    Dim manifest As String
    Dim i As Long, fnameM As String
    manifest = "Очередь файлов (" & (UBound(files) - LBound(files) + 1) & "):"
    For i = LBound(files) To UBound(files)
        fnameM = Mid$(CStr(files(i)), InStrRev(CStr(files(i)), "\") + 1)
        manifest = manifest & vbCrLf & (i - LBound(files) + 1) & ") " & fnameM & _
            " | дата/время: " & StampFromName(fnameM) & " | rec: " & RecFromName(fnameM)
    Next i
    modLog.WriteMilestone "Загрузка пакета", "LoadPackage", manifest

    ' Прогноз ожидания (Задача C)
    Dim lastSecsStr As String, lastRowsStr As String
    Dim lastSecs As Single, lastRows As Long, estMins As Single
    Dim progMsg As String
    lastSecsStr = Trim$(GetVariableDef("DATA/LAST_LOAD_SECONDS", ""))
    lastRowsStr = Trim$(GetVariableDef("DATA/LAST_LOAD_ROWS", ""))
    If lastSecsStr <> "" And lastRowsStr <> "" And IsNumeric(lastSecsStr) And IsNumeric(lastRowsStr) Then
        lastSecs = CSng(Replace(lastSecsStr, ".", ",")) ' терпимость к культуре
        lastRows = CLng(lastRowsStr)
        If lastRows > 0 Then
            estMins = (lastSecs / lastRows * totalRecs) / 60
            progMsg = "Refresh начат: " & totalRecs & " записей, ожидание ~" & Round(estMins, 1) & " мин"
        Else
            progMsg = "Refresh начат: " & totalRecs & " записей, прогноза нет (первый замер)"
        End If
    Else
        progMsg = "Refresh начат: " & totalRecs & " записей, прогноза нет (первый замер)"
    End If
    modLog.WriteMilestone "Загрузка пакета", "LoadPackage", progMsg
    Application.StatusBar = progMsg

    modLog.WriteDebug 2, "Загрузка пакета", "LoadPackage", "До Refresh"
    SetSourcePathParameter folder
    modPQSync.RefreshImportQuery
    Dim elapsedSecs As Single
    elapsedSecs = Timer - t0
    modLog.WriteDebug 2, "Загрузка пакета", "LoadPackage", "После Refresh (" & Round(elapsedSecs, 1) & " c)"
    Application.StatusBar = False

    Dim finalRows As Long
    finalRows = SafeRowCount()

    ' Записываем факт
    SetVariable "DATA/LAST_LOAD_SECONDS", CStr(Round(elapsedSecs, 1))
    SetVariable "DATA/LAST_LOAD_ROWS", CStr(finalRows)

    ' Тип «Веха» (ТЗ v1.0, п.2.1): сверка «Записано N из M» - в лист Logs и в файл.
    modLog.WriteMilestone "Загрузка пакета", "Query-ImportJSON", _
        "Записано " & finalRows & " из " & totalRecs

    modContentMTO.BuildPivots
    modLog.WriteDebug 2, "Загрузка пакета", "LoadPackage", "BuildPivots (" & Round(Timer - t0, 1) & " c)"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    modLog.WriteLogEntry Now, "Ошибка", "Загрузка пакета", "LoadPackage", _
        "Загрузка пакета прервана: " & Err.Description
End Sub

' Перечисляет *.json только верхнего уровня папки (без рекурсии), сортирует по дате/времени из имени.
' Возвращает Variant-массив путей (0-based) или Empty, если файлов нет.
Private Function ListJsonFiles(folder As String) As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        ListJsonFiles = Empty
        Exit Function
    End If

    ' Собираем пути в массив, затем сортируем пузырьком по имени (время выгрузки
    ' уже зашито в имя файла sppr_tablet_YYYYMMDD_HHMMSS_...). Без внешних COM-типов.
    Dim arr() As String
    Dim cnt As Long
    cnt = 0

    Dim f As Object
    For Each f In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "json" Then
            ReDim Preserve arr(0 To cnt)
            arr(cnt) = f.Path
            cnt = cnt + 1
        End If
    Next f

    If cnt = 0 Then
        ListJsonFiles = Empty
        Exit Function
    End If

    ' Пузырьковая сортировка по дате/времени, распарсенной из имени файла
    ' (sppr_tablet_YYYYMMDD_HHMMSS_*). Если дату извлечь не удалось - fallback
    ' на сравнение по имени файла (ТЗ v1.0, п.2.4).
    Dim i As Long, j As Long, tmp As String
    Dim keyI As String, keyJ As String
    For i = 0 To cnt - 2
        For j = i + 1 To cnt - 1
            keyI = StampFromName(arr(i))
            keyJ = StampFromName(arr(j))
            If keyI = "" Then keyI = arr(i)
            If keyJ = "" Then keyJ = arr(j)
            If StrComp(keyI, keyJ, vbTextCompare) > 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    ListJsonFiles = arr
End Function

' Извлекает дату/время YYYYMMDD_HHMMSS из имени файла (шаблон sppr_tablet_YYYYMMDD_HHMMSS_*).
' Возвращает "" если шаблон не найден. Чистый строковый скан, без внешних библиотек.
Private Function StampFromName(fileName As String) As String
    Dim i As Long
    StampFromName = ""
    For i = 1 To Len(fileName) - 14
        If Mid$(fileName, i + 8, 1) = "_" And IsDigits(Mid$(fileName, i, 8)) _
            And IsDigits(Mid$(fileName, i + 9, 6)) Then
            StampFromName = Mid$(fileName, i, 15)
            Exit Function
        End If
    Next i
End Function

' True, если s непустая и состоит только из цифр.
Private Function IsDigits(s As String) As Boolean
    Dim i As Long
    IsDigits = (Len(s) > 0)
    If Not IsDigits Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then
            IsDigits = False
            Exit Function
        End If
    Next i
End Function

' Число записей из имени файла (_NNNNNrec.json); 0, если не распозналось.
' Логика извлечения - как в SumRecsFromNames (та функция не менялась, ТЗ v1.0 п.2.5).
Private Function RecFromName(fileName As String) As Long
    Dim s As String, p As Long, digits As String, ch As String
    RecFromName = 0
    s = LCase$(fileName)
    p = InStr(s, "rec.")
    If p <= 1 Then Exit Function
    digits = ""
    p = p - 1
    Do While p >= 1
        ch = Mid$(s, p, 1)
        If ch >= "0" And ch <= "9" Then
            digits = ch & digits
            p = p - 1
        Else
            Exit Do
        End If
    Loop
    If digits <> "" Then RecFromName = CLng(digits)
End Function

' Суммирует число записей, зашитое в имени файла выгрузки (_NNNNNrec.json).
' Имя не распозналось - файл даёт 0 в сумму (в итоге M просто меньше реального).
Private Function SumRecsFromNames(files As Variant) As Long
    Dim total As Long, i As Long, p As Long
    Dim digits As String, ch As String, fname As String
    total = 0
    For i = LBound(files) To UBound(files)
        fname = LCase$(CStr(files(i)))
        p = InStr(fname, "rec.")
        If p > 1 Then
            digits = ""
            p = p - 1
            Do While p >= 1
                ch = Mid$(fname, p, 1)
                If ch >= "0" And ch <= "9" Then
                    digits = ch & digits
                    p = p - 1
                Else
                    Exit Do
                End If
            Loop
            If digits <> "" Then total = total + CLng(digits)
        End If
    Next i
    SumRecsFromNames = total
End Function

' Очищает таблицу tbDATA (удаляет строки, оставляя заголовок) для полной пересборки
' из папки. Вызывается только из LoadPackage.
Private Sub ClearTbData()
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tbDATA", vbTextCompare) = 0 Then
                ' If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete ' Приводит к зависанию на 255k строках! PowerQuery затрет сам
                Exit Sub
            End If
        Next lo
    Next ws
End Sub

' Очищает лог книги (таблицу tbLogs на листе Logs): после загрузки пакета
' в логе остаётся только итоговая запись «Записано N из M».
Private Sub ClearLogs()
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Sheets("Logs").ListObjects("tbLogs")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

' Копия книги в bak\ рядом с книгой: ReportMTO.xlsm.bak_YYYYMMDD_HHMMSS.
' SaveCopyAs не трогает открытую книгу - можно продолжать загрузку.
Private Sub BackupWorkbook()
    Dim bakDir As String
    bakDir = ThisWorkbook.Path & "\bak"
    If Dir(bakDir, vbDirectory) = "" Then MkDir bakDir

    Dim stamp As String
    stamp = Format(Now, "yyyymmdd_hhnnss")
    Dim bakPath As String
    bakPath = bakDir & "\ReportMTO.xlsm.bak_" & stamp

    ThisWorkbook.SaveCopyAs bakPath
    modLog.WriteLogEntry Now, "Инфо", "Загрузка пакета", "BackupWorkbook", _
        "Копия книги: " & bakPath
End Sub

' Обновляет значение Power Query-параметра prmSourcePath (Architecture Core §5.1).
'
' (!) ИСПРАВЛЕНО по итогам первого реального прогона в Excel (см. next-steps.md, п.2):
' исходная реализация меняла ThisWorkbook.Names("prmSourcePath").RefersTo - это неверно.
' prmSourcePath - обычная Power Query query (создаётся как «Пустой запрос» с телом
' = "путь", как и fn*-функции, либо как формальный Parameter через Manage Parameters),
' и живёт в коллекции ThisWorkbook.Queries, а не в ThisWorkbook.Names.
Private Sub SetSourcePathParameter(filePath As String)
    Dim safePath As String
    safePath = Replace(filePath, """", """""") ' экранирование кавычек внутри M-строки

    ' v5 (M-1): значение пишется ВМЕСТЕ с meta-записью параметра. Голый строковый литерал
    ' превращал prmSourcePath из ПАРАМЕТРА в обычный запрос, а на обычный запрос реагирует
    ' Formula Firewall: Query-ImportJSON ссылается на другой запрос и при этом сам обращается
    ' к File.Contents - обновление отклоняется. Разжалование происходило при КАЖДОЙ загрузке,
    ' поэтому первый прогон после сборки мог пройти, а следующий - нет.
    On Error GoTo ErrHandler
    ThisWorkbook.Queries("prmSourcePath").Formula = """" & safePath & """" & _
        " meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, , "Не удалось обновить Power Query-параметр 'prmSourcePath' " & _
        "(query с таким именем должна существовать в книге - см. инструкцию по сборке, шаг 4). " & _
        "Исходная ошибка: " & Err.Description
End Sub

Public Sub GenerateReport()
    ' v9.1: повторное нажатие кнопки во время сборки. DoEvents в ShowProgress
    ' пропускает клики, и без этой защиты вторая сборка пошла бы поверх первой.
    If mBusy Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "GenerateReport", _
            "Отчёт уже формируется - повторный запуск пропущен."
        Exit Sub
    End If
    mBusy = True
    On Error GoTo ErrHandler

    ShowProgress 0, "подготовка данных"

    Dim t0 As Single
    t0 = Timer
    modLog.WriteDebug 1, "Формирование отчёта", "GenerateReport", _
        "Старт. Строк в tbDATA: " & SafeRowCount() & "; DEBUG=" & modLog.GetDebugLevel()

    If SafeRowCount() = 0 Then
        modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "GenerateReport", _
            "Таблица tbDATA пуста - сначала загрузите файл выгрузки."
        GoTo CleanExit
    End If

    ' Защитный контракт: обязательные столбцы проверяем один раз до сборки, чтобы
    ' не падать на 13-й секунде с Err -2147221502 «Столбец не найден».
    If Not modContentMTO.ValidateRequiredColumns() Then
        GoTo CleanExit
    End If

    ShowProgress 3, "снимок данных tbDATA"
    modContentMTO.BuildPivots ' пересчёт на случай, если отчёт формируют без предварительной загрузки
    ShowProgress 10, "данные готовы"
    modLog.WriteDebug 1, "Формирование отчёта", "GenerateReport", _
        "BuildPivots завершён: " & Round(Timer - t0, 2) & " c"

    Dim slide3 As String, slide4 As String, slide5 As String
    Dim parsedOK As Boolean
    parsedOK = False

    ' --- Внешний ИИ: неуспех здесь НЕ блокирует выдачу отчёта (Core §11.3) ---
    ' v9.0 (задача 9.5): ошибка проверяется после КАЖДОГО шага и в журнал пишется,
    ' какой именно шаг не прошёл. Раньше одна общая проверка Err на три шага: нет
    ' ключа AI/ENDPOINT - и в журнале «запрос пропущен» без указания причины.
    On Error Resume Next
    Dim endpoint As String, apiKey As String, requestBody As String, responseText As String
    Dim aiStepErr As String
    aiStepErr = ""
    endpoint = GetVariable("AI/ENDPOINT")
    If Err.Number <> 0 Then
        aiStepErr = "шаг 1/3 - ключ AI/ENDPOINT на листе Variable: " & Err.Description
        Err.Clear
    End If
    If aiStepErr = "" Then
        apiKey = ResolveAiApiKey()
        If Err.Number <> 0 Then
            aiStepErr = "шаг 2/3 - ключ API: " & Err.Description
            Err.Clear
        ElseIf Len(apiKey) = 0 Then
            aiStepErr = "шаг 2/3 - ключ API не найден (AI_API_KEY / deepseek.key / Variable)"
        End If
    End If

    ShowProgress 12, "сборка запроса к ИИ"
    If aiStepErr = "" Then
        requestBody = modContentMTO.BuildPrompt()
        If Err.Number <> 0 Then
            aiStepErr = "шаг 3/3 - сборка промпта (ключ AI/MODEL?): " & Err.Description
            Err.Clear
        End If
    End If
    If aiStepErr <> "" Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
            "Запрос к ИИ не отправлен, " & aiStepErr & _
            ". Выводы слайдов будут собраны автоматически по правилам."
    End If

    If aiStepErr = "" And requestBody <> "" Then
        modLog.WriteDebug 1, "Формирование отчёта", "BuildPrompt", _
            "Запрос к ИИ собран: " & Len(requestBody) & " символов"
        modLog.WriteDebug 2, "Формирование отчёта", "BuildPrompt", _
            "Тело запроса: " & Left$(requestBody, 20000)
        ShowProgress 15, "ожидание ответа ИИ (до 60 с, полоса стоит)"
        responseText = modAIGateway.PostJSON(endpoint, apiKey, requestBody)
    Else
        modLog.WriteDebug 1, "Формирование отчёта", "DeepSeek", _
            "Запрос к ИИ пропущен: Err=" & Err.Number & " (" & Err.Description & _
            "), длина тела=" & Len(requestBody)
    End If

    Dim aiErrDesc As String
    If Err.Number <> 0 Then aiErrDesc = Err.Description
    Err.Clear
    On Error GoTo ErrHandler

    modLog.WriteDebug 1, "Формирование отчёта", "DeepSeek", _
        "Ответ ИИ: " & Len(responseText) & " символов" & _
        IIf(aiErrDesc <> "", "; ошибка вызова: " & aiErrDesc, "")
    modLog.WriteDebug 2, "Формирование отчёта", "DeepSeek", _
        "Начало ответа: " & Left$(responseText, 4000)

    If aiErrDesc <> "" Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
            "Слой ИИ пропущен: " & aiErrDesc
    End If

    parsedOK = modContentMTO.ParseAIResponse(responseText, slide3, slide4, slide5)
    If Not parsedOK Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
            "Ответ ИИ не распознан или недоступен - используются заглушки"
    End If

    ' --- Сборка отчёта ---
    ShowProgress 25, "сборка слайдов"
    Dim placeholders As Object
    Set placeholders = modContentMTO.BuildPlaceholders(slide3, slide4, slide5)
    modLog.WriteDebug 1, "Формирование отчёта", "BuildPlaceholders", _
        "Плейсхолдеры готовы: " & Round(Timer - t0, 2) & " c от старта"

    Dim templatePath As String
    templatePath = ThisWorkbook.Path & "\tmp_index.html"

    ShowProgress 94, "сборка HTML"
    Dim html As String
    html = modHTMLEngine.RenderTemplate(templatePath, placeholders)
    modLog.WriteDebug 1, "Формирование отчёта", "RenderTemplate", _
        "HTML собран: " & Len(html) & " символов, " & Round(Timer - t0, 2) & " c от старта"

    Dim resultFolder As String
    resultFolder = GetVariable("OUTPUT/RESULT_FOLDER")

    ShowProgress 97, "сохранение файла"
    Dim savedPath As String
    savedPath = modHTMLEngine.SaveHTMLFile(html, resultFolder)
    modLog.WriteDebug 1, "Формирование отчёта", "SaveHTMLFile", _
        "Папка: " & resultFolder & "; сохранён: " & savedPath

    If savedPath <> "" Then
        modLog.WriteLogEntry Now, "Инфо", "Формирование отчёта", resultFolder, "Сохранён: " & savedPath
    Else
        modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "SaveHTMLFile", _
            "Не удалось сохранить отчёт - проверьте права доступа к папке результата."
    End If

CleanExit:
    modAggregate.EndSnapshot
    EndProgress
    mBusy = False
    Exit Sub

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "GenerateReport", _
        "Ошибка " & Err.Number & ": " & Err.Description
    modLog.WriteDebug 1, "Формирование отчёта", "GenerateReport", _
        "Падение через " & Round(Timer - t0, 2) & " c от старта: " & Err.Number & _
        " (" & Err.Description & ")"
    modAggregate.EndSnapshot
    EndProgress mPrgStage
    mBusy = False
    modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "GenerateReport", _
        "Не удалось сформировать отчёт: " & Err.Description
End Sub

' Отладочный прогон без обращения к внешнему ИИ (A-7): собирает отчёт с гарантированными
' заглушками ИИ-выводов. Нужен для прогона Блоков 1-9 на обезличенном тестовом JSON,
' не тратя токены и не завися от доступности провайдера.
Public Sub DebugGenerateOffline()
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo ErrHandler
    ShowProgress 0, "подготовка данных (отладка, без ИИ)"

    If SafeRowCount() = 0 Then
        modLog.WriteLogEntry Now, "Ошибка", "Отладочный отчёт", "DebugGenerateOffline", _
            "Таблица tbDATA пуста."
        Exit Sub
    End If

    modContentMTO.BuildPivots

    ShowProgress 25, "сборка слайдов"
    Const STUB As String = "[offline] Внешний ИИ не вызывался - отладочный прогон."
    Dim placeholders As Object
    Set placeholders = modContentMTO.BuildPlaceholders(STUB, STUB, STUB)

    Dim html As String
    html = modHTMLEngine.RenderTemplate(ThisWorkbook.Path & "\tmp_index.html", placeholders)

    Dim folder As String
    folder = modHTMLEngine.ResolveOutputFolder(GetVariableDef("OUTPUT/RESULT_FOLDER", "result"))
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        On Error Resume Next
        fso.CreateFolder folder
        Err.Clear
        On Error GoTo ErrHandler
    End If
    Dim path As String
    path = folder & "\debug_" & Format(Now, "yyyymmdd_hhnnss") & ".html"

    ShowProgress 97, "сохранение файла"
    modHTMLEngine.WriteUtf8 path, html
    modAggregate.EndSnapshot
    EndProgress
    mBusy = False

    modLog.WriteLogEntry Now, "Инфо", "Отладочный отчёт", folder, "Сохранён: " & path
    Exit Sub

ErrHandler:
    modAggregate.EndSnapshot
    EndProgress mPrgStage
    mBusy = False
    modLog.WriteLogEntry Now, "Ошибка", "Отладочный отчёт", "DebugGenerateOffline", Err.Description
End Sub

Private Function SafeRowCount() As Long
    On Error Resume Next
    SafeRowCount = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").ListRows.Count
    On Error GoTo 0
End Function

' Резолвер API-ключа ИИ (v7.1): приоритет источников - переменная окружения AI_API_KEY,
' файл %APPDATA%\ReportMTO\deepseek.key (UTF-8), затем лист Variable (legacy, с предупреждением).
' Ключ никогда не пишется в лог и в HTML-отчёт.
Public Function ResolveAiApiKey() As String
    Dim key As String, src As String

    ' 1) Переменная окружения пользователя (setx AI_API_KEY ...).
    key = CleanApiKey(Environ$("AI_API_KEY"))
    If Len(key) > 0 Then src = "переменная окружения AI_API_KEY"

    ' 2) Файл с правами только для владельца (UTF-8; чтение modHTMLEngine.ReadUtf8).
    If Len(key) = 0 Then
        Dim appData As String
        appData = Environ$("APPDATA")
        If Len(appData) > 0 Then
            key = CleanApiKey(ReadKeyFile(appData & "\ReportMTO\deepseek.key"))
            If Len(key) > 0 Then src = "%APPDATA%\ReportMTO\deepseek.key"
        End If
    End If

    ' 3) Лист Variable - небезопасно, но совместимо с прежними книгами.
    If Len(key) = 0 Then
        On Error Resume Next
        key = CleanApiKey(GetVariable("AI/API_KEY"))
        If Err.Number <> 0 Then Err.Clear: key = ""
        On Error GoTo 0
        If Len(key) > 0 Then
            src = "лист Variable"
            modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
                "Ключ ИИ прочитан с листа Variable - защита листа не шифрует; перенесите ключ в AI_API_KEY или %APPDATA%\ReportMTO\deepseek.key"
        End If
    End If

    ' Диагностика без раскрытия секрета: из какого источника взят ключ, его длина и
    ' признак ожидаемого префикса. Без этой строки 401 от провайдера неотличим -
    ' пустой ключ, ключ с мусором и отозванный ключ выглядели в журнале одинаково.
    If Len(key) = 0 Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
            "Ключ ИИ не найден: ни AI_API_KEY, ни %APPDATA%\ReportMTO\deepseek.key, ни лист Variable"
    Else
        Dim pref As String
        If Left$(key, 3) = "sk-" Then pref = "да" Else pref = "нет"
        modLog.WriteDebug 1, "Формирование отчёта", "ResolveAiApiKey", _
            "Источник ключа: " & src & "; длина " & CStr(Len(key)) & _
            " символов; начинается с sk-: " & pref
    End If

    ResolveAiApiKey = key
End Function

' Чистка ключа перед подстановкой в заголовок Authorization: BOM, управляющие
' символы, пробелы и кавычки. Любой из них ломает формат "Bearer sk-..." и даёт
' 401 с текстом про формат заголовка, а не про сам ключ - диагностируется тяжело.
Private Function CleanApiKey(ByVal raw As String) As String
    Dim s As String, i As Long, ch As String, code As Long
    s = ""
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code > 32 And code <> 34 And code <> 39 And code <> 65279 Then s = s & ch
    Next i
    CleanApiKey = s
End Function

' Чтение файла с ключом: отсутствие файла и любая ошибка чтения - пустая строка,
' резолвер идёт к следующему источнику.
Private Function ReadKeyFile(ByVal path As String) As String
    ReadKeyFile = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    On Error Resume Next
    ReadKeyFile = modHTMLEngine.ReadUtf8(path)
    If Err.Number <> 0 Then Err.Clear: ReadKeyFile = ""
    On Error GoTo 0
End Function

' Чтение значения из листа Variable по ключу (например "AI/ENDPOINT").
' Контракт: лист Variable - таблица из двух столбцов, "Key" и "Value" (см. data.md §1.1).
Public Function GetVariable(key As String) As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Variable").ListObjects(1)

    Dim r As ListRow
    For Each r In lo.ListRows
        If CStr(r.Range.Cells(1, 1).Value) = key Then
            GetVariable = CStr(r.Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 1, , "Ключ '" & key & "' не найден на листе Variable"
End Function

' То же, но для необязательных ключей: отсутствие ключа или пустое значение -> defaultValue.
Public Function GetVariableDef(key As String, defaultValue As String) As String
    Dim v As String
    On Error Resume Next
    v = GetVariable(key)
    If Err.Number <> 0 Then Err.Clear: v = ""
    On Error GoTo 0
    If Trim$(v) = "" Then v = defaultValue
    GetVariableDef = v
End Function

' Запись значения в лист Variable. Если ключа нет - он создается в конце таблицы.
Public Sub SetVariable(key As String, value As String)
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Variable").ListObjects(1)
    
    Dim r As ListRow
    For Each r In lo.ListRows
        If CStr(r.Range.Cells(1, 1).Value) = key Then
            r.Range.Cells(1, 2).Value = value
            Exit Sub
        End If
    Next r
    
    ' Ключ не найден - добавляем
    Set r = lo.ListRows.Add
    r.Range.Cells(1, 1).Value = key
    r.Range.Cells(1, 2).Value = value
End Sub

' =====================================================================================
' Прогресс сборки отчёта (v9.1 от 29.09.2026). Два индикатора сразу:
'   - строка состояния Excel: «Отчёт: [####....] 40 % · Слайд 5 · 42 с»;
'   - полоса на листе Main под кнопками (фигуры PRG_BACK / PRG_BAR создаются сами).
' DoEvents после отрисовки: без него Excel во время сборки не перерисовывает окно,
' и прогресс не виден. Ошибка отрисовки сборку не роняет - индикатор вторичен.
' =====================================================================================
Public Sub ShowProgress(ByVal pct As Long, ByVal stage As String)
    Dim nCells As Long, filled As Long, bar As String, secs As Long
    If Not mBusy Then Exit Sub   ' вне сборки отчёта (самотест, сверка) индикатор не нужен
    On Error GoTo Quiet
    mPrgStage = stage
    If pct <= 0 Then mPrgT0 = Timer
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    secs = CLng(Timer - mPrgT0)
    If secs < 0 Then secs = secs + 86400    ' переход через полночь
    nCells = 20
    filled = CLng(pct * nCells / 100)
    ' Replace(Space$), а не String$: так символ вне ANSI гарантированно не искажается.
    bar = Replace$(Space$(filled), " ", ChrW$(&H25A0)) & Replace$(Space$(nCells - filled), " ", ChrW$(&H25A1))
    Application.StatusBar = "Отчёт: " & bar & " " & CStr(pct) & " % " & ChrW$(&HB7) & " " & _
        stage & " " & ChrW$(&HB7) & " " & CStr(secs) & " с"
    DrawProgress pct, stage & " " & ChrW$(&HB7) & " " & CStr(pct) & " % " & ChrW$(&HB7) & _
        " " & CStr(secs) & " с", False
    DoEvents
    Exit Sub
Quiet:
    Err.Clear
End Sub

' Конец сборки: успех - полоса и строка состояния убираются; ошибка - полоса
' остаётся красной с названием этапа, чтобы было видно, где остановилось.
Public Sub EndProgress(Optional ByVal failedStage As String = "")
    On Error GoTo Quiet
    Application.StatusBar = False
    If Len(failedStage) = 0 Then
        HideProgressShapes
    Else
        DrawProgress 100, "Ошибка на этапе: " & failedStage & " " & ChrW$(&H2014) & _
            " подробности в журнале", True
    End If
    Exit Sub
Quiet:
    Err.Clear
End Sub

Private Sub DrawProgress(ByVal pct As Long, ByVal txt As String, ByVal isError As Boolean)
    Dim ws As Worksheet, tr As Shape, br As Shape, lb As Shape, shp As Shape
    Dim topY As Double, leftX As Double
    Const W As Double = 420#
    Const H As Double = 26#
    On Error GoTo Quiet
    Set ws = ThisWorkbook.Worksheets("Main")
    ' Три фигуры: серая дорожка, цветная полоса поверх неё, прозрачная подпись сверху.
    Set tr = FindShape(ws, "PRG_TRACK")
    Set br = FindShape(ws, "PRG_BAR")
    Set lb = FindShape(ws, "PRG_TEXT")
    If tr Is Nothing Then
        ' Место - под самой нижней фигурой листа (кнопками), слева по ней же.
        topY = 10#: leftX = 10#
        For Each shp In ws.Shapes
            If Left$(shp.Name, 4) <> "PRG_" Then
                If shp.Top + shp.Height + 12# > topY Then
                    topY = shp.Top + shp.Height + 12#
                    leftX = shp.Left
                End If
            End If
        Next shp
        Set tr = ws.Shapes.AddShape(1, leftX, topY, W, H)     ' 1 = msoShapeRectangle
        tr.Name = "PRG_TRACK"
        tr.Fill.ForeColor.RGB = RGB(230, 230, 230)
        tr.Line.ForeColor.RGB = RGB(160, 160, 160)
    End If
    If br Is Nothing Then
        Set br = ws.Shapes.AddShape(1, tr.Left, tr.Top, 1#, tr.Height)
        br.Name = "PRG_BAR"
        br.Line.Visible = False
    End If
    If lb Is Nothing Then
        Set lb = ws.Shapes.AddShape(1, tr.Left, tr.Top, tr.Width, tr.Height)
        lb.Name = "PRG_TEXT"
        lb.Fill.Visible = False
        lb.Line.Visible = False
        lb.TextFrame2.TextRange.Font.Size = 10
        lb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
    End If
    tr.Visible = True: br.Visible = True: lb.Visible = True
    br.Left = tr.Left: br.Top = tr.Top: br.Height = tr.Height
    If pct <= 0 Then br.Width = 1# Else br.Width = tr.Width * pct / 100#
    If isError Then
        br.Fill.ForeColor.RGB = RGB(208, 59, 59)
    Else
        br.Fill.ForeColor.RGB = RGB(120, 170, 235)
    End If
    lb.Left = tr.Left: lb.Top = tr.Top: lb.Width = tr.Width: lb.Height = tr.Height
    br.ZOrder 0     ' msoBringToFront: полоса над дорожкой
    lb.ZOrder 0     ' подпись над полосой
    lb.TextFrame2.TextRange.Text = txt
    Exit Sub
Quiet:
    Err.Clear
End Sub

Private Function FindShape(ByVal ws As Worksheet, ByVal nm As String) As Shape
    Set FindShape = Nothing
    On Error Resume Next
    Set FindShape = ws.Shapes(nm)
    Err.Clear
End Function

Private Sub HideProgressShapes()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Main")
    ws.Shapes("PRG_TRACK").Visible = False
    ws.Shapes("PRG_BAR").Visible = False
    ws.Shapes("PRG_TEXT").Visible = False
    Err.Clear
End Sub
