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
Option Explicit

Public Sub LoadSourceFile()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("JSON files (*.json), *.json", , "Выберите файл выгрузки")
    If filePath = False Then Exit Sub

    Dim rowsBefore As Long
    rowsBefore = SafeRowCount()

    On Error GoTo ErrHandler

    SetSourcePathParameter CStr(filePath)
    modPQSync.RefreshImportQuery

    Dim rowsAfter As Long
    rowsAfter = SafeRowCount()

    modLog.WriteLogEntry Now, "Инфо", "Загрузка данных", Dir(CStr(filePath)), _
        "Строк до: " & rowsBefore & "; строк после: " & rowsAfter

    modContentMTO.BuildPivots
    MsgBox "Загрузка завершена. Строк в tbDATA: " & rowsAfter, vbInformation
    Exit Sub

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Загрузка данных", Dir(CStr(filePath)), Err.Description
    ' v5 (M-2): прежний текст утверждал, что виноват файл. Большинство отказов на этом пути
    ' к файлу отношения не имеют (подключение не найдено, Formula Firewall, обрыв на типизации),
    ' и формулировка уводила от причины.
    MsgBox "Загрузка не выполнена." & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Если в тексте выше упоминается Formula.Firewall - разово выключите уровни " & _
           "конфиденциальности: Данные -> Получить данные -> Параметры запроса -> Конфиденциальность." & vbCrLf & _
           "Подробности записаны на лист Logs.", vbCritical
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
    On Error GoTo ErrHandler

    Application.StatusBar = "Подготовка данных..."

    If SafeRowCount() = 0 Then
        MsgBox "Таблица tbDATA пуста - сначала нажмите «Загрузить» и выберите файл выгрузки.", vbExclamation
        GoTo CleanExit
    End If

    modContentMTO.BuildPivots ' пересчёт на случай, если отчёт формируют без предварительной загрузки

    Dim slide3 As String, slide4 As String, slide5 As String
    Dim parsedOK As Boolean
    parsedOK = False

    ' --- Внешний ИИ: неуспех здесь НЕ блокирует выдачу отчёта (Core §11.3) ---
    On Error Resume Next
    Dim endpoint As String, apiKey As String, requestBody As String, responseText As String
    endpoint = GetVariable("AI/ENDPOINT")
    apiKey = GetVariable("AI/API_KEY")

    Application.StatusBar = "Формирование запроса к ИИ..."
    requestBody = modContentMTO.BuildPrompt()

    If Err.Number = 0 And requestBody <> "" Then
        Application.StatusBar = "Ожидание ответа внешнего ИИ (до 60 сек)..."
        responseText = modAIGateway.PostJSON(endpoint, apiKey, requestBody)
    End If

    Dim aiErrDesc As String
    If Err.Number <> 0 Then aiErrDesc = Err.Description
    Err.Clear
    On Error GoTo ErrHandler

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
    Application.StatusBar = "Сборка HTML-отчёта..."
    Dim placeholders As Object
    Set placeholders = modContentMTO.BuildPlaceholders(slide3, slide4, slide5)

    Dim templatePath As String
    templatePath = ThisWorkbook.Path & "\tmp_index.html"

    Dim html As String
    html = modHTMLEngine.RenderTemplate(templatePath, placeholders)

    Dim resultFolder As String
    resultFolder = GetVariable("OUTPUT/RESULT_FOLDER")

    Dim savedPath As String
    savedPath = modHTMLEngine.SaveHTMLFile(html, resultFolder)

    If savedPath <> "" Then
        modLog.WriteLogEntry Now, "Инфо", "Формирование отчёта", resultFolder, "Сохранён: " & savedPath
        MsgBox "Отчёт сохранён: " & savedPath, vbInformation
    Else
        MsgBox "Не удалось сохранить отчёт - проверьте права доступа к папке результата.", vbCritical
    End If

CleanExit:
    modAggregate.EndSnapshot
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Формирование отчёта", "GenerateReport", _
        "Ошибка " & Err.Number & ": " & Err.Description
    modAggregate.EndSnapshot
    Application.StatusBar = False
    MsgBox "Не удалось сформировать отчёт." & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Подробности записаны на лист Logs.", vbCritical
End Sub

' Отладочный прогон без обращения к внешнему ИИ (A-7): собирает отчёт с гарантированными
' заглушками ИИ-выводов. Нужен для прогона Блоков 1-9 на обезличенном тестовом JSON,
' не тратя токены и не завися от доступности провайдера.
Public Sub DebugGenerateOffline()
    On Error GoTo ErrHandler

    If SafeRowCount() = 0 Then
        MsgBox "Таблица tbDATA пуста.", vbExclamation
        Exit Sub
    End If

    modContentMTO.BuildPivots

    Const STUB As String = "[offline] Внешний ИИ не вызывался - отладочный прогон."
    Dim placeholders As Object
    Set placeholders = modContentMTO.BuildPlaceholders(STUB, STUB, STUB)

    Dim html As String
    html = modHTMLEngine.RenderTemplate(ThisWorkbook.Path & "\tmp_index.html", placeholders)

    Dim folder As String
    folder = modHTMLEngine.ResolveOutputFolder(GetVariable("OUTPUT/RESULT_FOLDER"))
    Dim path As String
    path = folder & "\debug_" & Format(Now, "yyyymmdd_hhnnss") & ".html"

    modHTMLEngine.WriteUtf8 path, html
    modAggregate.EndSnapshot

    modLog.WriteLogEntry Now, "Инфо", "Отладочный отчёт", folder, "Сохранён: " & path
    MsgBox "Отладочный отчёт (без ИИ) сохранён: " & path, vbInformation
    Exit Sub

ErrHandler:
    modAggregate.EndSnapshot
    modLog.WriteLogEntry Now, "Ошибка", "Отладочный отчёт", "DebugGenerateOffline", Err.Description
    MsgBox "Ошибка отладочного прогона: " & Err.Description, vbCritical
End Sub

Private Function SafeRowCount() As Long
    On Error Resume Next
    SafeRowCount = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").ListRows.Count
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
