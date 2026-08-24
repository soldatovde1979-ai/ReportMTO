Attribute VB_Name = "modMain"
' modMain — CORE, оркестрация (Architecture Core, сценарии §11). Обработчики кнопок «Загрузить» /
' «Сформировать отчёт» на листе Main. Вызывает контрактные функции по фиксированным именам —
' их реализация для МТО находится в modContentMTO.bas:
'   modContentMTO.BuildPivots()               — построение всех аналитических блоков
'   modContentMTO.BuildPrompt() As String      — тело запроса к внешнему ИИ (JSON)
'   modContentMTO.ParseAIResponse(...)         — разбор ответа ИИ
'   modContentMTO.BuildPlaceholders() As Object — словарь {{ИМЯ}} -> значение для шаблона
Option Explicit

Public Sub LoadSourceFile()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename("JSON files (*.json), *.json", , "Выберите файл выгрузки")
    If filePath = False Then Exit Sub

    Dim rowsBefore As Long
    rowsBefore = SafeRowCount()

    SetSourcePathParameter CStr(filePath)

    On Error GoTo ErrHandler
    modPQSync.RefreshImportQuery
    On Error GoTo 0

    Dim rowsAfter As Long
    rowsAfter = SafeRowCount()

    modLog.WriteLogEntry Now, "Инфо", "Загрузка данных", Dir(CStr(filePath)), _
        "Строк до: " & rowsBefore & "; строк после: " & rowsAfter

    modContentMTO.BuildPivots
    MsgBox "Загрузка завершена. Строк в tbDATA: " & rowsAfter, vbInformation
    Exit Sub

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Загрузка данных", Dir(CStr(filePath)), Err.Description
    MsgBox "Ошибка загрузки: файл повреждён или имеет неверный формат." & vbCrLf & Err.Description, vbCritical
End Sub

' Обновляет значение Power Query-параметра prmSourcePath (Architecture Core §5.1).
'
' ⚠️ ИСПРАВЛЕНО по итогам первого реального прогона в Excel (см. next-steps.md, п.2):
' исходная реализация меняла ThisWorkbook.Names("prmSourcePath").RefersTo — это неверно.
' prmSourcePath — обычная Power Query query (создаётся как "Пустой запрос" с телом
' = "путь", как и fn*-функции, либо как формальный Parameter через Manage Parameters),
' и живёт в коллекции ThisWorkbook.Queries, а не в ThisWorkbook.Names. Установка через
' Names не влияла на то, что реально читает Query-ImportJSON при File.Contents(prmSourcePath).
Private Sub SetSourcePathParameter(filePath As String)
    Dim safePath As String
    safePath = Replace(filePath, """", """""") ' экранирование кавычек внутри M-строки

    On Error GoTo ErrHandler
    ThisWorkbook.Queries("prmSourcePath").Formula = """" & safePath & """"
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, , "Не удалось обновить Power Query-параметр 'prmSourcePath' " & _
        "(query с таким именем должна существовать в книге — см. инструкцию по сборке, шаг 4). " & _
        "Исходная ошибка: " & Err.Description
End Sub

Public Sub GenerateReport()
    modContentMTO.BuildPivots ' пересчёт на случай, если отчёт формируют без предварительной загрузки

    Dim endpoint As String, apiKey As String
    endpoint = GetVariable("AI/ENDPOINT")
    apiKey = GetVariable("AI/API_KEY")

    Dim requestBody As String
    requestBody = modContentMTO.BuildPrompt()

    Dim responseText As String
    responseText = modAIGateway.PostJSON(endpoint, apiKey, requestBody)

    Dim slide3 As String, slide4 As String, slide5 As String
    Dim parsedOK As Boolean
    parsedOK = modContentMTO.ParseAIResponse(responseText, slide3, slide4, slide5)

    If Not parsedOK Then
        modLog.WriteLogEntry Now, "Предупреждение", "Формирование отчёта", "DeepSeek", _
            "Ответ ИИ не распознан или недоступен — используются заглушки"
    End If

    Dim placeholders As Object
    Set placeholders = modContentMTO.BuildPlaceholders() ' уже содержит AI_INSIGHT_SLIDE_3/4/5 с учётом заглушек

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
        MsgBox "Не удалось сохранить отчёт — проверьте права доступа к папке результата.", vbCritical
    End If
End Sub

Private Function SafeRowCount() As Long
    On Error Resume Next
    SafeRowCount = ThisWorkbook.Sheets("tbDATA").ListObjects("tbDATA").ListRows.Count
    On Error GoTo 0
End Function

' Чтение значения из листа Variable по ключу (например "AI/ENDPOINT").
' Контракт: лист Variable — таблица из двух столбцов, "Key" и "Value" (см. data.md §1.1).
Public Function GetVariable(key As String) As String
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("Variable").ListObjects(1)

    Dim r As ListRow
    For Each r In lo.ListRows
        If r.Range.Cells(1, 1).Value = key Then
            GetVariable = CStr(r.Range.Cells(1, 2).Value)
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 1, , "Ключ '" & key & "' не найден на листе Variable"
End Function
