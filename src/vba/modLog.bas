Attribute VB_Name = "modLog"
' modLog - CORE, generic. Схема Logs фиксирована контрактом (data.md §2.3):
' Дата | Тип записи | Действие | Источник | Результат.
'
' v3.1: запись лога не может уронить вызывающий код (отсутствие листа Logs, защита листа
' и т.п.) - иначе обработчик ошибки в modMain сам падал бы при попытке залогировать ошибку.
'
' v7.2 (08.09.2026) - расширенное логирование:
'   Ключ DEBUG на листе Variable: 0 - прежнее поведение, 1 - краткие этапы/тайминги,
'   2 - очень полно (полный промпт и начало ответа ИИ; API-ключ не пишется никогда).
'   Каждая запись WriteLogEntry/WriteDebug дополнительно дописывается во внешний файл
'   ReportMTO_log.txt рядом с книгой (UTF-8, append). Лист Logs сохраняет обрезку result
'   до 1000 символов (контракт data.md §2.3), в файл идёт полный текст (до 100000).
'
' v8.2 (13.09.2026) - новая семантика DEBUG (ТЗ v1.0):
'   Лист Logs всегда только «Веха» и «Ошибка»; внешний файл ReportMTO.log:
'   DEBUG=0 - файл не пишется, DEBUG=1 - «Ошибка»+«Веха», DEBUG=2 - все записи.
Option Explicit

Private Const TXT_LIMIT As Long = 100000

' Уровень отладки из листа Variable (ключ "DEBUG": 0/1/2). Ключа нет или сбой чтения -> 0.
Public Function GetDebugLevel() As Long
    On Error Resume Next
    GetDebugLevel = CLng(Val(modMain.GetVariableDef("DEBUG", "0")))
    On Error GoTo 0
    If GetDebugLevel < 0 Or GetDebugLevel > 2 Then GetDebugLevel = 0
End Function

' Отладочная запись: пишется, только если DEBUG >= minLevel. Тип "Отладка" в лист
' Logs не пишется (лист - только вехи и ошибки); во внешний файл попадает при DEBUG=2.
Public Sub WriteDebug(minLevel As Long, action As String, source As String, message As String)
    If GetDebugLevel() < minLevel Then Exit Sub
    WriteLogEntry Now, "Отладка", action, source, message
End Sub

Public Sub WriteLogEntry(dt As Date, entryType As String, action As String, source As String, result As String)
    On Error GoTo Silent

    ' Фильтр листа (ТЗ v1.0, п.1.2): в tbLogs попадают только «Ошибка» и «Веха».
    ' «Инфо», «Предупреждение», «Отладка» в лист не пишутся никогда.
    If entryType = "Ошибка" Or entryType = "Веха" Then
        On Error Resume Next   ' сбой листа не должен отменять запись во внешний файл

        Dim lo As ListObject
        Set lo = ThisWorkbook.Sheets("Logs").ListObjects("tbLogs")

        Dim newRow As ListRow
        Set newRow = lo.ListRows.Add

        With newRow.Range
            .Cells(1, 1).Value = dt
            .Cells(1, 2).Value = entryType
            .Cells(1, 3).Value = action
            .Cells(1, 4).Value = source
            .Cells(1, 5).Value = Left$(result, 1000)   ' защита от гигантских Err.Description
        End With

        If Err.Number <> 0 Then Debug.Print "modLog: не удалось записать в Logs - " & Err.Description
        On Error GoTo Silent
    End If

    ' Файл пишется независимо от листа (ТЗ v1.0, п.1.3): «Инфо»-записи при DEBUG=2
    ' должны попадать в файл, даже если в лист они не пишутся.
    AppendToFile dt, entryType, action, source, result
    Exit Sub

Silent:
    ' Логирование - вспомогательная функция; её отказ не должен прерывать основной сценарий.
    Debug.Print "modLog: не удалось записать в Logs - " & Err.Description
End Sub

' Обёртка для вех: запись типа «Веха» (ТЗ v1.0). Веха попадает и в лист Logs,
' и во внешний файл при DEBUG>=1.
Public Sub WriteMilestone(action As String, source As String, message As String)
    WriteLogEntry Now, "Веха", action, source, message
End Sub

' Дописывает строку во внешний журнал ReportMTO.log рядом с книгой (UTF-8, append).
' Отказ файловой операции молча игнорируется: файл - зеркало листа Logs, не источник истины.
' Файл создаётся без BOM: Notepad и редакторы читают UTF-8 без BOM корректно.
Private Sub AppendToFile(dt As Date, entryType As String, action As String, source As String, result As String)
    On Error GoTo Done

    If Len(ThisWorkbook.Path) = 0 Then Exit Sub   ' книга не сохранена - некуда писать

    ' Фильтр внешнего файла (ТЗ v1.0, п.1.2):
    '   DEBUG=0 - файл не пишется;
    '   DEBUG=1 - только «Ошибка» и «Веха»;
    '   DEBUG=2 - все записи (включая «Инфо», «Предупреждение», «Отладка»).
    Dim lvl As Long
    lvl = GetDebugLevel()
    If lvl = 0 Then Exit Sub
    If (entryType = "Инфо" Or entryType = "Предупреждение" Or entryType = "Отладка") And lvl < 2 Then Exit Sub

    Dim path As String
    path = ThisWorkbook.Path & "\ReportMTO.log"

    Dim lineText As String
    lineText = Format(dt, "yyyy-mm-dd hh:nn:ss") & " | " & entryType & " | " & action & _
               " | " & source & " | " & Left$(result, TXT_LIMIT) & vbCrLf

    ' Байты строки UTF-8 без BOM (паттерн TextToUtf8Bytes из modAIGateway):
    ' WriteText ставит BOM в начало текстового потока - отсекаем его, чтобы
    ' не вставить BOM в середину файла при дозаписи.
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText lineText
    st.Position = 0
    Dim raw As Object
    Set raw = CreateObject("ADODB.Stream")
    raw.Type = 1
    raw.Open
    st.CopyTo raw
    st.Close
    If raw.Size >= 3 Then raw.Position = 3
    Dim newBytes As Variant
    newBytes = raw.Read(-1)
    raw.Close

    ' Дозапись в конец существующего файла (или создание нового).
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    If Len(Dir$(path)) > 0 Then bin.LoadFromFile path
    bin.Position = bin.Size
    bin.Write newBytes
    bin.SaveToFile path, 2   ' adSaveCreateOverWrite
    bin.Close

Done:
End Sub
