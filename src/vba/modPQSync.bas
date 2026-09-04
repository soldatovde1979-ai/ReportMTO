Attribute VB_Name = "modPQSync"
' modPQSync - CORE, generic. Синхронное обновление импортного запроса Power Query.
'
' v5 (02.09.2026), две правки:
'   S-1 - имя подключения больше не хардкодится. Было:
'           Set conn = ThisWorkbook.Connections("Query - ImportJSON")
'         Excel формирует имя подключения как "Query - " & <имя запроса>, а запрос называется
'         "Query-ImportJSON" (так его создаёт Build-ReportMTO.ps1) - то есть по умолчанию
'         подключение называется "Query - Query-ImportJSON". Совпадение имени обеспечивал
'         только шаг 4 сборочного скрипта, который сам переименовывает подключение, - но этот
'         шаг помечен в скрипте как экспериментальный и ни разу не проверенный, а при ручной
'         загрузке ("Закрыть и загрузить в... -> Таблица") переименование делает человек.
'         Промах по имени давал run-time error 9 на ПЕРВОЙ строке процедуры - до чтения файла.
'         Пользователь при этом видел "файл повреждён или имеет неверный формат" и искал
'         проблему в JSON.
'         Теперь основной путь - обновление QueryTable самой таблицы tbDATA (имя таблицы
'         зафиксировано контрактом Core, в отличие от имени подключения), поиск подключения
'         по имени остался запасным и идёт по вхождению подстроки.
'   S-2 - conn.OLEDBConnection вызывался безусловно. У подключения не-OLEDB типа (модель
'         данных, ODBC) обращение к этому свойству бросает исключение - добавлена проверка
'         conn.Type. Тот же класс ошибки уже ловили в Build-ReportMTO.ps1.
'
' (!) НЕ ПРОВЕРЕНО В EXCEL - у меня нет доступа к книге.
Option Explicit

Private Const PQ_QUERY_NAME As String = "Query-ImportJSON"
Private Const TARGET_TABLE  As String = "tbDATA"

Public Sub RefreshImportQuery()
    ' --- Путь 1 (основной): обновляем QueryTable целевой таблицы. Имя подключения не нужно.
    Dim lo As ListObject
    Set lo = FindListObjectAnywhere(TARGET_TABLE)

    If Not lo Is Nothing Then
        Dim qt As QueryTable
        On Error Resume Next
        Set qt = lo.QueryTable          ' у обычной, не внешней таблицы свойства нет - будет Nothing
        On Error GoTo 0
        If Not qt Is Nothing Then
            qt.BackgroundQuery = False  ' без этого .Refresh вернёт управление до окончания
            qt.Refresh BackgroundQuery:=False
            Exit Sub
        End If
    End If

    ' --- Путь 2 (запасной): подключение по вхождению имени запроса.
    Dim conn As WorkbookConnection
    Set conn = FindConnectionByQueryName(PQ_QUERY_NAME)

    If conn Is Nothing Then
        Err.Raise vbObjectError + 513, "modPQSync.RefreshImportQuery", _
            "Запрос '" & PQ_QUERY_NAME & "' не выгружен на лист '" & TARGET_TABLE & "'. " & _
            "В редакторе Power Query: Закрыть и загрузить в... -> Таблица -> лист " & TARGET_TABLE & ". " & _
            "Пока этот шаг не сделан, обновлять нечего."
    End If

    If conn.Type = xlConnectionTypeOLEDB Then
        conn.OLEDBConnection.BackgroundQuery = False
    End If
    conn.Refresh
End Sub

' Ищет ListObject по имени на всех листах книги: коллекции ThisWorkbook.ListObjects не существует.
Private Function FindListObjectAnywhere(tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectAnywhere = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindListObjectAnywhere = Nothing
End Function

' Подключение Power Query Excel называет "Query - <имя запроса>", но точное имя зависит от
' того, как именно запрос выгружался. Сравниваем по вхождению, а не по равенству.
Private Function FindConnectionByQueryName(queryName As String) As WorkbookConnection
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        If InStr(1, conn.Name, queryName, vbTextCompare) > 0 Then
            Set FindConnectionByQueryName = conn
            Exit Function
        End If
    Next conn
    Set FindConnectionByQueryName = Nothing
End Function
