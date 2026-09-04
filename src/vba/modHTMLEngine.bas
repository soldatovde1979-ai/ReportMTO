Attribute VB_Name = "modHTMLEngine"
' modHTMLEngine - CORE, generic Template Engine (Architecture Core §9).
' Не знает конкретных имён плейсхолдеров - их список и значения строит Content Spec
' (см. modContentMTO.bas: BuildPlaceholders), контракт - Scripting.Dictionary "{{ИМЯ}}" -> значение.
'
' v3.1 (правки по ревью 24.08.2026):
'   P0-3  - чтение и запись строго в UTF-8 через ADODB.Stream. Было: OpenTextFile(..., -1)
'           = UTF-16LE на UTF-8-шаблоне (кириллица разрушалась) и CreateTextFile(..., True)
'           = UTF-16 на выходе.
'   P1-8  - добавлена HtmlEscape: платформенная функция, применяется Content Spec ко всем
'           значениям, попадающим в HTML (в т.ч. к тексту от внешнего ИИ).
'   P1-11 - ResolveOutputFolder: разворачивает %проект%, .\, ..\ и Environ-переменные
'           относительно ThisWorkbook.Path.
Option Explicit

Public Function RenderTemplate(templatePath As String, placeholders As Object) As String
    Dim html As String
    html = ReadUtf8(templatePath)

    Dim key As Variant
    For Each key In placeholders.Keys
        html = Replace(html, "{{" & key & "}}", CStr(placeholders(key)))
    Next key

    RenderTemplate = html
End Function

' Возвращает полный путь сохранённого файла или "" при ошибке (см. spec.md §4.5).
Public Function SaveHTMLFile(html As String, resultFolder As String) As String
    On Error GoTo ErrHandler

    Dim folder As String
    folder = ResolveOutputFolder(resultFolder)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folder) Then
        ' Пытаемся создать один уровень - типичный случай "папки result\ ещё нет".
        On Error Resume Next
        fso.CreateFolder folder
        On Error GoTo ErrHandler
    End If

    If Not fso.FolderExists(folder) Then
        modLog.WriteLogEntry Now, "Ошибка", "Сохранение отчёта", folder, _
            "Папка недоступна (исходное значение OUTPUT/RESULT_FOLDER: " & resultFolder & ")"
        SaveHTMLFile = ""
        Exit Function
    End If

    Dim fileName As String
    fileName = "Report_" & Format(Now, "yyyymmdd_hhnnss") & ".html"

    Dim fullPath As String
    fullPath = fso.BuildPath(folder, fileName)

    WriteUtf8 fullPath, html

    SaveHTMLFile = fullPath
    Exit Function

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Сохранение отчёта", resultFolder, Err.Description
    SaveHTMLFile = ""
End Function

' =====================================================================================
' UTF-8 ввод/вывод (P0-3)
' =====================================================================================
Public Function ReadUtf8(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)  ' adReadAll
    st.Close
End Function

' Пишет UTF-8 БЕЗ BOM: ADODB.Stream в текстовом режиме всегда добавляет BOM,
' поэтому первые 3 байта отсекаются переносом в бинарный поток.
Public Sub WriteUtf8(path As String, content As String)
    Dim stText As Object, stBin As Object
    Set stText = CreateObject("ADODB.Stream")
    stText.Type = 2
    stText.Charset = "utf-8"
    stText.Open
    stText.WriteText content
    stText.Position = 3         ' пропустить BOM (EF BB BF)

    Set stBin = CreateObject("ADODB.Stream")
    stBin.Type = 1              ' adTypeBinary
    stBin.Open

    stText.CopyTo stBin
    stBin.SaveToFile path, 2    ' adSaveCreateOverWrite

    stText.Close
    stBin.Close
End Sub

' =====================================================================================
' Экранирование HTML (P1-8) - применяется ко ВСЕМ подставляемым значениям:
' полям из 1С (post/zn_type/ФИО могут содержать & и <) и тексту внешнего ИИ.
' =====================================================================================
Public Function HtmlEscape(s As String) As String
    Dim t As String
    t = Replace(s, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&#39;")
    HtmlEscape = t
End Function

' =====================================================================================
' Резолв пути выдачи (P1-11)
' Поддерживаются: %проект% и %project% (папка книги), .\ и ..\ (относительно книги),
' Environ-переменные вида %TEMP%, а также обычный абсолютный путь.
' =====================================================================================
Public Function ResolveOutputFolder(raw As String) As String
    Dim s As String
    s = Trim$(raw)
    If s = "" Then s = ".\result"

    s = Replace(s, "%проект%", ThisWorkbook.Path)
    s = Replace(s, "%ПРОЕКТ%", ThisWorkbook.Path)
    s = Replace(s, "%project%", ThisWorkbook.Path)

    ' Environ-переменные (%TEMP%, %USERPROFILE% и т.п.)
    Do While InStr(s, "%") > 0
        Dim p1 As Long, p2 As Long
        p1 = InStr(s, "%")
        p2 = InStr(p1 + 1, s, "%")
        If p2 = 0 Then Exit Do
        Dim varName As String, varVal As String
        varName = Mid$(s, p1 + 1, p2 - p1 - 1)
        varVal = Environ$(varName)
        If varVal = "" Then Exit Do   ' неизвестный токен - оставляем как есть, не зацикливаемся
        s = Left$(s, p1 - 1) & varVal & Mid$(s, p2 + 1)
    Loop

    ' Относительные пути - от папки книги.
    If Left$(s, 2) = ".\" Or Left$(s, 3) = "..\" Then
        s = ThisWorkbook.Path & "\" & s
    ElseIf InStr(s, ":") = 0 And Left$(s, 2) <> "\\" Then
        s = ThisWorkbook.Path & "\" & s
    End If

    ' Нормализация: убрать хвостовой слэш и свернуть .\ / ..\
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    s = fso.GetAbsolutePathName(s)
    On Error GoTo 0

    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    ResolveOutputFolder = s
End Function
