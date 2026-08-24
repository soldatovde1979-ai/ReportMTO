Attribute VB_Name = "modHTMLEngine"
' modHTMLEngine — CORE, generic Template Engine (Architecture Core §9).
' Не знает конкретных имён плейсхолдеров — их список и значения строит Content Spec
' (см. modContentMTO.bas: BuildPlaceholders), контракт — Scripting.Dictionary "{{ИМЯ}}" -> значение.
Option Explicit

Public Function RenderTemplate(templatePath As String, placeholders As Object) As String
    Dim fso As Object, ts As Object, html As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(templatePath, 1, False, -1) ' ForReading, Unicode
    html = ts.ReadAll
    ts.Close

    Dim key As Variant
    For Each key In placeholders.Keys
        html = Replace(html, "{{" & key & "}}", CStr(placeholders(key)))
    Next key

    RenderTemplate = html
End Function

' Возвращает полный путь сохранённого файла или "" при ошибке (см. spec.md §4.5).
Public Function SaveHTMLFile(html As String, resultFolder As String) As String
    On Error GoTo ErrHandler

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(resultFolder) Then
        modLog.WriteLogEntry Now, "Ошибка", "Сохранение отчёта", resultFolder, "Папка недоступна"
        SaveHTMLFile = ""
        Exit Function
    End If

    Dim fileName As String
    fileName = "Report_" & Format(Now, "yyyymmdd_hhnnss") & ".html"

    Dim fullPath As String
    fullPath = fso.BuildPath(resultFolder, fileName)

    Set ts = fso.CreateTextFile(fullPath, True, True) ' Overwrite, Unicode
    ts.Write html
    ts.Close

    SaveHTMLFile = fullPath
    Exit Function

ErrHandler:
    modLog.WriteLogEntry Now, "Ошибка", "Сохранение отчёта", resultFolder, Err.Description
    SaveHTMLFile = ""
End Function
