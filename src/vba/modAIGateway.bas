Attribute VB_Name = "modAIGateway"
' modAIGateway - CORE, generic. Только HTTP-транспорт (Architecture Core §7).
' Провайдер/модель/URL/ключ - не хардкодятся, приходят параметрами из листа Variable.
' Текст промпта и разбор ответа - Content Spec (см. modContentMTO.bas: BuildPrompt/ParseAIResponse).
'
' v3.1 (правки по ревью 24.08.2026):
'   P1-15 - Content-Type с charset=utf-8 И отправка тела байтовым массивом UTF-8.
'           MSXML при .Send строки не гарантирует UTF-8 - русскоязычный промпт мог уходить
'           как "????" (тот же класс бага, что №3 в next-steps.md, только на транспорте).
'           Ответ читается из responseBody как UTF-8, а не из responseText (кодировка
'           responseText зависит от заголовков ответа).
Option Explicit

Private Const TIMEOUT_MS As Long = 60000

' Возвращает тело ответа (String) или "" при ошибке/таймауте - вызывающий код (Content Spec)
' сам решает, как деградировать при пустом ответе (см. spec.md §4.4, Partial Failure).
Public Function PostJSON(endpointUrl As String, apiKey As String, bodyJson As String) As String
    On Error GoTo ErrHandler

    Dim t0 As Single
    t0 = Timer
    modLog.WriteDebug 1, "Вызов внешнего ИИ", "PostJSON", _
        "POST " & endpointUrl & "; тело: " & Len(bodyJson) & " символов"

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")  ' позднее связывание - обязательное требование Core

    http.Open "POST", endpointUrl, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetTimeouts TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS

    http.Send TextToUtf8Bytes(bodyJson)

    If http.Status = 200 Then
        PostJSON = Utf8BytesToText(http.responseBody)
        modLog.WriteDebug 1, "Вызов внешнего ИИ", "PostJSON", _
            "HTTP " & http.Status & "; ответ: " & Len(PostJSON) & " символов; " & _
            Round(Timer - t0, 2) & " c"
        modLog.WriteDebug 2, "Вызов внешнего ИИ", "PostJSON", _
            "Начало ответа: " & Left$(PostJSON, 4000)
    Else
        modLog.WriteLogEntry Now, "Предупреждение", "Вызов внешнего ИИ", endpointUrl, _
            "HTTP " & http.Status & ": " & http.statusText
        modLog.WriteDebug 1, "Вызов внешнего ИИ", "PostJSON", _
            "HTTP " & http.Status & ": " & http.statusText
        modLog.WriteDebug 2, "Вызов внешнего ИИ", "PostJSON", _
            "Тело ответа при HTTP-ошибке: " & Left$(Utf8BytesToText(http.responseBody), 4000)
        PostJSON = ""
    End If
    Exit Function

ErrHandler:
    modLog.WriteLogEntry Now, "Предупреждение", "Вызов внешнего ИИ", endpointUrl, _
        "Ошибка/таймаут: " & Err.Description
    modLog.WriteDebug 1, "Вызов внешнего ИИ", "PostJSON", _
        "Ошибка/таймаут через " & Round(Timer - t0, 2) & " c: " & Err.Description
    PostJSON = ""
End Function

Private Function TextToUtf8Bytes(s As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s
    st.Position = 3          ' отсечь BOM
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Open
    st.CopyTo bin
    bin.Position = 0
    TextToUtf8Bytes = bin.Read(-1)
    st.Close: bin.Close
End Function

Private Function Utf8BytesToText(bytes As Variant) As String
    On Error GoTo Fallback
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Open
    bin.Write bytes
    bin.Position = 0
    bin.Type = 2
    bin.Charset = "utf-8"
    Utf8BytesToText = bin.ReadText(-1)
    bin.Close
    Exit Function
Fallback:
    Utf8BytesToText = ""
End Function
