Attribute VB_Name = "modAIGateway"
' modAIGateway — CORE, generic. Только HTTP-транспорт (Architecture Core §7).
' Провайдер/модель/URL/ключ — не хардкодятся, приходят параметрами из листа Variable.
' Текст промпта и разбор ответа — Content Spec (см. modContentMTO.bas: BuildPrompt/ParseAIResponse).
Option Explicit

Private Const TIMEOUT_MS As Long = 60000

' Возвращает тело ответа (String) или "" при ошибке/таймауте — вызывающий код (Content Spec)
' сам решает, как деградировать при пустом ответе (см. spec.md §4.4, Partial Failure).
Public Function PostJSON(endpointUrl As String, apiKey As String, bodyJson As String) As String
    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")  ' позднее связывание — обязательное требование Core

    http.Open "POST", endpointUrl, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & apiKey
    http.SetTimeouts TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS, TIMEOUT_MS

    http.Send bodyJson

    If http.Status = 200 Then
        PostJSON = http.responseText
    Else
        modLog.WriteLogEntry Now, "Предупреждение", "Вызов внешнего ИИ", endpointUrl, _
            "HTTP " & http.Status & ": " & http.statusText
        PostJSON = ""
    End If
    Exit Function

ErrHandler:
    modLog.WriteLogEntry Now, "Предупреждение", "Вызов внешнего ИИ", endpointUrl, _
        "Ошибка/таймаут: " & Err.Description
    PostJSON = ""
End Function
