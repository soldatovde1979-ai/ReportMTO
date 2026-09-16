# modAIGateway

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
HTTP-шлюз к ИИ-провайдеру: отправка JSON-промпта и получение ответа без внешних библиотек (WinHttp + ADODB.Stream).

## Процедуры / Функции

### PostJSON(endpointUrl As String, apiKey As String, bodyJson As String) : String
- Вход: URL endpoint, API-ключ, тело запроса (JSON, UTF-8)
- Выход: тело ответа текстом ("" при ошибке — вызывающий модуль решает, как деградировать)
- Побочные эффекты: сетевой вызов к внешнему провайдеру ИИ; API-ключ не логируется

### TextToUtf8Bytes(s As String) : Variant
- Вход: строка
- Выход: массив байтов UTF-8
- Побочные эффекты: нет побочных эффектов

### Utf8BytesToText(bytes As Variant) : String
- Вход: массив байтов UTF-8
- Выход: строка (fallback — ANSI, если UTF-8 не удался)
- Побочные эффекты: нет побочных эффектов
