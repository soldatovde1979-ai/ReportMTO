# modAIGateway

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Только HTTP-транспорт к внешнему LLM (Architecture Core §7). Провайдер/модель/URL/ключ не хардкодятся — приходят параметрами; текст промпта и разбор ответа — Content Spec (`modContentMTO.BuildPrompt`/`ParseAIResponse`).

## Процедуры / Функции

### PostJSON(endpointUrl As String, apiKey As String, bodyJson As String) : String
- Назначение: синхронный POST к чат-API: заголовки `Content-Type: application/json; charset=utf-8` и `Authorization: Bearer <key>`, тело — UTF-8 байтами.
- Вход: `endpointUrl` — URL провайдера; `apiKey` — ключ; `bodyJson` — JSON-тело запроса.
- Выход: тело ответа строкой (только при HTTP 200); `""` при HTTP-ошибке, ошибке транспорта или таймауте — деградацию решает вызывающий код.
- Побочные эффекты: HTTP-запрос к внешнему сервису (таймаут 60 с, без ретраев); записи в журнал `modLog` (уровни 1–2): статус, размеры, тайминги, начало ответа; при не-200 — «Предупреждение». Секреты не логируются.

### TextToUtf8Bytes(s As String) : Variant (private)
- Назначение: перевести строку в массив байтов UTF-8 без BOM через ADODB.Stream.
- Вход: `s` — текст.
- Выход: Variant-массив байтов.
- Побочные эффекты: нет.

### Utf8BytesToText(bytes As Variant) : String (private)
- Назначение: перевести байты UTF-8 в строку.
- Вход: `bytes` — массив байтов ответа.
- Выход: строка; `""` при ошибке.
- Побочные эффекты: нет.
