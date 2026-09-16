# modAggregate

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Агрегации по снимку данных из ListObject в памяти: счёт, distinct-счёт, среднее, перцентили, уникальные значения, сортировки. Не знает про поля направления — имена столбцов принимает параметрами. Заменяет PivotTable там, где тот без Data Model не справляется (Distinct Count, % от группы, топ-N, попарные сравнения). Все агрегации возвращают `Scripting.Dictionary`: `"знач1|знач2|"` → число.

## Процедуры / Функции

### BeginSnapshot(lo As ListObject)
- Назначение: снять данные таблицы в память (Value2) и построить карту «имя столбца → индекс»; вызывается один раз перед серией агрегаций.
- Вход: `lo` — ListObject-источник (например `tbDATA`).
- Выход: нет.
- Побочные эффекты: заполняет модульные переменные снимка; пустая таблица не роняет модуль (`mRows = 0`); `WriteDebug 2` с составом столбцов.

### EndSnapshot()
- Назначение: освободить память снимка (необязательно вызывать).
- Вход: нет. Выход: нет.
- Побочные эффекты: сброс модульных переменных.

### IsReady() : Boolean
- Назначение: признак созданного снимка.
- Вход: нет. Выход: True/False. Побочные эффекты: нет.

### RowCount() : Long
- Назначение: число строк снимка.
- Вход: нет. Выход: количество строк. Побочные эффекты: нет.

### ColIndex(columnName As String) : Long
- Назначение: индекс столбца по имени.
- Вход: `columnName` — имя столбца (сравнение без регистра).
- Выход: 1-based индекс; при отсутствии снимка или столбца — raise (`vbObjectError+3`/`+2`).
- Побочные эффекты: `WriteDebug 1` при отсутствующем столбце.

### HasColumn(columnName As String) : Boolean
- Назначение: проверка наличия столбца.
- Вход: `columnName` — имя столбца. Выход: True/False. Побочные эффекты: нет.

### CellText(r As Long, columnName As String) : String
- Назначение: значение ячейки снимка как нормализованный текст (Boolean → `True`/`False` независимо от локали).
- Вход: `r` — строка; `columnName` — столбец.
- Выход: текст. Побочные эффекты: нет.

### CellRaw(r As Long, columnName As String) : Variant
- Назначение: сырое значение ячейки снимка.
- Вход: `r` — строка; `columnName` — столбец.
- Выход: Variant-значение. Побочные эффекты: нет.

### NormText(v As Variant) : String (private)
- Назначение: нормализация значения к тексту: Empty/Null/Error → `""`, Boolean → `"True"`/`"False"`.
- Вход: `v` — значение. Выход: текст. Побочные эффекты: нет.

### NormBoolLiteral(s As String) : String (private)
- Назначение: привести написания булевых литералов (`TRUE`/`ИСТИНА`/`1`/`ДА` и др.) к `True`/`False`.
- Вход: `s` — строка. Выход: нормализованный литерал. Побочные эффекты: нет.

### GroupCount(groupCols As Variant, Optional filters As Variant) : Object
- Назначение: количество строк по группировке.
- Вход: `groupCols` — массив имён столбцов группировки; `filters` — массив фильтров `"col=знач"` / `"col<>знач"` / `"col<>"` (непусто) / `"col@=a;b;c"` (в списке), соединяются через И.
- Выход: Dictionary `"знач1|знач2|" → число`.
- Побочные эффекты: нет.

### GroupCountDistinct(groupCols As Variant, distinctCol As String, Optional filters As Variant) : Object
- Назначение: количество уникальных значений столбца по группам.
- Вход: `groupCols` — группировка; `distinctCol` — столбец уникальности; `filters` — фильтры.
- Выход: Dictionary `ключ группы → число уникальных`.
- Побочные эффекты: нет.

### GroupAverage(groupCols As Variant, valueCol As String, Optional filters As Variant) : Object
- Назначение: среднее числового столбца по группам (пустые/нечисловые пропускаются).
- Вход: `groupCols`, `valueCol`, `filters`.
- Выход: Dictionary `ключ группы → среднее`; группы без значений отсутствуют.
- Побочные эффекты: нет.

### GroupPercentile(groupCols As Variant, valueCol As String, p As Double, Optional filters As Variant) : Object
- Назначение: перцентиль числового столбца по группам (метод линейной интерполяции EXCEL.PERCENTILE.INC).
- Вход: `groupCols`, `valueCol`, `p` (0..1), `filters`.
- Выход: Dictionary `ключ группы → перцентиль`; группы без значений в результат не включаются.
- Побочные эффекты: нет; `p` вне 0..1 — raise `vbObjectError+5`.

### Percentile(valueCol As String, p As Double, Optional filters As Variant, Optional ByRef hasValue As Boolean) : Double
- Назначение: перцентиль по всему отфильтрованному набору без группировки.
- Вход: `valueCol`, `p`, `filters`, `hasValue` — ByRef-флаг наличия значений.
- Выход: перцентиль; при отсутствии значений 0 и `hasValue = False`.
- Побочные эффекты: нет.

### ValidatePercentile(p As Double) (private)
- Назначение: проверка `p` в диапазоне 0..1.
- Вход: `p`. Выход: нет. Побочные эффекты: raise при выходе из диапазона.

### QSortD(ByRef a() As Double, lo As Long, hi As Long) (private)
- Назначение: сортировка массива Double (quick sort, отрезки <32 — вставками, без переполнения стека).
- Вход: `a` — массив; `lo`/`hi` — границы. Выход: отсортированный массив через ByRef. Побочные эффекты: нет.

### PctlOfSorted(ByRef a() As Double, n As Long, p As Double) : Double (private)
- Назначение: перцентиль по уже отсортированному ряду (rank = (n-1)*p, интерполяция между соседями).
- Вход: `a` — отсортированный массив; `n` — длина; `p`. Выход: перцентиль. Побочные эффекты: нет.

### DistinctValues(columnName As String, Optional filters As Variant) : Object
- Назначение: уникальные значения столбца с количеством строк (обёртка над GroupCount по одному столбцу).
- Вход: `columnName`, `filters`. Выход: Dictionary `значение → число`. Побочные эффекты: нет.

### SortDictionaryKeysByValue(d As Object, descending As Boolean) : Variant
- Назначение: отсортировать ключи словаря по числовому значению.
- Вход: `d` — Dictionary (ключ → число); `descending` — направление.
- Выход: массив ключей; пустой `Array()` для пустого/Nothing словаря.
- Побочные эффекты: нет.

### SortKeys(d As Object, numeric As Boolean) : Variant
- Назначение: отсортировать ключи словаря по самому ключу (как числа или как текст) — для осей матриц.
- Вход: `d` — Dictionary; `numeric` — режим сравнения.
- Выход: массив ключей; пустой `Array()` для пустого словаря.
- Побочные эффекты: нет.

### SafeDbl(v As Variant) : Double (private)
- Назначение: числовое значение ключа — берётся первая часть составного ключа до `|` (для осей недель `202643|ДГМ|`).
- Вход: `v` — ключ. Выход: число. Побочные эффекты: нет.
