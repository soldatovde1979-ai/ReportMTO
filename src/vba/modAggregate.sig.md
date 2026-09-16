# modAggregate

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Снимок таблицы tbDATA в память (один проход) и агрегации по нему: счётчики, средние, перцентили, словари значений — источник цифр для всех контентных модулей.

## Процедуры / Функции

### BeginSnapshot(lo As ListObject)
- Вход: ListObject tbDATA
- Выход: нет (создаёт снимок в памяти)
- Побочные эффекты: заполняет module-level кэш (mData, mCols, mReady)

### EndSnapshot()
- Вход: нет
- Выход: нет
- Побочные эффекты: очищает кэш снимка

### IsReady() : Boolean
- Вход: нет
- Выход: создан ли снимок
- Побочные эффекты: нет побочных эффектов

### RowCount() : Long
- Вход: нет
- Выход: число строк снимка
- Побочные эффекты: нет побочных эффектов

### ColIndex(columnName As String) : Long
- Вход: имя колонки
- Выход: индекс колонки; ошибка vbObjectError+3, если снимка нет
- Побочные эффекты: нет побочных эффектов

### HasColumn(columnName As String) : Boolean
- Вход: имя колонки
- Выход: есть ли колонка в снимке
- Побочные эффекты: нет побочных эффектов

### CellText(r As Long, columnName As String) : String
- Вход: номер строки, имя колонки
- Выход: нормализованный текст ячейки (NormText)
- Побочные эффекты: нет побочных эффектов

### CellRaw(r As Long, columnName As String) : Variant
- Вход: номер строки, имя колонки
- Выход: сырое значение ячейки
- Побочные эффекты: нет побочных эффектов

### GroupCount(groupCols As Variant, Optional filters As Variant) : Object
- Вход: колонки группировки, необязательные фильтры
- Выход: Dictionary «составной ключ| -> число строк»
- Побочные эффекты: нет побочных эффектов

### GroupCountDistinct(groupCols As Variant, distinctCol As String, Optional filters As Variant) : Object
- Вход: колонки группировки, колонка уникальных значений, фильтры
- Выход: Dictionary с числом уникальных значений distinctCol по группам
- Побочные эффекты: нет побочных эффектов

### GroupAverage(groupCols As Variant, valueCol As String, Optional filters As Variant) : Object
- Вход: колонки группировки, колонка значений, фильтры
- Выход: Dictionary со средним по группам (пустые группы не включаются)
- Побочные эффекты: нет побочных эффектов

### GroupPercentile(groupCols As Variant, valueCol As String, p As Double, Optional filters As Variant) : Object
- Вход: колонки группировки, колонка значений, перцентиль 0..1, фильтры
- Выход: Dictionary с перцентилем по группам
- Побочные эффекты: нет побочных эффектов

### Percentile(valueCol As String, p As Double, ByRef hasValue As Boolean) : Double
- Вход: колонка значений, перцентиль 0..1
- Выход: перцентиль по всей колонке; hasValue=False, если значений нет (возврат 0)
- Побочные эффекты: нет побочных эффектов

### DistinctValues(columnName As String, Optional filters As Variant) : Object
- Вход: колонка, фильтры
- Выход: Dictionary «значение -> количество»
- Побочные эффекты: нет побочных эффектов

### SortDictionaryKeysByValue(d As Object, descending As Boolean) : Variant
- Вход: словарь, направление сортировки
- Выход: массив ключей, отсортированный по значениям; пустой словарь -> пустой Array()
- Побочные эффекты: нет побочных эффектов

### SortKeys(d As Object, numeric As Boolean) : Variant
- Вход: словарь, признак числовой сортировки
- Выход: массив ключей, отсортированный по самим ключам (числово или текстово)
- Побочные эффекты: нет побочных эффектов

### Приватные хелперы
NormText, NormBoolLiteral, RowMatchesFilters, GroupKey, ColIndexes, RequireSnapshot, ValidatePercentile, QSortD, PctlOfSorted, SafeDbl — нормализация значений, фильтры строк, составные ключи группировки, внутренняя сортировка и расчёт перцентиля. Внешних источников не трогают.
