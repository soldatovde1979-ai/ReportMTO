# modPivotBuilder

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
CORE-модуль, generic. Конструирует PivotTable по переданным именам полей; не знает про Блоки/направления — раскладку задаёт Content Spec (`modContentMTO.BuildPivots`). На 24.08.2026 направлением МТО НЕ используется (Блоки переведены на `modAggregate` + generic-рендер матрицы); остаётся как готовый механизм для будущих направлений.

## Процедуры / Функции

### CreatePivotCache(sourceRange As Range) : PivotCache
- Назначение: создать кэш сводной таблицы по диапазону-источнику.
- Вход: `sourceRange` — диапазон исходных данных.
- Выход: объект `PivotCache`.
- Побочные эффекты: нет.

### BuildPivotTable(cache As PivotCache, destCell As Range, ptName As String, rowFields As Variant, colFields As Variant, filterFields As Variant, dataFieldSpecs As Collection) : PivotTable
- Назначение: построить сводную таблицу с заданными полями строк/столбцов/фильтра и полями данных.
- Вход: `cache` — кэш сводной; `destCell` — ячейка размещения; `ptName` — имя таблицы (существующая с тем же именем предварительно очищается); `rowFields`/`colFields`/`filterFields` — Variant-массивы имён полей (можно `Array()`); `dataFieldSpecs` — Collection элементов `Array(имяПоля, xlFunction, caption)`.
- Выход: созданный объект `PivotTable`.
- Побочные эффекты: создаёт/пересоздаёт PivotTable на листе (мутация книги).
