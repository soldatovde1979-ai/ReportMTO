# modPivotBuilder

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Программное построение сводных таблиц Excel (PivotCache + PivotTable) на листах книги.

## Процедуры / Функции

### CreatePivotCache(sourceRange As Range) : PivotCache
- Вход: диапазон источника
- Выход: PivotCache книги
- Побочные эффекты: создаёт кэш сводной в книге

### BuildPivotTable(cache As PivotCache, destCell As Range, ptName As String, rowFields As Variant, colFields As Variant, filterFields As Variant, dataFieldSpecs As Collection) : PivotTable
- Вход: кэш, ячейка назначения, имя сводной, массивы полей строк/колонок/фильтров, коллекция спецификаций полей данных
- Выход: построенная PivotTable
- Побочные эффекты: создаёт сводную таблицу на листе; перезаписывает сводную с тем же именем при повторе
