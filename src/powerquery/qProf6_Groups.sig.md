# qProf6_Groups

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
Профильный срез шага WithGroups конвейера Query-ImportJSON (после fnComputeGroupMetrics). В сборку не входит.

## Процедуры / Функции

### qProf6_Groups : table
- Вход: prmSourcePath (параметр, файл или папка)
- Выход: таблица {"Rows"} с числом строк WithGroups
- Побочные эффекты: чтение файлов JSON; в книгу не пишет
