# modContentDisc

Статус: CLEAN
Обновлено: 16.09.2026

## Назначение
HTML-блоки слайдов 2–4 (дисциплина подписания): недели, посты, люди, статистика подписания, неподписанные наряды. Единица счёта — событие подписания (или заказ-наряд, где указано).

## Процедуры / Функции

### ResetDisc()
- Вход: нет
- Выход: нет
- Побочные эффекты: сбрасывает кэш mReady

### BuildWeeksTable(dir As String) : String
- Вход: дирекция
- Выход: HTML-таблица «% планшет» по неделям
- Побочные эффекты: читает снимок modAggregate; пишет только в кэш

### BuildBlock1(sDir As String) : String
- Вход: дирекция
- Выход: HTML блока 1 (сводка подписания + подсветка «% планшет» шкалой)
- Побочные эффекты: читает снимок; использует modColor.PercentToColor

### BuildPostsTable(dir As String) : String
- Вход: дирекция
- Выход: HTML-таблица подписания по постам
- Побочные эффекты: читает снимок modAggregate

### BuildPeople(dir As String) : String
- Вход: дирекция
- Выход: HTML-таблица подписания по людям (без ФИО наружу)
- Побочные эффекты: читает снимок modAggregate

### BuildSignStat(dir As String) : String
- Вход: дирекция
- Выход: HTML статистики подписания (по заказ-нарядам)
- Побочные эффекты: читает снимок modAggregate

### BuildNoSignSplit(sDir As String, fld As Long, emptyLab As String, colLab As String) : String
- Вход: дирекция, индекс поля, подписи пустой/заполненной групп
- Выход: HTML-таблица «подписано/не подписано» по полю
- Побочные эффекты: читает снимок modAggregate

### BuildKpiUnsigned() : String
- Вход: нет
- Выход: HTML KPI неподписанных
- Побочные эффекты: читает снимок modAggregate

### BuildUnsignedAge() : String
- Вход: нет
- Выход: HTML корзин возраста неподписанных
- Побочные эффекты: читает снимок modAggregate

### BuildUnsignedPost() / BuildUnsignedOwner() / BuildUnsignedZnType() : String
- Вход: нет
- Выход: HTML разрезы неподписанных по посту / подразделению / виду ремонта
- Побочные эффекты: читают снимок modAggregate

### FillDiscPlaceholders(d As Object)
- Вход: словарь плейсхолдеров отчёта
- Выход: нет
- Побочные эффекты: записывает BLOCK_*-плейсхолдеры слайдов 2–4 в словарь d

### Приватные хелперы
EnsureDisc, ZoneOf, ArmField, TimeField, PctHeat, PctArrowTd, AvgSpan, SsRow, AgeBucket4, SignedCount, UnsignedBy — кэш событий, сопоставление полей дирекций, оформление строк/стрелок. Внешних источников не трогают.
