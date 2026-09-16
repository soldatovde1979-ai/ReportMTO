# План правок по задачам A–C (актуализация analysis_loadpackage_hang_v1.0)

> Версия: v1.0 от 15.09.2026. Основание: [`analysis_loadpackage_hang_v1.0.md`](analysis_loadpackage_hang_v1.0.md:1).
> Назначение: зафиксировать фактическое состояние задач A–C после ревизии кода
> и перечислить **оставшиеся** правки построчно. Анализ v1.0 описывает код до правок
> v8.2/v8.3; этот план сверяет его с текущим состоянием исходников.

## 1. Фактическое состояние задач A–C (что уже сделано)

| Задача | Статус | Где в коде |
|---|---|---|
| B. Прогноз/индикация во время Refresh | **Выполнено** | [`modMain.bas`](src/vba/modMain.bas:150) — прогноз по `DATA/LAST_LOAD_SECONDS`/`DATA/LAST_LOAD_ROWS`, `Application.StatusBar`, запись замеров после Refresh ([`modMain.bas`](src/vba/modMain.bas:183)) |
| C.1. Ретеншн без `try` в режиме папки | **Выполнено** | ветка `IsFolder` в `ToDT`/`RowAnchor`/`ColDates` — [`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:147) |
| C.2. `fnComputeGroupMetrics` одним проходом | **Выполнено** | v8.3: колонки-носители `__start_dt`/`__end_dt` + `List.Min`/`List.Max` — [`fnComputeGroupMetrics.pq`](src/powerquery/fnComputeGroupMetrics.pq:16) |
| C.3. `fnComputeKey` без построчного `try` | **Выполнено** | v8.3: `HasThreeArgs` проверяется один раз — [`fnComputeKey.pq`](src/powerquery/fnComputeKey.pq:20) |
| C.5. Режим папки минует `fnUpsert` | **Выполнено** | `Upserted = if IsFolder then NewData` — [`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:128); `qExistingData` в цепочку `Retained` при `IsFolder=true` не входит, значит не вычисляется (ленивость M) |
| C.4. Ревизия буферов | **Частично** | буферы убраны из `fnComputeGroupMetrics`; остался `Table.Buffer` в [`fnDedupByKey.pq`](src/powerquery/fnDedupByKey.pq:40) — см. раздел 3 |
| A. Профилирование конвейера | **Частично** | `qProf1–qProf7` созданы, прогонщик [`tools/_tmp_profile_pq.ps1`](tools/_tmp_profile_pq.ps1:1) есть; прогон и фиксация замеров не выполнены — см. раздел 2 |

## 2. Остаток задачи A: прогон профилей и фиксация замеров

Кода писать не нужно — инструменты уже есть. Действия:

1. **Перегенерировать `qProf1–qProf7`** скриптом [`tools/_tmp_generate_qprof.ps1`](tools/_tmp_generate_qprof.ps1:1)
   (сейчас qProf-файлы — копии текущего [`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:1);
   после любой правки конвейера генерацию повторять).
2. **Прогнать [`tools/_tmp_profile_pq.ps1`](tools/_tmp_profile_pq.ps1:1)** — два замера:
   `_prof1` (1 файл, 4976 записей) и `_prof2` (2 файла, 22 528 записей). Каждый прогон
   выдаёт строки `STAGE qProfN: <rows> rows, <sec> s`; разница соседних стадий — стоимость
   функции (таблица соответствия стадий — в анализе §4, задача A).
3. **Зафиксировать замеры** в таблицу ниже (дополнить этот документ после прогона).
4. **Прогон на полном объёме** (5 файлов, 217 768 записей) — отдельно, только если
   экстраполяция с 22 528 записей не выявит виновника однозначно.

Замечания к прогону:

- [`tools/_tmp_profile_pq.ps1`](tools/_tmp_profile_pq.ps1:4) жёстко открывает
  `build\ReportMTO v7.0.xlsm` и **сохраняет её** после добавления qProf-запросов и
  листов-таблиц. Прогон выполнять на **копии книги**, чтобы не загрязнять релизную сборку.
- В сборочный скрипт qProf не входят (`$pqOrder` в [`tools/build-report-mto.ps1`](tools/build-report-mto.ps1:268)
  их не содержит) — в боевую книгу они не попадают, это правильно; менять сборку не нужно.

### Таблица замеров (заполняется после прогона)

| Стадия | 1 файл, 4976 зап. | 2 файла, 22 528 зап. | Дельта (стоимость шага) |
|---|---|---|---|
| qProf1_Source | ? c | ? c | чтение + парсинг JSON |
| qProf2_Expanded | ? c | ? c | разворачивание записей |
| qProf3_Normalized | ? c | ? c | fnNormalizeFields |
| qProf4_Key | ? c | ? c | fnComputeKey |
| qProf5_Dedup | ? c | ? c | fnDedupByKey |
| qProf6_Groups | ? c | ? c | fnComputeGroupMetrics |
| qProf7_Retention | ? c | ? c | ретеншн + финальный `Table.Buffer` |

## 3. Остаток задачи C.4: правка `fnDedupByKey` (единственная правка кода)

### 3.1. Что меняю и почему

Файл [`src/powerquery/fnDedupByKey.pq`](src/powerquery/fnDedupByKey.pq:1), строки 38–41:

```m
Sorted =
    if HasTime
    then Table.Buffer(Table.Sort(tbl, {{"status_date", Order.Descending}}))
    else Table.Buffer(tbl),
```

Правка:

```m
Sorted =
    if HasTime
    then Table.Sort(tbl, {{"status_date", Order.Descending}})
    else tbl,
```

**Почему:**

- `fnDedupByKey` вызывается безусловно в обоих режимах ([`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:118)),
  и в режиме папки `Table.Buffer` держит второй полный набор 217 тыс. строк в памяти —
  источник стоимости №2.1 из анализа.
- `Table.Sort` сам материализует результат (сортировка не ленива), а `Table.Distinct`
  сохраняет порядок первого вхождения — детерминизм «победителя» обеспечивается
  сортировкой, а не буфером.
- `Table.Buffer` в ветке без `HasTime` тоже убирается: при отсутствии `status_date`
  сортировки нет и детерминизм `Distinct` не гарантирован никем — но эта ветка является
  защитной (поле не попало в выборку) и на реальных данных не возникает.

### 3.2. Сопутствующая правка комментариев

Обновить комментарии [`fnDedupByKey.pq`](src/powerquery/fnDedupByKey.pq:15) и
[`fnDedupByKey.pq`](src/powerquery/fnDedupByKey.pq:37): заменить утверждение
«Table.Buffer обязателен» на «детерминизм обеспечивается Table.Sort + порядком первого
вхождения в Table.Distinct, фиксируется тестом». Версию в шапке поднять до v8.4
(пункт B.2.4 в стиле существующих записей).

### 3.3. Критерий приёмки (перед коммитом)

1. Тест ЗН-006: две записи «Готов к приемке» (06:00 и 10:00) с одним Key →
   выживает 10:00, `deltaHours` считается от неё (тест уже описан в шапке файла).
2. Замеры qProf5_Dedup до/после: стоимость шага должна уменьшиться либо остаться
   в пределах погрешности (если вырастет — правку откатить).

### 3.4. Что сознательно НЕ трогаю

| Место | Причина оставить |
|---|---|
| `List.Buffer` в `Source` — [`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:93) | нужен по ТЗ п.3.1: убирает повторное чтение папки |
| `UpsertedBuf = Table.Buffer(Upserted)` — [`Query-ImportJSON.pq`](src/powerquery/Query-ImportJSON.pq:136) | финальный буфер перед выгрузкой; именно он разрешён п. C.4 |
| `NewBuf = Table.Buffer(newData)` — [`fnUpsert.pq`](src/powerquery/fnUpsert.pq:16) | режим одного файла, запрос самоссылающийся; вне пути пакетной загрузки |
| `Buffered = Table.Buffer(Cleaned)` — [`qExistingData.pq`](src/powerquery/qExistingData.pq:33) | в режиме папки не вычисляется вовсе (см. C.5) |

## 4. Порядок выполнения

1. **Без правок кода**: перегенерировать qProf, прогнать профили на копии книги,
   заполнить таблицу раздела 2.
2. По результатам qProf5_Dedup принять решение о правке раздела 3
   (порог для отказа — рост стоимости шага или выход дельты за погрешность замеров).
3. Применить правку 3.1–3.2, прогнать критерии приёмки 3.3.
4. Обновить [`analysis_loadpackage_hang_v1.0.md`](analysis_loadpackage_hang_v1.0.md:1)
   итогами замеров (раздел 2 исходного анализа ссылается на код до v8.2/v8.3).
5. После полного прогона (если потребуется) — зафиксировать «Записано N из M» и
   общее время как новую базовую точку для прогноза задачи B.
