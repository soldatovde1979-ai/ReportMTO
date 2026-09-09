# Задачи

> Версия 1.4 от 08.09.2026. Закрыты остатки tz_Reports2.md: плейсхолдеры BLOCK_6_*_DEPT,
> пояснение при пустом REPORT/SLIDE_ZONES, «Создали ЗН» дашборда без FBase (см. ниже).
> Версия 1.3 от 08.09.2026. Закрыта задача расширенного логирования (DEBUG 0/1/2 + внешний
> журнал ReportMTO_log.txt); по пути найден и исправлен M-баг, ломавший загрузку данных (см. ниже).
> Версия 1.2 от 08.09.2026. Закрыты задачи по ошибкам ИИ-слоя и фиксации чек-ина (см. ниже).
> Версия 1.1 от 08.09.2026. Добавлена задача по ошибкам ИИ-слоя из логов книги.
> Версия 1.0 от 08.09.2026. Создан по поручению пользователя.

## Срочно

- [x] 09.09.2026: ТЗ v1.2 — переход отчёта на 4 слайда, код внесён. T1 fnNormalizeFields.pq v7
      (dateWeek/dateMonth/isSigned); T2 modAggregate.bas v3.2 (GroupPercentile/Percentile,
      собственный QuickSort); T3-T7, T9 modContentMTO.bas v7.3 (FBase пуст — in_bounds не
      фильтруется; псевдонимы «Сотрудник N»; рамка td.pct вместо заливки; SVG-примитивы;
      блоки слайдов 1-4; промпт/разбор под 4 слайда; BuildPlaceholders и DebugCheckPlaceholders);
      T8 tmp_index.html v3.0 и build\tmp_index.html (4 слайда, переключатель темы, data-drill
      снят); T10 REPORT/RETENTION_WEEKS (пусто, зарезервирован) в обе build-книги
      (tools/add_retention_key.ps1). Переимпорт в build-книгу (tools/apply_vba_tmp.ps1),
      compile_check: COMPILE_OK. НЕ СДЕЛАНО: woff2 Oswald/IBM Plex не встроены — файлов в
      репозитории нет, интернет на машине недоступен (TODO-блок в шаблоне); прогон
      DebugGenerateOffline и сверка чисел на данных не выполнялись (рабочая книга вне
      репозитория); M-код v7 в build-книгу не загружен (нужно «Загрузить» данные).
- [x] 09.09.2026: автономный прогон рабочей книги и починка зависания. Прогоны вставали в
      BuildPlaceholders на 136108 строках (>30 мин, прерывались): причина — O(N^2) конкатенация
      дампа расшифровки в BuildDataDump (43,5 МБ). Исправлено: двухпроходный массив + Join и
      поэтапные тайминги слайдов (modContentMTO.bas). Полный прогон с ИИ: result\Report_20260909_134836.html
      за ~6,5 мин, все 7 выводов ИИ распознаны. Остаточные медленные этапы: Слайд 1 ~100 c,
      Слайды 2/3 (Блок 6 ×4) ~270 c.
- [x] install.ps1 v1.8 / install_prod.ps1 v1.4 (09.09.2026): подробный по-модульный вывод
      (VBA_ADDED/DIFF/SAME, PQ_ADDED/DIFF/SAME/SKIP, TPL_DIFF/SAME) и обязательная пост-установочная
      верификация по исходникам (VBA/PQ/TPL_VERIFY_OK|FAIL, VERIFY_OK|FAIL; провал -> exit 1).
      Лог человекочитаемый (русские пояснения + позиция первого отличия и коды символов, UTF-8 консоль).
      Диагностикой найдены и убраны ложные различия: Attribute-строка VBA (не возвращается
      CodeModule.Lines) и конечный перевод строки PQ (обрезается Excel при Save). Контрольный прогон
      install_prod 09.09.2026: все SAME + все VERIFY_OK, exit 0.
- [x] 09.09.2026: install_prod падал на VERIFY_FAIL VBA:modContentMTO. Причина: VBE при импорте
      .bas сам добавляет `:` после `Else` с оператором на той же строке (строка `Else vals(6)=...`
      → в книге `Else: vals(6)=...`), поэтому книга всегда отличалась от исходника. Исправлен
      исходник modContentMTO.bas (`Else:`), повторный install_prod: все SAME + VERIFY_OK, exit 0;
      compile_check: COMPILE_OK. Отдельно: в modContentMTO.bas 10 юникод-символов вне cp1251
      (⚠▲▼→∈×−) превращаются в «?» при импорте через ANSI — верификация этого не ловит (roundtrip).
- [x] 09.09.2026: «Method or data member not found» на modAggregate.Percentile в корневой книге.
      Причина: install.ps1 ставил только modMain/modContentMTO, а modAggregate v3.2 (Percentile,
      GroupPercentile) в книгу не попадал — новый modContentMTO вызывал отсутствующие функции.
      Сделано: modAggregate добавлен в install.ps1 (установка + верификация), корневая книга и
      build обновлены (VERIFY_OK), compile_check корневой: COMPILE_OK. В e2e добавлен STEP0_COMPILE
      (полная компиляция проекта до импорта тестового модуля).
- [x] 09.09.2026: e2e доведён до зелёного. Починено: «Subscript out of range» в
      BuildKpiOverview (окно из одной недели: opened(2)/closed(2)/hang(2)/med(2) без
      защиты hasPrev — ретеншн отсекает вторую неделю тестового JSON) — предвычисление
      cur/prev; эталоны modSelfTest.bas v1.2 переписаны под 4 слайда и 22 строки,
      e2e-скрипт: STEP0_COMPILE, ожидания 22, маркеры STEP4 (mto-report/slide-nav вместо
      удалённого drill-dump), печать всех CHECK в терминал. Прогон: ALL PASSED
      (35 CHECK=1 + upsert). Прод обновлён (install_prod: modContentMTO/modAggregate
      переустановлены, VERIFY_OK). Отчёт на проде: result\Report_20260909_233710.html
      (1,37 МБ, 36 c) — ИИ вернул 0 символов, выводы заглушки (ключи на месте:
      AI_API_KEY/ deepseek.key; причина — недоступность провайдера, не код).
- [x] tests/expected.md дополнен актуальными эталонами e2e v7 (22 строки, 4 слайда,
      ключевые цифры). Полная переработка остального документа (эталоны v6.1) — не делалась.
- [x] tz_Reports2.md: проверено, что уже реализовано, и закрыты остатки — `BuildBlock6Weekly` с
      параметром `byDept` и новые плейсхолдеры `BLOCK_6_DENT_DEPT`/`BLOCK_6_DGM_DEPT` (понедельная
      раскладка по подразделениям) на слайдах 2/3; пустой `REPORT/SLIDE_ZONES` -> пояснение в
      `BLOCK_1_*_ZONES`; «Создали ЗН» дашборда считается без FBase (по всей истории, включая
      «НЕ ПОДПИСАНО»). Правки: modContentMTO v7.2, tmp_index.html v2.1 (корень и build\). Компиляция
      VBA в Excel и e2e не прогонялись.
- [x] Обновить [`docs/rules.md`](rules.md) — дополнен 08.09.2026 (правила cmd/PowerShell, search_files/JSON, выгрузка 2026)
- [x] ФАЗА 1 (09.09.2026): Compile error «ByRef argument type mismatch» в ValidateRequiredColumns
      устранён объявлением `Dim required() As String` + `Split(...)` (modContentMTO.bas); компиляция
      всего VBAProject пройдена автоматически через COM Run — COMPILE_OK (tools/compile_check.ps1).
      Вопрос владельцу: менять ли Core HasColumn на `ByVal columnName As String` — ждёт решения.
- [x] ФАЗА 2 (09.09.2026): ретеншн tbDATA (инструкция §4) — qKeepWeeks.pq (лист Variable, ключ
      DATA/KEEP_WEEKS), финальный шаг R-1 в Query-ImportJSON v7.3 (опорная дата status_date, иначе
      date; отсечка от самой свежей недели данных; 0 = выключено; строка без даты сохраняется).
      Ключ DATA/KEEP_WEEKS=52 добавлен в tblVariable build-книг (tools/add_keep_weeks_key.ps1).
      install.ps1 на build: все VERIFY_OK. Проверка ретеншна на тестовом JSON
      (tools/check_retention.ps1): KW0=23 / KW52=22 / KW1=22, повторные прогоны идемпотентны —
      отсечена единственная строка вне окна (2025-08-25). Перенос в рабочую книгу — не сделан
      (книга вне репозитория, нужен ручной «Загрузить» и контроль KEEP_WEEKS=1).
- [x] ФАЗА 3.2 (09.09.2026): документация приведена к инварианту «е» — «Готов к приемке» в
      docs/spec.md (enum, deltaHours, образец M-кода) и docs/specs/content-spec.md с пояснением
      «1С пишет через «е», сравнение нормализует (NormStatus)»; docs/archive не тронут.
- [ ] Ждут решения владельца: (1) HasColumn ByVal в Core; (2) ФАЗА 3.1 развилка norm_status —
      столбец в fnNormalizeFields (а) или нормализация в modAggregate (б); (3) ФАЗА 4.1 парсер
      операторов фильтра (латентный дефект «по левому вхождению»); (4) допущение ретеншна —
      отсечка от свежей недели данных и KEEP_WEEKS=52 как стартовое значение.
- [ ] НЕ СДЕЛАНО: ФАЗА 4.2 синхронизация claude_spec/claude_data/claude_next-steps.md —
      файлы приложены к проекту Claude вне репозитория; в workspace их нет.
- [ ] Разобрать ошибки ИИ-слоя из логов книги: «Слой ИИ пропущен: Столбец 'year_status' не найден в tbDATA», «Ошибка -2147221502: Столбец 'year_status' не найден в tbDATA», «Ответ ИИ не распознан или недоступен - используются заглушки». Переоткрыто 08.09.2026: первопричина — рабочая книга (вне репозитория, 136108 строк) содержит M-код v5, где `*_status` не вычисляются из `status_date`, а выгрузка 2026 их не отдаёт -> столбцы `year_status`/`week_status` исчезли из tbDATA, VBA падал на ColIndex. Сделано: VBA больше не опирается на `year_status` (год из `yearWeek`), добавлен `ValidateRequiredColumns`; M-код v6 доставлен в корневую ReportMTO.xlsm. Осталось: доставить v6 в рабочую книгу из лога и перезагрузить данные.
- [x] Зафиксировать, что при чек-ине и arm = НЕ ПОДПИСАНО планшет приложен к метке, но документ не подписан — arm корректен (записано в [`docs/spec.md`](spec.md) и [`docs/data.md`](data.md) 08.09.2026)
- [x] Расширенное логирование (08.09.2026): ключ `DEBUG` на листе Variable (0 — штатно, 1 — этапы и
      тайминги, 2 — полный промпт до 20000 и первые 4000 ответа, без API-ключа); все записи дублируются
      в `ReportMTO_log.txt` рядом с книгой (UTF-8, append); при «Столбец не найден» лог показывает
      фактический список столбцов tbDATA. Ключ DEBUG=0 добавлен в build-книгу и starter
      ([`tools/add_debug_key.ps1`](../tools/add_debug_key.ps1)). По пути найден и исправлен M-баг:
      в fnNormalizeFields апостроф `'` вместо `//` перед комментарием — Refresh падал
      «[Expression.Error] Ожидался токен Identifier», загрузка данных была сломана целиком.
      Проверено: smoke DEBUG=2 (файл, UTF-8, append) и e2e STEP4 (сборка отчёта) — PASS.

## Бэклог

- [ ] Сделать свёртку БД
- [ ] Обновить эталоны e2e/self-test под текущую логику (прогон 09.09.2026): e2e ждёт 25 строк и
      CHECK-константы 7 слайдов/BLOCK6/EMP1 (tests/modSelfTest.bas); факт — 22 строки (дедупликация
      ключей + ретеншн KEEP_WEEKS=52: KW0=23/KW52=22 по tools/check_retention.ps1) и промпт 4 слайда
      (ТЗ v1.2); STEP4 не создал свежий debug_*.html: DebugGenerateOffline падает с «Subscript out
      of range» (лог %TEMP%\ReportMTO_log.txt, запись 09.09.2026 22:36). Обновить
      tests/expected.md и CHECK-константы tests/modSelfTest.bas, локализовать строку ошибки
      (DEBUG=1 / VBE break), затем перепрогнать e2e.
