# Манифест v6.1

sha256, первые 12 символов. Проверка: `sha256sum <файл>`.

| Файл | sha256 | |
|---|---|---|
| `src/powerquery/Query-ImportJSON.pq` | `5dd7c1242542` |  |
| `src/powerquery/fnComputeGroupMetrics.pq` | `7135dec63565` |  |
| `src/powerquery/fnComputeKey.pq` | `ede67cc03e45` | **изменён в v6.1** |
| `src/powerquery/fnNormalizeFields.pq` | `f06bfdbbd783` |  |
| `src/powerquery/fnUpsert.pq` | `4b7f3a0c82fa` |  |
| `src/powerquery/qDiagImport.pq` | `abf7318e24cd` |  |
| `src/powerquery/qExistingData.pq` | `2a0fb656ea98` | **изменён в v6.1** |
| `src/vba/modAIGateway.bas` | `b202cad57d75` |  |
| `src/vba/modAggregate.bas` | `a92b2938235c` |  |
| `src/vba/modColor.bas` | `0eae9fcbb93d` |  |
| `src/vba/modContentMTO.bas` | `01b68ab06ab3` | **изменён в v6.1** |
| `src/vba/modHTMLEngine.bas` | `dba6d3bcbdd3` |  |
| `src/vba/modLog.bas` | `003e1b6f7bd2` |  |
| `src/vba/modMain.bas` | `b4f84e859e6c` |  |
| `src/vba/modPQSync.bas` | `64aa929411df` |  |
| `src/vba/modPivotBuilder.bas` | `ddf194213903` |  |
| `tmp_index.html` | `4d14e89fd502` |  |
| `tools_Build-ReportMTO.ps1` | `790f99877d9b` | **изменён в v6.1** |
| `tests/test_sppr_tablet_v1.json` | `78bd8c74900a` |  |
| `tests/test_upsert_same_keys_v1.json` | `99da18d45643` |  |
| `tests/ОЖИДАЕМОЕ_v3.md` | `1d3d7c713381` |  |

`.bas` и `.pq` — UTF-8 без BOM, LF. `tools_Build-ReportMTO.ps1` — UTF-8 с BOM.
В `.bas` нет символов вне Windows-1251 (проверено): сборочный скрипт конвертирует их
в кодировку системы перед импортом, и всё вне 1251 стало бы `?`.
