# Манифест v6

sha256, первые 12 символов. Проверка: `sha256sum <файл>`.

| Файл | sha256 |
|---|---|
| `src/powerquery/Query-ImportJSON.pq` | `5dd7c1242542` |
| `src/powerquery/fnComputeGroupMetrics.pq` | `7135dec63565` |
| `src/powerquery/fnComputeKey.pq` | `33ac8a623ce0` |
| `src/powerquery/fnNormalizeFields.pq` | `f06bfdbbd783` |
| `src/powerquery/fnUpsert.pq` | `4b7f3a0c82fa` |
| `src/powerquery/qDiagImport.pq` | `abf7318e24cd` |
| `src/powerquery/qExistingData.pq` | `1b19b4ded8e1` |
| `src/vba/modAIGateway.bas` | `b202cad57d75` |
| `src/vba/modAggregate.bas` | `a92b2938235c` |
| `src/vba/modColor.bas` | `0eae9fcbb93d` |
| `src/vba/modContentMTO.bas` | `30984c60ed74` |
| `src/vba/modHTMLEngine.bas` | `dba6d3bcbdd3` |
| `src/vba/modLog.bas` | `003e1b6f7bd2` |
| `src/vba/modMain.bas` | `b4f84e859e6c` |
| `src/vba/modPQSync.bas` | `64aa929411df` |
| `src/vba/modPivotBuilder.bas` | `ddf194213903` |
| `tmp_index.html` | `4d14e89fd502` |
| `tools_Build-ReportMTO.ps1` | `ce49ab96a4b8` |
| `tests/test_sppr_tablet_v1.json` | `78bd8c74900a` |
| `tests/test_upsert_same_keys_v1.json` | `99da18d45643` |
| `tests/ОЖИДАЕМОЕ_v2.md` | `7783b61bfe6c` |

Все `.bas` и `.pq` — UTF-8 без BOM, LF. `tools_Build-ReportMTO.ps1` — UTF-8 с BOM.
В `.bas` нет ни одного символа вне Windows-1251 (проверено), поэтому сборка их не портит.
