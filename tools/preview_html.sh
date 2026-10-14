#!/usr/bin/env bash
# Рендер HTML в PNG без Excel и без браузера на экране.
#
# Зачем: сложную логику можно вычитать глазами по коду, а раскладку - нельзя.
# Дефекты вида «стрелка получила рамку бейджа» или «значение уехало вниз»
# видны только на картинке. Скрипт даёт картинку за одну команду.
#
# Использование:
#   tools/preview_html.sh <файл.html> [выход.png] [ширина] [высота]
#
# Пример:
#   tools/preview_html.sh build/report.html /tmp/report.png 1500 900
set -euo pipefail

SRC="${1:?укажите путь к HTML}"
OUT="${2:-${SRC%.html}.png}"
W="${3:-1500}"
H="${4:-900}"

CHROME=""
for c in \
  /opt/pw-browsers/chromium-1194/chrome-linux/chrome \
  /opt/pw-browsers/chromium/chrome-linux/chrome \
  "$(command -v chromium || true)" \
  "$(command -v google-chrome || true)"
do
  if [[ -n "$c" && -x "$c" ]]; then CHROME="$c"; break; fi
done
if [[ -z "$CHROME" ]]; then
  echo "Chromium не найден. Проверьте PLAYWRIGHT_BROWSERS_PATH." >&2
  exit 1
fi

[[ -f "$SRC" ]] || { echo "Нет файла: $SRC" >&2; exit 1; }

"$CHROME" --headless --no-sandbox --disable-gpu --hide-scrollbars \
  --force-device-scale-factor=2 --window-size="${W},${H}" \
  --screenshot="$OUT" "$SRC" 2>/dev/null

echo "Готово: $OUT"
