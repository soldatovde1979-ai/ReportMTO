# -*- coding: utf-8 -*-
# Конвертация системных шрифтов в woff2 для встраивания в tmp_index.html (base64 @font-face).
# Без subset: в дампе расшифровки могут быть произвольные символы (ФИО), глифы не режем.
# flavor='woff2' (атрибут TTFont, не аргумент save) сохраняет полный набор глифов.
import os
from fontTools.ttLib import TTFont

SRCS = {
    'segoeui': 'C:/Windows/Fonts/segoeui.ttf',   # Segoe UI Regular (400)
    'segoeuib': 'C:/Windows/Fonts/segoeuib.ttf', # Segoe UI Semibold (600)
    'consola': 'C:/Windows/Fonts/consola.ttf',   # Consolas (400)
}

os.makedirs('tools/fonts', exist_ok=True)

for name, src in SRCS.items():
    out = 'tools/fonts/%s.woff2' % name
    font = TTFont(src)
    font.flavor = 'woff2'
    font.save(out)
    print('%s -> %s (%d bytes)' % (name, out, os.path.getsize(out)))
