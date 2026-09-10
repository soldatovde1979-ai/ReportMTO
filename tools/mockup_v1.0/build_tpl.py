# -*- coding: utf-8 -*-
"""Шаблон tmp_index.html: из эталона, подменой содержимого слотов на {{ПЛЕЙСХОЛДЕР}}."""
import re, os
HOME=os.environ['HOME']; SRC=os.path.join(HOME,'mnt','ReportMTO')
src=open(os.path.join(SRC,'temp','MTO_макет_отчета_v4.0.html'),encoding='utf-8').read()

names=[]
def sub(m):
    names.append(m.group(1)); return '{{%s}}' % m.group(1)
tpl=re.sub(r'<!--SLOT:([A-Z_0-9]+)-->.*?<!--/SLOT:\1-->', sub, src, flags=re.S)
assert '<!--SLOT' not in tpl, 'остались неподменённые слоты'

# шапка: заголовок, период, факты, подпись недели
tpl=tpl.replace('<title>Отчёт МТО · 202613 · планшеты и техника</title>',
                '<title>{{REPORT_TITLE}}</title>')
tpl=re.sub(r'<h1>Отчёт МТО</h1>', '<h1>{{REPORT_TITLE}}</h1>', tpl)
tpl=re.sub(r'<p class="lede">.*?</p>', '<p class="lede">{{REPORT_LEDE}}</p>', tpl, flags=re.S)
tpl=re.sub(r'<dl class="facts">.*?</dl>', '<dl class="facts">{{FACTS}}</dl>', tpl, flags=re.S)
tpl=re.sub(r'<span class="ttl">Ремзона · неделя[^<]*</span>',
           '<span class="ttl">Ремзона · {{REPORT_WEEK_LABEL}}</span>', tpl)
tpl=re.sub(r'<span class="ttl">Парк и заезды · неделя[^<]*</span>',
           '<span class="ttl">Парк и заезды · {{REPORT_WEEK_LABEL}}</span>', tpl)
tpl=re.sub(r'(<span class="ttl">Аналитика по постам · (?:ДЭНТ|ДГМ) · )неделя[^<]*(</span>)',
           r'\1{{REPORT_WEEK_LABEL}}\2', tpl)
tpl=re.sub(r'<footer>.*?</footer>',
           '<footer>{{REPORT_FOOTER}}</footer>', tpl, flags=re.S)

out=os.path.join(SRC,'temp','tmp_index_v4.0.html')
open(out,'w',encoding='utf-8').write(tpl)
ph=sorted(set(re.findall(r'\{\{([A-Z_0-9]+)\}\}', tpl)))
print('слотов подменено:', len(names))
print('плейсхолдеров всего:', len(ph))
for x in ph: print('   {{%s}}' % x)
print('размер, КБ:', round(len(tpl.encode())/1024))
print('файл:', out)
