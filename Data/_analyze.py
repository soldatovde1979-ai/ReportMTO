import json
import collections

d = json.load(open(r'Data/sppr_tablet_20260824_161449_25000rec.json', encoding='utf-8'))
out = ['total records: %d' % len(d)]
uniq = set((r['number'], r['date'], r['defect_desc'], r['ready_for'], r['direction']) for r in d)
out.append('unique(num,date,defect,ready_for,direction): %d' % len(uniq))
uniq2 = set((r['number'], r['date'], r['defect_desc']) for r in d)
out.append('unique(num,date,defect): %d' % len(uniq2))
out.append('unique number: %d' % len({r['number'] for r in d}))
combo = collections.Counter((r['ready_for'], r['direction']) for r in d)
out.append('combos: %s' % dict(combo))
out.append('records with employee: %d' % sum(1 for r in d if r['employee']))
out.append('in_bounds=True: %d' % sum(1 for r in d if r['in_bounds']))
open(r'Data/_check.txt', 'w', encoding='utf-8').write('\n'.join(out))
print('done')
