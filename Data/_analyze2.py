import json, collections

d = json.load(open(r'C:\Projects\ReportMTO\Data\sppr_tablet_20260824_161449_25000rec.json', encoding='utf-8'))
out = []
out.append('records %d' % len(d))
out.append('keys %s' % list(d[0].keys()))

rv = collections.Counter(x['ready_for'] for x in d); out.append('ready_for %s' % dict(rv))
dr = collections.Counter(x['direction'] for x in d); out.append('direction %s' % dict(dr))
arm = collections.Counter(x['arm'] for x in d); out.append('arm %s' % dict(arm))
emp_dep = collections.Counter(x['emp_dep'] for x in d); out.append('emp_dep %s' % dict(emp_dep))
sd = collections.Counter((x['status_date'] == '') for x in d); out.append('status_date empty %s' % dict(sd))
ib = collections.Counter(x['in_bounds'] for x in d); out.append('in_bounds %s' % dict(ib))

emp_vals = set(x['employee'] for x in d if str(x['employee']).strip())
out.append('employee distinct non-empty %d' % len(emp_vals))

post_vals = collections.Counter(x['post'].strip() if isinstance(x['post'],str) else str(x['post']) for x in d)
out.append('post distinct %d' % len(post_vals))

out.append('number distinct %d' % len(set(x['number'] for x in d)))

mt = collections.Counter(x['model_type'] for x in d)
out.append('model_type %s' % dict(mt))

zt = collections.Counter(x['zn_type'] for x in d)
out.append('zn_type %s' % dict(zt))

open(r'C:\Projects\ReportMTO\Data\_analysis.txt', 'w', encoding='utf-8').write('\n'.join(out))
print('ok')
