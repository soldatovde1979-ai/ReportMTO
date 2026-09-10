# -*- coding: utf-8 -*-
"""
Структурный линтер VBA — v1.1 от 10.09.2026.

Версия 1.1: проверка зарезервированных слов распространена на модульные
  Private/Public и на параметры процедур. Раньше ловились только Dim/Const,
  и `Private mE As Object` (то есть `Me`) прошёл в модуль незамеченным.
Версия 1.0: первый выпуск.

Ловит то, что в этом проекте уже ломало сборку и что видно без Excel:
  * незакрытые блоки Sub/Function/If/For/Do/With/Select;
  * дубли имён процедур;
  * вызовы процедур модуля с неверным числом аргументов;
  * необъявленные переменные при Option Explicit;
  * незакрытые строковые литералы и знак переноса внутри строки;
  * символы, не представимые в ANSI-1251 (VBE читает .bas только в 1251);
  * превышение лимитов VBA: длина строки и число переносов;
  * Attribute VB_Name, не совпадающий с именем файла.

Не заменяет компиляцию: типы, свойства объектов и позднее связывание не проверяются.
"""
import re, sys, os, glob

# Зарезервированные слова VBA: как имя переменной дают «Expected: identifier».
RESERVED = set( '''
as byref byval call case const declare dim do each else elseif end enum erase error event
exit false for friend function get global gosub goto if implements in input is let lib like
loop me mod new next not nothing null on option optional or paramarray preserve print private
property public raiseevent redim rem resume return select set static step stop sub then to
true type unload until wend while with withevents xor and eqv imp write open close line
'''.split())

KW_OPEN = {
    'sub': ('end sub',), 'function': ('end function',), 'property': ('end property',),
    'type': ('end type',), 'enum': ('end enum',), 'with': ('end with',),
    'select': ('end select',),
}
DECL_RE = re.compile(r'^\s*(?:(public|private|friend)\s+)?(?:static\s+)?(sub|function|property\s+(?:get|let|set))\s+([A-Za-z_]\w*)\s*(\(.*)?$', re.I)
END_RE  = re.compile(r'^\s*end\s+(sub|function|property)\b', re.I)

class Issue:
    def __init__(s, path, line, code, msg):
        s.path, s.line, s.code, s.msg = path, line, code, msg
    def __str__(s):
        return f'{os.path.basename(s.path)}:{s.line}: [{s.code}] {s.msg}'

def strip_strings(t):
    """Убирает содержимое строковых литералов, возвращает (текст, флаг незакрытой строки)."""
    out = []; i = 0; inside = False
    while i < len(t):
        c = t[i]
        if c == '"':
            if inside and i + 1 < len(t) and t[i+1] == '"':
                i += 2; continue
            inside = not inside; out.append('"'); i += 1; continue
        out.append(' ' if inside else c); i += 1
    return ''.join(out), inside

def strip_comment(t):
    """Отрезает комментарий с учётом строковых литералов и только потом
    сообщает о незакрытой строке — иначе кавычка в комментарии даёт ложную тревогу."""
    inside = False; cut = None; i = 0
    while i < len(t):
        c = t[i]
        if c == '"':
            if inside and i + 1 < len(t) and t[i+1] == '"':
                i += 2; continue
            inside = not inside
        elif c == "'" and not inside:
            cut = i; break
        i += 1
    code = t if cut is None else t[:cut]
    if re.match(r"^\s*rem\b", code, re.I): code = ''
    _, unterm = strip_strings(code)
    return code, unterm

def logical_lines(raw_lines):
    """Склейка переносов `_`. Возвращает [(номер первой строки, текст, число переносов)]."""
    out = []; buf = ''; start = 0; cont = 0
    for i, ln in enumerate(raw_lines, 1):
        body = ln.rstrip('\n').rstrip('\r')
        if not buf: start = i
        code, _ = strip_strings(body)
        if code.rstrip().endswith('_'):
            buf += body.rstrip()[:-1]; cont += 1; continue
        out.append((start, buf + body, cont)); buf = ''; cont = 0
    if buf: out.append((start, buf, cont))
    return out

def lint(path):
    issues = []
    raw = open(path, encoding='utf-8', errors='replace').read()
    lines = raw.splitlines(True)

    # 1. кодируемость в ANSI-1251
    for i, ln in enumerate(lines, 1):
        try: ln.encode('cp1251')
        except UnicodeEncodeError as e:
            bad = ln[e.start:e.end]
            issues.append(Issue(path, i, 'ENC1251', f'символ {bad!r} не представим в ANSI-1251'))

    # 2. Attribute VB_Name
    want = os.path.splitext(os.path.basename(path))[0]
    m = re.match(r'Attribute VB_Name = "([^"]+)"', raw)
    if not m: issues.append(Issue(path, 1, 'VBNAME', 'нет строки Attribute VB_Name'))
    elif m.group(1) != want:
        issues.append(Issue(path, 1, 'VBNAME', f'Attribute VB_Name = "{m.group(1)}", а файл {want}.bas'))

    # 3. Option Explicit
    explicit = re.search(r'^\s*Option\s+Explicit\s*(?:\'.*)?$', raw, re.M | re.I) is not None
    if not explicit: issues.append(Issue(path, 1, 'NOEXPLICIT', 'нет Option Explicit'))

    ll = logical_lines(lines)
    stack = []          # (тип, номер строки)
    procs = {}          # имя -> (номер строки, число аргументов, минимум обязательных)
    state = {'cur': None, 'declared': set(), 'module_vars': set()}
    used = []; calls = []

    for lineno, text, cont in ll:
        if cont > 24:
            issues.append(Issue(path, lineno, 'CONT', f'{cont} переносов строки, лимит VBA — 24'))
        if len(text) > 1000:
            issues.append(Issue(path, lineno, 'LONGLINE', f'логическая строка {len(text)} символов, лимит ~1023'))
        code, unterm = strip_comment(text)
        if unterm:
            issues.append(Issue(path, lineno, 'STRING', 'незакрытая строковая константа'))
        for stmt in split_statements(code):
            analyse_stmt(stmt, lineno, stack, procs, issues, path, state)
        code_for_names = code
        for mm in re.finditer(r'\b([A-Za-z_]\w*)\s*\(', code_for_names):
            calls.append((mm.group(1).lower(), lineno, code_for_names))

    for kind, lineno in stack:
        issues.append(Issue(path, lineno, 'BLOCK', f'блок {kind} не закрыт'))
    return issues, procs, used, state['declared'], state['module_vars'], calls, explicit

def analyse_stmt(code, lineno, stack, procs, issues, path, state):
    low = code.lower().strip()
    if not low: return

    d = DECL_RE.match(code)
    if d and not re.match(r'^\s*(?:public|private)?\s*declare\b', code, re.I):
        name = d.group(3)
        if name.lower() in procs:
            issues.append(Issue(path, lineno, 'DUP',
                f'процедура {name} объявлена повторно (первый раз в строке {procs[name.lower()][0]})'))
        args = d.group(4) or '()'
        inner = args[args.find('(')+1:args.rfind(')')] if '(' in args else ''
        parts = [p for p in split_args(inner) if p.strip()]
        opt = sum(1 for p in parts if re.search(r'\boptional\b', p, re.I))
        procs[name.lower()] = (lineno, len(parts), len(parts) - opt)
        stack.append((d.group(2).split()[0].lower(), lineno))
        state['cur'] = name; state['declared'] = set()
        for p in parts:
            mm = re.search(r'(?:byval|byref|optional|paramarray)?\s*([A-Za-z_]\w*)', p.strip(), re.I)
            if mm: state['declared'].add(mm.group(1).lower())
        return
    e = END_RE.match(code)
    if e:
        pop_expect(stack, e.group(1).lower(), issues, path, lineno)
        state['cur'] = None; return

    if re.match(r'^\s*with\b', low): stack.append(('with', lineno)); return
    if re.match(r'^\s*end\s+with\b', low): pop_expect(stack, 'with', issues, path, lineno); return
    if re.match(r'^\s*select\s+case\b', low): stack.append(('select', lineno)); return
    if re.match(r'^\s*end\s+select\b', low): pop_expect(stack, 'select', issues, path, lineno); return
    if re.match(r'^\s*end\s+if\b', low): pop_expect(stack, 'if', issues, path, lineno); return
    # блочный If: Then в конце оператора и ничего после него
    if re.match(r'^\s*if\b', low) and re.search(r'\bthen\s*$', low):
        stack.append(('if', lineno)); return
    if re.match(r'^\s*else\s*$', low) or re.match(r'^\s*elseif\b.*\bthen\s*$', low):
        return
    if re.match(r'^\s*(?:exit|continue)\s+(?:for|do|sub|function)\b', low): return
    if re.match(r'^\s*for\b', low): stack.append(('for', lineno)); return
    if re.match(r'^\s*next\b', low): pop_expect(stack, 'for', issues, path, lineno); return
    if re.match(r'^\s*do\b', low): stack.append(('do', lineno)); return
    if re.match(r'^\s*loop\b', low): pop_expect(stack, 'do', issues, path, lineno); return

    for mm in re.finditer(r'\b(?:dim|static|const|redim(?:\s+preserve)?)\s+(.+)', code, re.I):
        for part in split_args(mm.group(1)):
            nm = re.match(r'\s*([A-Za-z_]\w*)', part)
            if nm and nm.group(1).lower() in RESERVED:
                issues.append(Issue(path, lineno, 'RESERVED',
                    f'«{nm.group(1)}» — зарезервированное слово VBA, именем переменной быть не может'))
    for mm in re.finditer(r'\b(?:dim|static|const|redim(?:\s+preserve)?)\s+(.+)', code, re.I):
        for part in split_args(mm.group(1)):
            nm = re.match(r'\s*([A-Za-z_]\w*)', part)
            if nm: (state['declared'] if state['cur'] else state['module_vars']).add(nm.group(1).lower())
    for mm in re.finditer(r'^\s*(?:public|private)\s+(?!sub|function|property|const|declare|type|enum)(.+)',
                          code, re.I):
        for part in split_args(mm.group(1)):
            nm = re.match(r'\s*(?:withevents\s+)?([A-Za-z_]\w*)', part, re.I)
            if not nm: continue
            # Модульные Private/Public проверяются на зарезервированные слова так же,
            # как Dim: `Private mE As Object` - это объявление переменной `Me`.
            if nm.group(1).lower() in RESERVED:
                issues.append(Issue(path, lineno, 'RESERVED',
                    f'«{nm.group(1)}» — зарезервированное слово VBA, именем переменной быть не может'))
            state['module_vars'].add(nm.group(1).lower())
    # Параметры процедур - тем же правилом.
    dm = DECL_RE.match(code)
    if dm and dm.group(4):
        for part in split_args(dm.group(4).strip('()')):
            nm = re.match(r'\s*(?:optional\s+|byval\s+|byref\s+|paramarray\s+)*([A-Za-z_]\w*)',
                          part, re.I)
            if nm and nm.group(1).lower() in RESERVED:
                issues.append(Issue(path, lineno, 'RESERVED',
                    f'«{nm.group(1)}» — зарезервированное слово VBA, именем параметра быть не может'))

def split_statements(code):
    """Логическая строка -> отдельные операторы (VBA разделяет их двоеточием).
    Метка вида `Fail:` двоеточием не считается."""
    out = []; cur = ''; depth = 0; ins = False
    for ch in code:
        if ch == '"': ins = not ins
        if not ins:
            if ch in '([': depth += 1
            elif ch in ')]': depth -= 1
            elif ch == ':' and depth == 0:
                out.append(cur); cur = ''; continue
        cur += ch
    out.append(cur)
    res = []
    for st in out:
        if re.match(r'^\s*[A-Za-z_]\w*\s*$', st) and len(res) == 0 and len(out) > 1:
            continue          # метка перед оператором
        res.append(st)
    return [x for x in res if x.strip()]

def split_args(s):
    out = []; depth = 0; cur = ''
    ins = False
    for ch in s:
        if ch == '"': ins = not ins
        if not ins:
            if ch in '([': depth += 1
            elif ch in ')]': depth -= 1
            elif ch == ',' and depth == 0:
                out.append(cur); cur = ''; continue
        cur += ch
    out.append(cur)
    return out

def pop_expect(stack, kind, issues, path, lineno):
    if stack and stack[-1][0] == kind: stack.pop()
    else:
        top = stack[-1][0] if stack else 'ничего'
        issues.append(Issue(path, lineno, 'BLOCK', f'закрытие {kind}, а открыт {top}'))

def main(paths):
    total = []
    for p in paths:
        issues, procs, used, declared, mvars, calls, explicit = lint(p)
        total += issues
        print(f'--- {os.path.basename(p):24s} процедур: {len(procs):3d}  замечаний: {len(issues)}')
        for i in issues[:40]: print('    ', i)
    print(f'\nИТОГО замечаний: {len(total)}')
    return 1 if total else 0

if __name__ == '__main__':
    args = sys.argv[1:]
    if not args:
        args = sorted(glob.glob(os.path.join(os.path.dirname(__file__), '..', '..', 'src', 'vba', '*.bas')))
    sys.exit(main(args))
