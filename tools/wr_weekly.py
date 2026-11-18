"""코웨이 TV Weekly Report(새 양식, 2026-11~) → TV_WR (doc.tv.wr) — 기존 11월 TV_WR(11/8본) 구조에 맞춤.
사용: python3 -I wrnov.py weekly.xlsx out.json
hist · src · tgt · v 는 사이트에 있는 값을 그대로 쓰므로 여기서는 만들지 않는다."""
import sys, re, json, datetime, warnings
import openpyxl
warnings.filterwarnings('ignore')

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC, data_only=True)

ITEM_MAP = {'카페 얼음': '카페얼음정수기', '페블2': '페블체어2', '얼음정수기 (25년~)': '아이콘얼음정수기',
            '코웨이아이콘3': '아이콘3', '코웨이아이콘프로': '아이콘프로'}
canon = lambda s: ITEM_MAP.get(str(s).strip(), str(s).strip())


def hidden_cols(ws):
    h = set()
    for _, d in ws.column_dimensions.items():
        if d.hidden:
            h.update(range(d.min, d.max + 1))
    return h


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def r2(v, d=2):
    v = num(v)
    return None if v is None else (round(v, d) if d else int(round(v)))


def iso_dot(s, ym):
    """'26.02.24' · '26.10.8' · datetime → yyyy-mm-dd"""
    if isinstance(s, datetime.datetime):
        return s.strftime('%Y-%m-%d')
    m = re.match(r'\s*(\d{2})\.(\d{1,2})\.(\d{1,2})', str(s or ''))
    return f'20{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}' if m else ''


# ---------------- TV 요약 → items · tot · period ----------------
ws = wb['TV 요약']
period = str(ws.cell(3, 2).value).strip()
m = re.match(r'(\d{2})\.(\d{1,2})', period)
YEAR, MON = 2000 + int(m.group(1)), int(m.group(2))
ym = f'{YEAR}-{MON:02d}'
items, cur, tot = [], None, None
ONAIR = {}
for r in range(10, ws.max_row + 1):
    a, b, c = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
    if ws.row_dimensions[r].hidden:
        # 숨긴 줄 — 품목 이름만 기억(첫 소재 줄이 숨겨져 있어도 다음 보이는 소재가 그 품목에 붙도록)
        if b not in (None, '') and a not in (None, '') and str(a).strip() not in ('소계', '`', '계'):
            cur = {'n': canon(a), 'raw': str(a).strip(), 'cr': [], 'sub': None}
            items.append(cur)
        elif b in (None, '') and str(a or '').strip() in ('소계', '`'):
            cur = None
        continue
    vals = [ws.cell(r, k).value for k in range(4, 13)]
    if str(a or '').strip() == '계':
        tot = {'eq': r2(vals[2]), 'r1': num(vals[3]), 'r3': num(vals[4]), 'bud': r2(vals[0], 0), 'grp': r2(vals[1])}
        continue
    if b not in (None, ''):          # 소재 줄
        if a not in (None, '') and str(a).strip() not in ('소계', '`'):
            cur = {'n': canon(a), 'raw': str(a).strip(), 'cr': [], 'sub': None}
            items.append(cur)
        if cur is None:
            continue
        cr = {'n': str(b).strip(), 'on': iso_dot(c, ym), 'bud': r2(vals[0], 0), 'grp': r2(vals[1]), 'eq': r2(vals[2]),
              'r1': num(vals[3]), 'r3': num(vals[4]),
              'c': {'grp': r2(vals[5]), 'eq': r2(vals[6]), 'r1': num(vals[7]), 'r3': num(vals[8])}}
        cur['cr'].append(cr)
        ONAIR[(cur['n'], cr['n'])] = cr['on']
    elif cur is not None and any(v not in (None, '') for v in vals):   # 소계 줄(이름 칸은 '소계' · '`' · 빈칸)
        cur['sub'] = {'bud': r2(vals[0], 0), 'grp': r2(vals[1]), 'eq': r2(vals[2]), 'r1': num(vals[3]), 'r3': num(vals[4]),
                      'c': {'grp': r2(vals[5]), 'eq': r2(vals[6]), 'r1': num(vals[7]), 'r3': num(vals[8])}}
        cur = None
items = [i for i in items if i['cr']]
for i in items:   # 소계 줄이 없으면(소재 하나) 그 소재 값
    if i['sub'] is None:
        x = i['cr'][0]
        i['sub'] = {k: x[k] for k in ('bud', 'grp', 'eq', 'r1', 'r3', 'c')}

# ---------------- 당월_주별 → chs · weeks · wtot ----------------
ws = wb['당월_주별']
hc = hidden_cols(ws)
weeks, wcols = [], []
for col in range(13, ws.max_column + 1, 6):
    h = ws.cell(9, col).value
    if not h or col in hc:
        continue
    mm = re.match(r'\s*(\d+)주\s*\((\d+)(?:~(\d+))?일\)', str(h))
    if not mm:
        continue
    s, e = int(mm.group(2)), int(mm.group(3) or mm.group(2))
    weeks.append({'l': f'{mm.group(1)}주', 's': f'{ym}-{s:02d}', 'e': f'{ym}-{e:02d}'})
    wcols.append(col)


def wk(row):
    return [[r2(ws.cell(row, c).value), r2(ws.cell(row, c + 1).value, 0)] for c in wcols]


rows_all = []          # (row, g, pp, ch, hidden)
g, pp_rows = None, []
for r in range(12, ws.max_row + 1):
    a, b, c = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
    if a not in (None, ''):
        g = '지상파' if '지상파' in str(a) else '케이블'
    bs = str(b or '').strip()
    if 'TV Set Total' in bs:
        wtot = {'w': wk(r), 'eq': r2(ws.cell(r, 7).value), 'bud': r2(ws.cell(r, 6).value, 0), 'cost': r2(ws.cell(r, 8).value, 0)}
        break
    if 'Total' in bs:
        continue
    cs = str(c or '').strip()
    if not cs:
        continue
    if cs.endswith('Total'):   # PP Total 줄 → 그 위 채널들의 PP 이름
        ppn = cs[:-5].strip()
        for x in pp_rows:
            x[2] = ppn
        pp_rows = []
        continue
    rec = [r, g, None, cs, bool(ws.row_dimensions[r].hidden)]
    if g == '지상파':
        rec[2] = cs
    else:
        pp_rows.append(rec)
    rows_all.append(rec)

# ---------------- 당월_품목별 세부 → det · dorder ----------------
wd = wb['당월_품목별 세부']
hcd = hidden_cols(wd)
blocks, item = [], None
for col in range(5, wd.max_column + 1):
    if col in hcd:
        continue
    if wd.cell(10, col).value != '횟수':
        continue
    t7 = wd.cell(7, col).value
    if t7 not in (None, ''):
        item = str(t7).strip()
    lab = wd.cell(9, col).value
    if item is None:
        continue
    if item == '코웨이 총 합계':
        blocks.append(('__ALL__', item, '계', '코웨이 총 합계', col))
        continue
    if lab in (None, ''):
        continue
    lab = str(lab).strip()
    if lab.endswith(' 계'):
        blocks.append((canon(item), item, '계', lab, col))
    else:
        blocks.append((canon(item), item, lab, lab, col))
# 채널 줄 · 합계 줄
det_ch, det_tot = [], {}
for r in range(11, wd.max_row + 1):
    b, c = str(wd.cell(r, 2).value or '').strip(), str(wd.cell(r, 3).value or '').strip()
    if 'TV Set Total' in b:
        det_tot['T'] = r
        break
    if '지상파TV Total' in b:
        det_tot['지상파'] = r
        continue
    if '케이블TV Total' in b:
        det_tot['케이블'] = r
        continue
    if not c or c.endswith('Total') or wd.row_dimensions[r].hidden:
        continue
    det_ch.append((r, c))


def cell5(r, col):
    v = [wd.cell(r, col + k).value for k in range(5)]
    return [r2(v[0], 0) or 0, r2(v[1], 3) or 0, r2(v[2], 3) or 0, num(v[3]) or 0, num(v[4]) or 0]


# 품목 세부에 실적이 있는 채널은 주별에서 숨겨져 있어도 chs 에 넣는다
need = set()
for (_, _, _, _, col) in blocks:
    for r, c in det_ch:
        x = cell5(r, col)
        if x[0] or x[1]:
            need.add(c.lower())
norm = lambda s: re.sub(r'\s+', '', s).lower()
chs = []
for r, g_, pp, ch, hid in rows_all:
    if hid and norm(ch) not in {norm(n) for n in need}:
        continue
    chs.append({'g': g_, 'pp': pp, 'ch': ch, 'tc': r2(ws.cell(r, 4).value, 0) if num(ws.cell(r, 4).value) else None,
                'bud': r2(ws.cell(r, 6).value, 0), 'cost': r2(ws.cell(r, 8).value, 0), 'eq': r2(ws.cell(r, 7).value), 'w': wk(r)})
cidx = {norm(c['ch']): i for i, c in enumerate(chs)}
missing = sorted({c for _, c in det_ch if norm(c) not in cidx and any(cell5(r, b[4])[0] for b in blocks for rr, cc in [(r, c)])})

det, dorder, order_items = {}, [], {}
for (it, raw_it, n, raw, col) in blocks:
    t = {k: cell5(r, col) for k, r in det_tot.items()}
    cells = []
    for r, c in det_ch:
        x = cell5(r, col)
        if (x[0] or x[1]) and norm(c) in cidx:
            cells.append([cidx[norm(c)]] + x)
    k = f'{it}|{n}'
    if t.get('T', [0])[0] or cells:
        det[k] = {'cells': cells, 't': {'T': t['T'], '지상파': t['지상파'], '케이블': t['케이블']}}
    if it not in order_items:
        order_items[it] = {'n': it, 'raw': raw_it if it != '__ALL__' else '코웨이 총 합계', 'cr': []}
        dorder.append(order_items[it])
    e = {'k': k, 'n': n, 'on': ONAIR.get((it, n), '') if n != '계' else '', 'raw': raw}
    if n != '계':
        e['sn'] = n
    order_items[it]['cr'].append(e)
# __ALL__ 은 맨 뒤(기존과 같게)
dorder.sort(key=lambda d: d['n'] == '__ALL__')

# ---------------- 당월_시청률 → rt ----------------
wr = wb['당월_시청률']
rt = {}
for g_, c0 in (('지상파', 1), ('종편', 12), ('케이블', 23)):
    lst = []
    for r in range(5, wr.max_row + 1, 2):
        ch = wr.cell(r, c0).value
        if ch in (None, ''):
            continue
        f = lambda rr, k: (lambda v: None if v is None else round(v, 3))(num(wr.cell(rr, c0 + k).value))
        lst.append({'ch': str(ch).strip(), 'cw': f(r, 8), 'avg': f(r + 1, 8),
                    'cww': [f(r, 2 + i) for i in range(6)], 'avgw': [f(r + 1, 2 + i) for i in range(6)]})
    rt[g_] = lst

# ---------------- 당월_광고비 기준별 효율 → eff ----------------
we = wb['당월_광고비 기준별 효율']
hdr = str(we.cell(9, 24).value or '')
mm = re.search(r'\((~[^)]+)\)', hdr)
eff = {'asof': mm.group(1) if mm else '', 'days': int(num(we.cell(6, 24).value) or 0),
       'span': int(num(we.cell(5, 2).value) or 0), 'prog': round(num(we.cell(7, 24).value) or 0, 4), 'rows': [], 'gt': {}}


def erow(r):
    tc, bud, eq = num(we.cell(r, 4).value), num(we.cell(r, 5).value), num(we.cell(r, 6).value)
    tgt = bud / tc if tc and bud else None
    return [round(tc, 1) if tc else tc, int(round(bud or 0)), round(eq or 0, 2),
            round(tgt, 2) if tgt else None, round(eq / tgt, 4) if tgt else None]


g_ = None
for r in range(10, we.max_row + 1):
    if we.row_dimensions[r].hidden:
        continue
    a, b, c = (str(we.cell(r, k).value or '').strip() for k in (1, 2, 3))
    if a.startswith('케이블'):
        g_ = '케이블'
    elif a == '지상파' or (g_ is None and c):
        g_ = g_ or '지상파'
    if 'TV Set Total' in a:
        x = erow(r)
        eff['tot'] = x
        continue
    if '지상파TV Total' in a:
        eff['gt']['지상파'] = erow(r)
        g_ = '케이블'
        continue
    if '케이블TV Total' in a:
        eff['gt']['케이블'] = erow(r)
        continue
    if g_ == '지상파' and c:
        eff['rows'].append(['지상파', c] + erow(r))
    elif g_ == '케이블' and c.endswith('Total'):
        eff['rows'].append(['케이블', c[:-5].strip()] + erow(r))

wr_out = {'ym': ym, 'period': period, 'items': items, 'tot': tot, 'chs': chs, 'weeks': weeks, 'wtot': wtot,
          'det': det, 'dorder': dorder, 'rt': rt, 'eff': eff}
json.dump(wr_out, open(OUT, 'w'), ensure_ascii=False)
print('period', period, 'items', [(i['n'], len(i['cr'])) for i in items])
print('tot', tot)
print('weeks', weeks)
print('wtot', wtot)
print('chs', len(chs), 'missing-in-chs', missing)
print('det keys', list(det))
print('eff', {k: v for k, v in eff.items() if k != 'rows'}, len(eff['rows']))
print('rt', {k: len(v) for k, v in rt.items()})
