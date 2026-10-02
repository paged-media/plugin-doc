"""Line geometry, Word vs ours, per aligned paragraph: how many lines each
paragraph takes and where they start and end (Word's from its PDF, ours
from measure.spec.ts's per-line rects).
usage: line_geometry.py word.pdf ours.json [rows] [filter-substring]"""
import json, re, sys, subprocess, collections, html, bisect
sys.path.insert(0, sys.argv[0].rsplit('/', 1)[0])
from compare import norm
pdf, ours = sys.argv[1], sys.argv[2]
N = int(sys.argv[3]) if len(sys.argv) > 3 else 30
FILT = sys.argv[4] if len(sys.argv) > 4 else None
x = subprocess.run(['pdftotext', '-bbox-layout', pdf, '-'], capture_output=True, text=True).stdout
O = json.load(open(ours))
P0 = O['pages'][0]; mt = P0['marginTopPt']; ph = P0['sizePt'][1]; mb = P0['marginBottomPt']
wl = []  # word visual lines: dict(page, y, y2, x1, x2, text, off)
stream = ''
for pi, p in enumerate(re.split(r'<page ', x)[1:]):
    ws = []
    for a, b, c, d, t in re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', p):
        a, b, c, d = map(float, (a, b, c, d))
        if b < mt - 2 or d > ph - mb + 14: continue
        ws.append((b, a, c, d, html.unescape(t)))
    ws.sort()
    rows = []
    for b, a, c, d, t in ws:
        if rows and abs(rows[-1]['y'] - b) < 4 or (rows and b < rows[-1]['y2'] - 6):
            r = rows[-1]; r['w'].append((a, c, t)); r['y2'] = max(r['y2'], d); continue
        rows.append({'page': pi, 'y': b, 'y2': d, 'w': [(a, c, t)]})
    for r in rows:
        r['w'].sort(); r['x1'] = r['w'][0][0]; r['x2'] = max(w[1] for w in r['w']); r['text'] = ' '.join(w[2] for w in r['w'])
        n = norm(r['text'])
        if n: r['off'] = len(stream); stream += n; wl.append(r)
offs = [r['off'] for r in wl]
for r in wl: r['end'] = r['off'] + len(norm(r['text']))
ends = [r['end'] for r in wl]
paras = [p for s in O['stories'] for p in s['paras']]
cur = 0
for p in paras:
    p['n'] = norm(p['text']); p['wpos'] = None
    if not p['n']: continue
    key = p['n'][:40]; win = 400 if len(key) < 8 else 60000
    i = stream.find(key, cur, cur + win + len(key))
    if i < 0 and len(key) >= 16: i = stream.find(key[:16], cur, cur + win)
    if i >= 0: p['wpos'] = i; cur = i + min(len(p['n']), len(key))
al = [p for p in paras if p['wpos'] is not None and p.get('rects')]
def olines(p):
    m = collections.OrderedDict()
    for pg, t, l, w, h in p['rects']:
        k = (pg, round(t))
        if k in m: m[k][3] = max(m[k][3], l + w); m[k][2] = min(m[k][2], l)
        else: m[k] = [pg, t, l, l + w, h]
    return list(m.values())
cnt = collections.Counter(); diffs = []; gaps = collections.defaultdict(list)
tw = to = 0
for k, p in enumerate(al):
    end = al[k + 1]['wpos'] if k + 1 < len(al) else len(stream)
    # A line belongs to the paragraph holding its END (Word's list number
    # opens the line before the paragraph's own first character).
    lo = bisect.bisect_right(ends, p['wpos']); hi = bisect.bisect_right(ends, end)
    W = wl[lo:hi]
    p['W'] = W; p['L'] = olines(p)
    d = len(p['L']) - len(W); cnt[d] += 1; tw += len(W); to += len(p['L'])
    if d: diffs.append((d, p))
print('paragraphs', len(al), 'word lines', tw, 'ours', to, 'equal', cnt[0])
print('delta histogram', sorted(cnt.items()))
by = collections.Counter()
for d, p in diffs: by[p['style'].split('/')[-1]] += d
print('net by style', by.most_common(10), by.most_common()[-6:])
shown = 0
for d, p in sorted(diffs, key=lambda r: -abs(r[0])):
    if FILT and FILT not in p['text'] and FILT not in p['style']: continue
    if shown >= N: break
    shown += 1
    print(f"\n[{d:+d}] {p['style'].split('/')[-1]}  word p{p['W'][0]['page']+1 if p['W'] else '?'} ours p{p['L'][0][0]+1}: {p['text'][:70]!r}")
    for r in p['W'][:6]: print(f"   W x {r['x1']:6.1f}..{r['x2']:6.1f} y {r['y']:6.1f}  {r['text'][:80]}")
    for l in p['L'][:7]: print(f"   O x {l[2]:6.1f}..{l[3]:6.1f} y {l[1]:6.1f} h {l[4]}")
