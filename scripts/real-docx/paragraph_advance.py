"""Vertical advance between consecutive paragraphs, Word vs ours.
For neighbours (a, b) in OUR paragraph order that are both aligned into
Word's text and sit on one page on each side: advance = top(b) - top(a).
usage: paragraph_advance.py word.pdf ours.json [rows] [min abs delta]"""
import json, re, sys, subprocess, collections, html, bisect
sys.path.insert(0, sys.argv[0].rsplit('/', 1)[0])
from compare import norm
pdf, ours = sys.argv[1], sys.argv[2]
N = int(sys.argv[3]) if len(sys.argv) > 3 else 40
MIN = float(sys.argv[4]) if len(sys.argv) > 4 else 1.0
x = subprocess.run(['pdftotext', '-bbox-layout', pdf, '-'], capture_output=True, text=True).stdout
O = json.load(open(ours))
P0 = O['pages'][0]; mt = P0['marginTopPt']; ph = P0['sizePt'][1]; mb = P0['marginBottomPt']
wl = []; stream = ''
for pi, p in enumerate(re.split(r'<page ', x)[1:]):
    ws = []
    for a, b, c, d, t in re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', p):
        a, b, c, d = map(float, (a, b, c, d))
        if b < mt - 2 or d > ph - mb + 14: continue
        ws.append((b, a, c, d, html.unescape(t)))
    ws.sort(); rows = []
    for b, a, c, d, t in ws:
        if rows and (abs(rows[-1]['y'] - b) < 4 or b < rows[-1]['y2'] - 6):
            r = rows[-1]; r['w'].append((a, c, t)); r['y2'] = max(r['y2'], d); continue
        rows.append({'page': pi, 'y': b, 'y2': d, 'w': [(a, c, t)]})
    for r in rows:
        r['w'].sort(); r['text'] = ' '.join(w[2] for w in r['w']); n = norm(r['text'])
        if n: r['off'] = len(stream); stream += n; r['end'] = len(stream); wl.append(r)
ends = [r['end'] for r in wl]
paras = []
for s in O['stories']:
    for i, p in enumerate(s['paras']): p['story'] = s['storyId']; p['i'] = i; paras.append(p)
cur = 0
for p in paras:
    p['n'] = norm(p['text']); p['wpos'] = None
    if not p['n']: continue
    key = p['n'][:40]; win = 400 if len(key) < 8 else 60000
    i = stream.find(key, cur, cur + win + len(key))
    if i < 0 and len(key) >= 16: i = stream.find(key[:16], cur, cur + win)
    if i >= 0: p['wpos'] = i; cur = i + min(len(p['n']), len(key))
for p in paras:
    if p['wpos'] is None or not p.get('rects'): continue
    k = bisect.bisect_right(ends, p['wpos'])
    if k < len(wl): p['w'] = wl[k]
rows = []; by = collections.defaultdict(lambda: [0, 0.0]); tw = to = 0
prev = None; between = 0
for p in paras:
    if 'w' not in p:
        between += 1 if not p['n'] else 100   # blanks count 1; unaligned text poisons the pair
        continue
    if prev is not None and prev['story'] == p['story'] and between < 100:
        a, b = prev, p
        if a['w']['page'] == b['w']['page'] and a['rects'][0][0] == b['rects'][0][0]:
            hw = b['w']['y'] - a['w']['y']; ho = b['rects'][0][1] - a['rects'][0][1]
            if hw > 0 and ho > 0:
                d = ho - hw; tw += hw; to += ho
                key = (a['style'].split('/')[-1], b['style'].split('/')[-1], between)
                by[key][0] += 1; by[key][1] += d
                rows.append((round(d, 1), round(hw, 1), round(ho, 1), between, a['w']['page'] + 1, a['rects'][0][0] + 1, key[0], key[1], a['text'][:38], b['text'][:24]))
    prev = p; between = 0
print('pairs', len(rows), 'word', round(tw), 'ours', round(to), 'ratio', round(to / tw, 4) if tw else '-', 'pairs off by >', MIN, ':', sum(1 for r in rows if abs(r[0]) > MIN))
print('net by (style a, style b, blanks between) [n, sum pt]:')
for k, v in sorted(by.items(), key=lambda kv: -abs(kv[1][1]))[:16]: print('  ', k, v[0], round(v[1], 1))
for r in sorted(rows, key=lambda r: -abs(r[0]))[:N]:
    if abs(r[0]) >= MIN: print(' ', r)
