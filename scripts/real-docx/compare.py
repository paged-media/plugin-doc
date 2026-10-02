"""ADR 029 decision 6 (docs/reference/acceptance-real-docx.md): compare Word's page
map (its PDF) with ours (the editor's per-paragraph page geometry). A
page's START PARAGRAPH is the body
paragraph holding the page's first body line (a paragraph continuing from
the previous page counts). Empty paragraphs are invisible on both sides."""
import json, re, sys, unicodedata
LIG = {'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬀ': 'ff', 'ﬃ': 'ffi', 'ﬄ': 'ffl'}
def norm(s):
    s = ''.join(LIG.get(c, c) for c in s)
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s.lower() if c.isalnum())
def main(word_map, ours_json, out=None):
    W = json.load(open(word_map)); O = json.load(open(ours_json))
    # Word stream with page offsets.
    stream, page_at = '', []
    for p in W['pages']:
        page_at.append(len(stream))
        stream += norm(' '.join(p['body']))
    # Our paragraphs in document order.
    paras = []
    for s in O['stories']:
        for p in s['paras']:
            paras.append({'story': s['storyId'], 'text': p['text'], 'n': norm(p['text']), 'pages': p['pages']})
    # Greedy alignment of each non-empty paragraph into Word's stream.
    cur = 0
    for p in paras:
        p['wpos'] = None
        if not p['n']:
            continue
        key = p['n'][:40]
        win = 400 if len(key) < 8 else 60000
        i = stream.find(key, cur, cur + win + len(key))
        if i < 0 and len(key) >= 16:
            i = stream.find(key[:16], cur, cur + win)
        if i >= 0:
            p['wpos'] = i
            cur = i + min(len(p['n']), len(key))
    aligned = [k for k, p in enumerate(paras) if p['wpos'] is not None]
    nonempty = [k for k, p in enumerate(paras) if p['n']]
    # Word page starts: last aligned paragraph with wpos <= page start.
    word_start = []
    for pg, at in enumerate(page_at):
        best = None
        for k in aligned:
            if paras[k]['wpos'] <= at:
                best = k
            else:
                break
        word_start.append(best)
    npages_o = len(O['pages'])
    our_start = []
    for pg in range(npages_o):
        k = next((k for k in nonempty if pg in paras[k]['pages']), None)
        our_start.append(k)
    rows, ok = [], 0
    for pg in range(len(page_at)):
        w = word_start[pg]; o = our_start[pg] if pg < npages_o else None
        same = w is not None and w == o
        ok += same
        rows.append({'page': pg + 1, 'word': w, 'ours': o, 'same': same,
                     'word_text': paras[w]['text'][:70] if w is not None else None,
                     'ours_text': paras[o]['text'][:70] if o is not None else None,
                     'word_first_line': (W['pages'][pg]['body'] or [''])[0][:70]})
    res = {'word_pages': len(page_at), 'our_pages': npages_o, 'matched': ok,
           'pct': round(100 * ok / len(page_at), 1), 'aligned': len(aligned), 'nonempty': len(nonempty), 'rows': rows}
    if out:
        json.dump(res, open(out, 'w'), indent=1, ensure_ascii=False)
    print(f"word {res['word_pages']} pages, ours {npages_o}; page-start match {ok}/{len(page_at)} = {res['pct']}%; aligned {len(aligned)}/{len(nonempty)} paragraphs")
    for r in rows:
        if not r['same']:
            d = (r['ours'] - r['word']) if (r['ours'] is not None and r['word'] is not None) else None
            print(f"  p{r['page']:>3} word#{r['word']} ours#{r['ours']} (Δ{d}) | W: {r['word_text']!r} | O: {r['ours_text']!r}")
if __name__ == '__main__':
    main(*sys.argv[1:])
