"""Word's page map: per page, the body lines (top to bottom) inside the
section's margin box (headers/footers excluded by geometry)."""
import subprocess, sys, json, re, zipfile, html
def margins(docx):
    d = zipfile.ZipFile(docx).read('word/document.xml').decode()
    secs = re.findall(r'<w:sectPr\b.*?</w:sectPr>', d, re.S)
    out = []
    for s in secs:
        m = re.search(r'<w:pgMar ([^>]*)/>', s)
        a = dict(re.findall(r'w:(\w+)="(-?\d+)"', m.group(1))) if m else {}
        out.append({k: int(v) / 20 for k, v in a.items()})
    return out
def lines(pdf):
    x = subprocess.run(['pdftotext', '-bbox-layout', pdf, '-'], capture_output=True, text=True).stdout
    pages = []
    for pg in re.findall(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', x, re.S):
        w, h, body = float(pg[0]), float(pg[1]), pg[2]
        ls = []
        for ln in re.findall(r'<line xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</line>', body, re.S):
            words = [html.unescape(t) for t in re.findall(r'<word[^>]*>(.*?)</word>', ln[4], re.S)]
            ls.append({'x': float(ln[0]), 'y': float(ln[1]), 'y2': float(ln[3]), 'text': ' '.join(words)})
        ls.sort(key=lambda l: (round(l['y']), l['x']))
        pages.append({'w': w, 'h': h, 'lines': ls})
    return pages
if __name__ == '__main__':
    docx, pdf, out = sys.argv[1:4]
    mg = margins(docx)
    m = mg[-1] if mg else {}
    top = min(s.get('top', 72) for s in mg) if mg else 72
    bottom = min(s.get('bottom', 72) for s in mg) if mg else 72
    pages = lines(pdf)
    res = []
    for i, p in enumerate(pages):
        body = [l for l in p['lines'] if l['y2'] > top - 2 and l['y'] < p['h'] - bottom + 2]
        res.append({'page': i + 1, 'size': [p['w'], p['h']], 'body': [l['text'] for l in body], 'all': [ (round(l['y'],1), l['text']) for l in p['lines']]})
    json.dump({'margins': mg, 'pages': res}, open(out, 'w'), indent=1, ensure_ascii=False)
    for r in res: print(r['page'], '|', (r['body'][0] if r['body'] else '<EMPTY>')[:90])
