import re, html
from docx import Document
from docx.shared import Pt, RGBColor
src = open('guide.md').read().splitlines()
def runs(par, text):
    for i, part in enumerate(re.split(r'\*\*', text)):
        if not part: continue
        for j, p2 in enumerate(re.split(r'(?<!\w)_(.+?)_(?!\w)', part)):
            r = par.add_run(p2); r.bold = (i % 2 == 1); r.italic = (j % 2 == 1)
doc = Document()
st = doc.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(11)
H = []; title = True
for l in src:
    if not l.strip(): continue
    if l.startswith('# '):
        if title: doc.add_heading(l[2:], 0); title = False
        else: doc.add_heading(l[2:], 1)
        H.append(f'<h{1 if len(H)==0 else 2}>{html.escape(l[2:])}</h{1 if len(H)==0 else 2}>')
    elif l.startswith('## '):
        doc.add_heading(l[3:], 2); H.append(f'<h3>{html.escape(l[3:])}</h3>')
    elif l.startswith('- '):
        runs(doc.add_paragraph(style='List Bullet'), l[2:])
        H.append('<li>' + l[2:] + '</li>')
    else:
        runs(doc.add_paragraph(), l); H.append('<p>' + l + '</p>')
doc.save('NUR774_Exam3_Mnemonics.docx')
body = '\n'.join(H)
body = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', body)
body = re.sub(r'(?<!\w)_(.+?)_(?!\w)', r'<i>\1</i>', body)
body = re.sub(r'(</li>)\n(?!<li>)', r'\1</ul>\n', body)
body = re.sub(r'(?<!</li>)\n<li>', '\n<ul><li>', body)
css = """body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.35;margin:0}
h1{font-size:22pt;color:#1f3864}h2{font-size:15pt;color:#fff;background:#2f5496;padding:4px 8px;margin-top:18px}
h3{font-size:12.5pt;color:#2f5496;border-bottom:1px solid #2f5496;margin:12px 0 4px}ul{margin:2px 0 6px;padding-left:20px}li{margin:2px 0}
h3,li{break-inside:avoid}"""
open('guide.html','w').write(f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{body}</body></html>')
