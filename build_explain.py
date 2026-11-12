import re, html
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
src=open('explain.md').read().splitlines()
def runs(p,t):
    for i,part in enumerate(re.split(r'\*\*',t)):
        for j,q in enumerate(re.split(r'(?<!\w)_(.+?)_(?!\w)',part)):
            if q: r=p.add_run(q); r.bold=i%2==1; r.italic=j%2==1
def shade(p,c):
    s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),c); p._p.get_or_add_pPr().append(s)
def inl(t):
    t=html.escape(t,quote=False); t=re.sub(r'\*\*(.+?)\*\*',r'<b>\1</b>',t); return re.sub(r'(?<!\w)_(.+?)_(?!\w)',r'<i>\1</i>',t)
d=Document(); st=d.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(11)
H=[]; i=0; lst=None
def close():
    global lst
    if lst: H.append(f'</{lst}>'); lst=None
while i<len(src):
    l=src[i]
    if not l.strip(): close(); i+=1; continue
    if l.startswith('|'):
        close(); rows=[]
        while i<len(src) and src[i].startswith('|'):
            if not re.match(r'^\|[-| ]+\|$',src[i]): rows.append([c.strip() for c in src[i].strip('|').split('|')])
            i+=1
        t=d.add_table(rows=len(rows),cols=len(rows[0])); t.style='Table Grid'
        for r,row in enumerate(rows):
            for c,val in enumerate(row):
                cell=t.cell(r,c); cell.text=''; p=cell.paragraphs[0]; runs(p,val)
                for rr in p.runs: rr.font.size=Pt(9.5); rr.bold = rr.bold or r==0
                if r==0: shade(p,'D9E2F3')
        d.add_paragraph()
        H.append('<table>'+''.join('<tr>'+''.join(f'<{"th" if r==0 else "td"}>{inl(v)}</{"th" if r==0 else "td"}>' for v in row)+'</tr>' for r,row in enumerate(rows))+'</table>')
        continue
    if l.startswith('# '): close(); d.add_heading(l[2:],0); H.append(f'<h1>{inl(l[2:])}</h1>')
    elif l.startswith('## '): close(); d.add_heading(l[3:],1); H.append(f'<h2>{inl(l[3:])}</h2>')
    elif l.startswith('### '): close(); d.add_heading(l[4:],2); H.append(f'<h3>{inl(l[4:])}</h3>')
    elif l.startswith('> '): close(); p=d.add_paragraph(); shade(p,'FFF4CC'); runs(p,l[2:]); H.append(f'<div class="mn">{inl(l[2:])}</div>')
    elif l.startswith('- '):
        if lst!='ul': close(); H.append('<ul>'); lst='ul'
        runs(d.add_paragraph(style='List Bullet'),l[2:]); H.append(f'<li>{inl(l[2:])}</li>')
    elif re.match(r'^\d+\. ',l):
        if lst!='ol': close(); H.append('<ol>'); lst='ol'
        txt=re.sub(r'^\d+\. ','',l); runs(d.add_paragraph(style='List Number'),txt); H.append(f'<li>{inl(txt)}</li>')
    else: close(); runs(d.add_paragraph(),l); H.append(f'<p>{inl(l)}</p>')
    i+=1
close()
d.save('NUR774_Murmurs_Circulation_Defects.docx')
css="""@page{size:4in 7.2in;margin:.22in .2in}body{font-family:Arial,sans-serif;font-size:11pt;line-height:1.35;margin:0}
h1{color:#1f3864;font-size:17pt;margin:0 0 6px}h2{background:#2f5496;color:#fff;padding:5px 8px;border-radius:6px;font-size:13.5pt;break-before:page;margin-top:0}
h3{color:#2f5496;font-size:12pt;margin:12px 0 4px}ul,ol{padding-left:18px;margin:3px 0 8px}li{margin:3px 0;break-inside:avoid}p{margin:6px 0}
.mn{background:#fff4cc;border-left:4px solid #e0a800;padding:6px 8px;margin:8px 0;break-inside:avoid}
table{border-collapse:collapse;width:100%;font-size:8.3pt;margin:6px 0}th,td{border:1px solid #9aa9c7;padding:3px;vertical-align:top}th{background:#d9e2f3}tr{break-inside:avoid}"""
open('explain_phone.html','w').write(f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{"".join(H)}</body></html>')
