import random, html, collections
exec(open('ch38_q.py').read())
DROP=["Which findings belong to the UPPER RIGHT","Which statements about S1 are correct","Per Box 38.2, which sounds point","Which factors can make ANY murmur louder","Which characteristics are included when describing","Which statements about murmur GRADE","Which statements about SYSTOLIC murmurs","A family is told their child has an innocent murmur","Which conditions are in the differential diagnosis of a Still","Which findings describe a VENOUS HUM","What information does a chest radiograph","Which measures help prevent avoidable","After endocarditis prophylaxis is given","Which situations warrant a genetic testing","Which maternal conditions are associated","Noonan syndrome is associated","Which CHF signs are listed for children and teens","Which are traditional pediatric heart failure","Which are complications of an UNTREATED ASD","Which findings may be found in a child with an ASD","Cyanosis is LESS readily apparent","Hypercyanotic (\"tet\") spells increase","Which tests detect ventricular arrhythmias","Which ELECTRICAL abnormalities","Which ACQUIRED causes","An ejection click may be confused","Up to what percentage","A murmur can be heard with the stethoscope held just off","Which innocent-murmur finding may make the murmur DISAPPEAR","Why might a CBC be ordered","Infant and child CPR certification","During which weeks of gestation","Which statement about B-type natriuretic","Which type of VSD is most common","Which ECG change develops in an OLDER child"]
for dd in DROP: assert sum(c[1].startswith(dd) for c in C)==1, dd
C=[c for c in C if not any(c[1].startswith(d) for d in DROP)]; assert len(C)==100, len(C)
rng=random.Random(8); L="ABCDEFG"
singles=[i for i,c in enumerate(C) if 'apply' not in c[1]]
tg=list("ABCD"*14)[:len(singles)]; rng.shuffle(tg); tmap=dict(zip(singles,tg))
items=[]
for i,(t,stem,ok,bad,why,mn) in enumerate(C):
    sata='apply' in stem
    if sata: o=ok+bad; rng.shuffle(o)
    else: b=bad[:]; rng.shuffle(b); k=L.index(tmap[i]); o=b[:k]+ok+b[k:]
    items.append(dict(t=t,stem=stem,opts=o,ans=[L[j] for j,x in enumerate(o) if x in ok],why=why,mn=mn,sata=sata))
print(collections.Counter(x['ans'][0] for x in items if not x['sata']), sum(x['sata'] for x in items))
TITLE="NUR 774 Chapter 38 Cardiovascular: 100 Hard Questions"
SUB="45 Select All That Apply + 55 single answer, grouped by topic, from Chapter 38 (heart sounds, murmurs, diagnostics, primary care, specific defects, sudden cardiac death). Each question is followed by its answer, rationale, and a mnemonic."
NOTE="Heads-up: Chapter 38 says delay live vaccines 6 months after bypass/blood products. That is different from the 11-month wait after Kawasaki IVIG (2 g/kg). Know both."
# ---------- web (tap to reveal)
e=html.escape; W=[]; cur=None
for n,x in enumerate(items,1):
    if x['t']!=cur: cur=x['t']; W.append(f'<h2>{e(cur)}</h2>')
    tag='<span class="sata">SELECT ALL THAT APPLY</span>' if x['sata'] else '<span class="single">SINGLE ANSWER</span>'
    W.append(f'<div class="q"><div class="n">Question {n} of 100 {tag}</div><p class="stem">{e(x["stem"].replace(SA,""))}</p><ol type="A">'+''.join(f'<li>{e(o)}</li>' for o in x['opts'])+
      f'</ol><details><summary>Tap to show answer</summary><div class="ans"><p class="key">✅ {", ".join(x["ans"])}</p><ul>'+''.join(f'<li>{e(x["opts"][L.index(a)])}</li>' for a in x['ans'])+
      f'</ul><p><b>Why:</b> {e(x["why"])}</p><div class="mn">🧠 {e(x["mn"])}</div></div></details></div>')
css="""body{font-family:-apple-system,Arial,sans-serif;max-width:640px;margin:0 auto;padding:14px;background:#f4f6fb;font-size:16px;line-height:1.4}
h1{color:#1f3864;font-size:22px}h2{color:#fff;background:#2f5496;padding:6px 10px;border-radius:8px;font-size:17px;margin-top:22px}
.q{background:#fff;border-radius:12px;padding:12px 14px;margin:12px 0;box-shadow:0 1px 3px #0002}.n{color:#666;font-size:13px}
.sata,.single{display:inline-block;font-size:11px;font-weight:bold;padding:2px 7px;border-radius:9px;color:#fff;background:#c0392b;margin-left:4px}.single{background:#2f5496}
.stem{font-weight:bold}ol{padding-left:24px}ol li{margin:6px 0}
summary{background:#1f7a3a;color:#fff;padding:10px;border-radius:8px;text-align:center;font-weight:bold;cursor:pointer;list-style:none}summary::-webkit-details-marker{display:none}
details[open] summary{background:#888}.ans{background:#eef3fb;border-radius:8px;padding:8px 12px;margin-top:8px}.key{font-size:22px;font-weight:bold;color:#1f7a3a;margin:2px 0}
.mn{background:#fff4cc;border-left:4px solid #e0a800;padding:7px 9px;font-weight:bold;border-radius:4px}.note{background:#fdecea;border-left:4px solid #c0392b;padding:8px;border-radius:4px;font-size:14px}
.btns button{padding:8px 12px;margin:4px 4px 0 0;border-radius:8px;border:1px solid #2f5496;background:#fff;color:#2f5496;font-weight:bold}"""
js="<script>function all(o){document.querySelectorAll('details').forEach(d=>d.open=o)}</script>"
open('ch38.html','w').write(f'<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ch 38 Cardio 100 Qs</title><style>{css}</style>{js}</head><body><h1>{TITLE}</h1><p>{SUB}</p><p class="note">{e(NOTE)}</p><div class="btns"><button onclick="all(false)">Hide all answers</button><button onclick="all(true)">Show all answers</button></div>{"".join(W)}<p style="text-align:center"><a href="index.html">← All study files</a></p></body></html>')
# ---------- Word (answer box right after, cover with paper)
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
def shade(p,c):
    s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),c); p._p.get_or_add_pPr().append(s)
d=Document(); st=d.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(11)
d.add_heading(TITLE,0); d.add_paragraph(SUB+" Cover the shaded boxes with a sticky note or index card to quiz yourself.").runs[0].italic=True
p=d.add_paragraph(); shade(p,'FDECEA'); p.add_run(NOTE).bold=True
cur=None
for n,x in enumerate(items,1):
    if x['t']!=cur: cur=x['t']; d.add_heading(cur,1)
    q=d.add_paragraph(); q.paragraph_format.keep_with_next=True; q.paragraph_format.space_before=Pt(8)
    r=q.add_run(f'{n}. '); r.bold=True
    if x['sata']: r=q.add_run('[SATA] '); r.bold=True; r.font.color.rgb=RGBColor(0xC0,0x39,0x2B)
    q.add_run(x['stem']).bold=True
    for l,o in zip(L,x['opts']):
        op=d.add_paragraph(f'   {l}. {o}'); op.paragraph_format.space_after=Pt(0); op.paragraph_format.keep_with_next=True
    a=d.add_paragraph(); shade(a,'E8EEF7'); a.paragraph_format.space_before=Pt(4); a.paragraph_format.space_after=Pt(0); a.paragraph_format.keep_with_next=True
    a.add_run('Answer: ').bold=True; a.add_run(', '.join(x['ans'])+': '+'; '.join(x['opts'][L.index(k)] for k in x['ans']))
    w=d.add_paragraph(); shade(w,'E8EEF7'); w.paragraph_format.space_after=Pt(0); w.paragraph_format.keep_with_next=True; w.add_run('Rationale: ').bold=True; w.add_run(x['why'])
    m=d.add_paragraph(); shade(m,'FFF4CC'); m.paragraph_format.space_after=Pt(10); m.add_run('🧠 Mnemonic: ').bold=True; m.add_run(x['mn'])
d.save('NUR774_Ch38_Cardio_100Q.docx')
# ---------- phone PDF flashcards (answer on next page)
P=[f'<section class="card"><h1>{TITLE}</h1><p>{SUB}</p><p class="note">{e(NOTE)}</p><p><b>How to use:</b> answer the question, then swipe to the next page for the answer, rationale, and mnemonic.</p></section>']
for n,x in enumerate(items,1):
    tag='<span class="sata">SELECT ALL THAT APPLY</span>' if x['sata'] else '<span class="single">SINGLE ANSWER</span>'
    P.append(f'<section class="card"><div class="num">Question {n} of 100 · {e(x["t"])}</div>{tag}<p class="stem">{e(x["stem"].replace(SA,""))}</p><ol type="A">'+''.join(f'<li>{e(o)}</li>' for o in x['opts'])+'</ol><div class="next">Swipe for answer →</div></section>')
    P.append(f'<section class="card a"><div class="num">Answer {n}</div><p class="key">{", ".join(x["ans"])}</p><ul>'+''.join(f'<li><b>{e(x["opts"][L.index(k)])}</b></li>' for k in x['ans'])+f'</ul><p><b>Why:</b> {e(x["why"])}</p><div class="mn">🧠 {e(x["mn"])}</div></section>')
pcss="""@page{size:4in 7.2in;margin:0}body{margin:0;font-family:Arial,sans-serif}.card{box-sizing:border-box;width:4in;min-height:7.2in;padding:.25in;break-before:page;font-size:12pt;line-height:1.35}
h1{color:#1f3864;font-size:19pt}.num{color:#666;font-size:9.5pt;margin-bottom:6px}.stem{font-weight:bold}ol{padding-left:22px}ol li{margin:6px 0}
.sata,.single{display:inline-block;font-size:8.5pt;font-weight:bold;padding:2px 7px;border-radius:9px;color:#fff;background:#c0392b}.single{background:#2f5496}
.next{color:#999;font-size:9pt;text-align:right;margin-top:10px}.a{background:#eef3fb}.key{font-size:24pt;font-weight:bold;color:#1f7a3a;margin:4px 0}ul{padding-left:18px}li{margin:3px 0}
.mn{background:#fff4cc;border-left:4px solid #e0a800;padding:7px 9px;font-weight:bold}.note{background:#fdecea;border-left:4px solid #c0392b;padding:7px;font-size:10.5pt}"""
open('ch38_phone.html','w').write(f'<html><head><meta charset="utf-8"><style>{pcss}</style></head><body>{"".join(P)}</body></html>')
