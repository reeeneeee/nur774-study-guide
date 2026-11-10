import random, html, collections
exec(open('ear_data.py').read())
rng=random.Random(2026); L="ABCDEFG"
singles=[(e,k) for e in E for k,q in enumerate(e['qs']) if not q[4]]
targets=list("ABCD"*6)[:len(singles)]; rng.shuffle(targets)
tmap={ (id(e),k):t for (e,k),t in zip(singles,targets)}
for e in E:
    new=[]
    for k,(stem,opts,ans,why,sata) in enumerate(e['qs']):
        correct=[opts[L.index(a.strip())] for a in ans.split(',')]
        wrong=[o for o in opts if o not in correct]
        if sata:
            o2=opts[:]; rng.shuffle(o2)
        else:
            rng.shuffle(wrong); t=L.index(tmap[(id(e),k)]); o2=wrong[:t]+correct+wrong[t:]
        a2=', '.join(L[i] for i,o in enumerate(o2) if o in correct)
        new.append((stem,o2,a2,why,sata))
    e['qs']=new
print(collections.Counter(q[2] for e in E for q in e['qs'] if not q[4]), [q[2] for e in E for q in e['qs'] if q[4]])
TITLE="NUR 774: Ear Pictures Explained"
SUB="Every picture from the Normal and Abnormal Middle Ear Anatomy slides: what it is, what you're looking at, a mnemonic, and practice questions with answers and rationales."
# ---------- Word
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_BREAK
def shade(p,c):
    s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),c); p._p.get_or_add_pPr().append(s)
d=Document(); st=d.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(11)
d.add_heading(TITLE,0); d.add_paragraph(SUB).runs[0].italic=True
d.add_heading('Quick ID cheat sheet',1)
for e in E: 
    p=d.add_paragraph(style='List Bullet'); p.add_run(f"Slide {e['s']}: ").bold=True; p.add_run(f"{e['title']}. ").bold=False; p.add_run(e['mn'].split('.')[0]+'.').italic=True
qn_=0
for e in E:
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading(f"Slide {e['s']}: {e['title']}",1)
    d.add_picture(f"ear_img/slide{e['s']}.jpg", width=Inches(3.6))
    c=d.add_paragraph(e['credit']); c.runs[0].italic=True; c.runs[0].font.size=Pt(8)
    p=d.add_paragraph(); p.add_run('What it is: ').bold=True; p.add_run(e['what'])
    d.add_paragraph().add_run('What you see in the picture:').bold=True
    for s in e['see']: d.add_paragraph(s, style='List Bullet')
    m=d.add_paragraph(); shade(m,'FFF4CC'); m.add_run('🧠 Mnemonic: ').bold=True; m.add_run(e['mn'])
    d.add_paragraph().add_run('Nursing / exam tips:').bold=True
    for s in e['tip']: d.add_paragraph(s, style='List Bullet')
    for stem,opts,ans,why,sata in e['qs']:
        qn_+=1
        q=d.add_paragraph(); q.paragraph_format.keep_with_next=True; q.paragraph_format.space_before=Pt(8)
        r=q.add_run(f'Q{qn_}. '); r.bold=True
        if sata: r=q.add_run('[SATA] '); r.bold=True; r.font.color.rgb=RGBColor(0xC0,0x39,0x2B)
        q.add_run(stem).bold=True
        for l,o in zip(L,opts):
            op=d.add_paragraph(f'   {l}. {o}'); op.paragraph_format.space_after=Pt(0); op.paragraph_format.keep_with_next=True
        a=d.add_paragraph(); shade(a,'E8EEF7'); a.paragraph_format.space_before=Pt(4); a.paragraph_format.space_after=Pt(0); a.paragraph_format.keep_with_next=True
        a.add_run('Answer: ').bold=True; a.add_run(ans+': '+'; '.join(opts[L.index(x.strip())] for x in ans.split(',')))
        w=d.add_paragraph(); shade(w,'E8EEF7'); w.add_run('Rationale: ').bold=True; w.add_run(why)
d.save('NUR774_Ear_Pictures_Explained.docx')
# ---------- Phone PDF
x=html.escape
P=[f'<section class="card cover"><h1>{TITLE}</h1><p>{x(SUB)}</p><p><b>How to use:</b> look at each picture and name it first, then swipe for the answer. Questions follow each picture, with the answer on the next page.</p><h3>Quick ID cheat sheet</h3><ul class="cheat">'+''.join(f"<li><b>Slide {e['s']}:</b> {x(e['title'])}</li>" for e in E)+'</ul></section>']
qn_=0
for e in E:
    P.append(f'<section class="card"><div class="num">Slide {e["s"]}: what is this?</div><img src="ear_img/slide{e["s"]}.jpg"><p class="guess">Name it before you swipe →</p></section>')
    P.append(f'<section class="card info"><div class="num">Slide {e["s"]}</div><h2>{x(e["title"])}</h2><img class="small" src="ear_img/slide{e["s"]}.jpg"><p class="cr">{x(e["credit"])}</p><p><b>What it is:</b> {x(e["what"])}</p><p><b>What you see:</b></p><ul>'+''.join(f'<li>{x(s)}</li>' for s in e['see'])+f'</ul></section><section class="card info"><div class="num">Slide {e["s"]}: remember it</div><h2>{x(e["title"])}</h2><div class="mn">🧠 {x(e["mn"])}</div><p><b>Exam tips:</b></p><ul>'+''.join(f'<li>{x(s)}</li>' for s in e['tip'])+'</ul></section>')
    for stem,opts,ans,why,sata in e['qs']:
        qn_+=1
        tag='<span class="sata">SELECT ALL THAT APPLY</span>' if sata else '<span class="single">SINGLE ANSWER</span>'
        P.append(f'<section class="card"><div class="num">Question {qn_} (slide {e["s"]})</div>{tag}<img class="thumb" src="ear_img/slide{e["s"]}.jpg"><p class="stem">{x(stem.replace(" (Select all that apply.)",""))}</p><ol type="A">'+''.join(f'<li>{x(o)}</li>' for o in opts)+'</ol><div class="next">Swipe for answer →</div></section>')
        P.append(f'<section class="card ans"><div class="num">Answer {qn_}</div><p class="key">{x(ans)}</p><ul>'+''.join(f'<li><b>{x(opts[L.index(a.strip())])}</b></li>' for a in ans.split(','))+f'</ul><p><b>Why:</b> {x(why)}</p></section>')
css="""@page{size:4in 7.2in;margin:0}body{margin:0;font-family:Arial,sans-serif}
.card{box-sizing:border-box;width:4in;min-height:7.2in;padding:.25in .25in;break-before:page;font-size:11.5pt;line-height:1.33}
h1{color:#1f3864;font-size:19pt;margin:0 0 8px}h2{color:#1f3864;font-size:15pt;margin:4px 0}h3{color:#2f5496;margin:10px 0 4px}
.cheat{font-size:9.5pt;padding-left:16px}.cheat li{margin:2px 0}
.num{color:#666;font-size:9.5pt;margin-bottom:6px}img{width:100%;border-radius:8px}img.small{width:52%;display:block;margin:4px auto}img.thumb{width:45%;display:block;margin:8px 0}
.guess{text-align:center;color:#2f5496;font-weight:bold;margin-top:14px}.cr{font-size:7.5pt;color:#888;text-align:center;margin:0}
ul{padding-left:18px;margin:2px 0 6px}li{margin:3px 0}.mn{background:#fff4cc;border-left:4px solid #e0a800;padding:7px 9px;margin:8px 0;font-weight:bold}
.sata,.single{display:inline-block;font-size:8.5pt;font-weight:bold;padding:2px 7px;border-radius:9px;color:#fff;background:#c0392b}.single{background:#2f5496}
.stem{font-weight:bold;margin:8px 0}ol{padding-left:22px;margin:0}ol li{margin:6px 0}.next{color:#999;font-size:9pt;margin-top:12px;text-align:right}
.ans{background:#eef3fb;-webkit-print-color-adjust:exact}.key{font-size:24pt;font-weight:bold;color:#1f7a3a;margin:4px 0}.info{font-size:11pt}"""
open('ear.html','w').write(f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{"".join(P)}</body></html>')
