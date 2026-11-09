import random, re, html, collections
exec(open('questions.py').read())
CUT = ["An elevated ASO titer","Which statement about foreign body aspiration","Croup is most commonly caused","A large, uncorrected left-to-right","Atrioventricular (AV) septal","A nurse is caring for an infant on a prostaglandin","Which statement about salmeterol","A child has abrupt fever, chills","The parent of an 8-month-old","The 1–3–6 hearing","A child who swims daily","A 16-year-old is starting the HPV","A child with fifth disease","What is the minimum age for routine ibuprofen","A teen who wears contact lenses","A 3-day-old with hypoplastic"]
for c in CUT: assert sum(q[1].startswith(c) for q in Q)==1, c
Q = [q for q in Q if not any(q[1].startswith(c) for c in CUT)]
assert len(Q)==100, len(Q)
rng = random.Random(774)
rng.shuffle(Q)
items=[]
letters=list("ABCD"*25); rng.shuffle(letters)
for (t,stem,ok,bad,why),L in zip(Q,letters):
    b=bad[:]; rng.shuffle(b); k="ABCD".index(L); opts=b[:k]+[ok]+b[k:]; items.append((t,stem,opts,L,why))
print(collections.Counter(i[3] for i in items))
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
d=Document(); s=d.styles['Normal']; s.font.name='Calibri'; s.font.size=Pt(11)
d.add_heading('NUR 774 Exam 3: 100-Question Practice Exam',0)
intro='Respiratory, ear, eye, cardiovascular, medications, and immunizations. Built from your Integrated Study Guide. Choose the ONE best answer. Answer key with explanations starts after question 100. Aim for 80%+ (80/100).'
d.add_paragraph(intro)
H=[f'<h1>NUR 774 Exam 3: 100-Question Practice Exam</h1><p><i>{intro}</i></p>']
for n,(t,stem,opts,a,why) in enumerate(items,1):
    p=d.add_paragraph(); r=p.add_run(f'{n}. {stem}'); r.bold=True; p.paragraph_format.space_after=Pt(2); p.paragraph_format.keep_with_next=True
    H.append(f'<div class="q"><p class="s"><b>{n}.</b> {html.escape(stem)}</p><ol type="A">')
    for L,o in zip("ABCD",opts):
        op=d.add_paragraph(f'   {L}. {o}'); op.paragraph_format.space_after=Pt(0); op.paragraph_format.keep_with_next=(L!='D')
        H.append(f'<li>{html.escape(o)}</li>')
    d.paragraphs[-1].paragraph_format.space_after=Pt(10)
    H.append('</ol></div>')
d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
d.add_heading('Answer Key and Explanations',1)
H.append('<h1 class="pb">Answer Key and Explanations</h1><p class="quick">'+' &nbsp; '.join(f'<b>{n}</b>-{i[3]}' for n,i in enumerate(items,1))+'</p>')
d.add_paragraph('Quick key: '+'  '.join(f'{n}-{i[3]}' for n,i in enumerate(items,1)))
for n,(t,stem,opts,a,why) in enumerate(items,1):
    p=d.add_paragraph(); r=p.add_run(f'{n}. {a} — {opts["ABCD".index(a)]}'); r.bold=True
    p.add_run(f'  [{t}]').italic=True
    w=d.add_paragraph(why); w.paragraph_format.left_indent=Pt(18); w.paragraph_format.space_after=Pt(8)
    H.append(f'<div class="a"><b>{n}. {a}: {html.escape(opts["ABCD".index(a)])}</b> <i class="t">[{t}]</i><br>{html.escape(why)}</div>')
d.add_heading('Score yourself by topic',2)
tips='Count your misses per topic tag in the answer key. Any topic with 2+ misses: reread that section of the mnemonics guide, then redo only those questions tomorrow.'
d.add_paragraph(tips); H.append(f'<h2>Score yourself by topic</h2><p>{tips}</p>')
d.save('NUR774_Exam3_Practice_100Q.docx')
css="""body{font-family:Calibri,Arial,sans-serif;font-size:10.5pt;line-height:1.3}h1{color:#1f3864;font-size:18pt}h2{color:#2f5496}
.q{break-inside:avoid;margin-bottom:9px}.s{margin:0 0 2px}ol{margin:0;padding-left:26px}li{margin:1px 0}
.pb{break-before:page}.a{break-inside:avoid;margin:0 0 7px}.t{color:#666;font-size:9pt}.quick{font-size:10pt;line-height:1.7}"""
open('exam.html','w').write(f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{"".join(H)}</body></html>')
