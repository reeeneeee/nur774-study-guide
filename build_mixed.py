import random, html, collections, re
exec(open('questions.py').read()); exec(open('sata.py').read())
exec(re.search(r'CUT = \[.*?\]\n', open('build_exam.py').read(), re.S).group(0))
CUT += ["Which group of findings most strongly suggests a VIRAL","Which instruction is correct for a child using an inhaled","During an asthma exacerbation","A 4-month-old has bronchiolitis","A 15-month-old has nonsevere AOM in BOTH","A 2-year-old has AOM along with purulent","A child has amber fluid","Ear irrigation is CONTRAINDICATED","Which finding distinguishes orbital","Which congenital heart defect is CYANOTIC","Which is NOT a component of Tetralogy","Which finding is an early sign of heart failure","Which murmur characteristic is always","A healthy 5-year-old has a soft","A 3-year-old has had fever for 6 days","A 12-month-old has a mild cold","A healthy 12-month-old lives with","A 6-week-old has episodes of apnea","A 2-year-old has a barking cough","A 15-year-old has severe throat pain","A child with chickenpox may return","Which vaccines are routinely given at the 2-month","A 5-year-old is entering kindergarten","Why is aspirin avoided","Which live vaccines are first given","Coughing fits, an inspiratory whoop","A child is newly diagnosed with mild persistent","A child with croup and stridor at rest","A 3-year-old has nonsevere AOM in one ear","A toddler wakes with thick yellow"]
for c in CUT: assert sum(q[1].startswith(c) for q in Q)==1, c
Q=[q for q in Q if not any(q[1].startswith(c) for c in CUT)]
assert len(Q)==70 and len(S)==30, (len(Q),len(S))
rng=random.Random(7741)
items=[]
letters=list("ABCD"*18)[:70]; rng.shuffle(letters)
for (t,stem,ok,bad,why),L in zip(Q,letters):
    b=bad[:]; rng.shuffle(b); k="ABCD".index(L); opts=b[:k]+[ok]+b[k:]
    items.append(dict(t=t,stem=stem,opts=opts,ans=[L],why=why,sata=False))
for t,stem,ok,bad,why in S:
    opts=ok+bad; rng.shuffle(opts)
    items.append(dict(t=t,stem=stem+" (Select all that apply.)",opts=opts,ans=[ "ABCDEFG"[i] for i,o in enumerate(opts) if o in ok],why=why,sata=True))
rng.shuffle(items)
print(collections.Counter(i['ans'][0] for i in items if not i['sata']))
LET="ABCDEFG"
def anstext(i): return ', '.join(i['ans'])+': '+'; '.join(i['opts'][LET.index(a)] for a in i['ans'])
TITLE='NUR 774 Exam 3: 100-Question Practice Exam'
SUB='70 single-answer + 30 Select All That Apply (SATA). Built from your Integrated Study Guide. SATA tip: judge each option as its own true/false question.'
# ---- Word
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
d=Document(); st=d.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(11)
d.add_heading(TITLE,0); d.add_paragraph(SUB).runs[0].italic=True
def shade(p,color='E8EEF7'):
    pPr=p._p.get_or_add_pPr(); s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),color); pPr.append(s)
for n,i in enumerate(items,1):
    p=d.add_paragraph(); p.paragraph_format.keep_with_next=True; p.paragraph_format.space_after=Pt(2)
    r=p.add_run(f'{n}. '); r.bold=True
    if i['sata']:
        r=p.add_run('[SATA] '); r.bold=True; r.font.color.rgb=RGBColor(0xC0,0x39,0x2B)
    p.add_run(i['stem']).bold=True
    for L,o in zip(LET,i['opts']):
        op=d.add_paragraph(f'   {L}. {o}'); op.paragraph_format.space_after=Pt(0); op.paragraph_format.keep_with_next=True
    a=d.add_paragraph(); shade(a); a.paragraph_format.space_before=Pt(4); a.paragraph_format.keep_with_next=True; a.paragraph_format.space_after=Pt(0)
    a.add_run('Answer: ').bold=True; a.add_run(anstext(i))
    w=d.add_paragraph(); shade(w); w.paragraph_format.space_after=Pt(12)
    w.add_run('Rationale: ').bold=True; w.add_run(i['why']); w.add_run(f'   [{i["t"]}]').italic=True
d.save('NUR774_Exam3_Practice_SATA.docx')
# ---- Phone PDF (flashcards: question page, then answer page)
e=html.escape
P=[f'<section class="card cover"><h1>{TITLE}</h1><p>{SUB}</p><p><b>How to use:</b> read the question, pick your answer, then swipe to the next page for the answer and rationale.</p></section>']
for n,i in enumerate(items,1):
    tag='<span class="sata">SELECT ALL THAT APPLY</span>' if i['sata'] else '<span class="single">SINGLE ANSWER</span>'
    stem=i['stem'].replace(' (Select all that apply.)','')
    P.append(f'<section class="card"><div class="num">Question {n} of 100</div>{tag}<p class="stem">{e(stem)}</p><ol type="A">'+''.join(f'<li>{e(o)}</li>' for o in i['opts'])+'</ol><div class="next">Swipe for answer →</div></section>')
    P.append(f'<section class="card ans"><div class="num">Answer {n}</div><p class="key">{e(", ".join(i["ans"]))}</p><ul>'+''.join(f'<li>{e(i["opts"][LET.index(a)])}</li>' for a in i['ans'])+f'</ul><p class="why"><b>Why:</b> {e(i["why"])}</p><p class="topic">{e(i["t"])}</p></section>')
css="""@page{size:4in 7.2in;margin:0}body{margin:0;font-family:Arial,sans-serif}
.card{box-sizing:border-box;width:4in;height:7.2in;padding:.3in .28in;break-after:page;overflow:hidden;font-size:12.5pt;line-height:1.35}
.cover h1{color:#1f3864;font-size:20pt}.num{color:#666;font-size:10pt;margin-bottom:6px}
.sata,.single{display:inline-block;font-size:9pt;font-weight:bold;padding:2px 7px;border-radius:9px;color:#fff;background:#c0392b}.single{background:#2f5496}
.stem{font-weight:bold;margin:10px 0}ol{padding-left:22px;margin:0}ol li{margin:7px 0}
.next{color:#999;font-size:9.5pt;margin-top:14px;text-align:right}
.ans{background:#eef3fb}.key{font-size:26pt;font-weight:bold;color:#1f7a3a;margin:4px 0}.ans ul{padding-left:18px}.ans li{margin:4px 0;font-weight:bold}
.why{margin-top:12px}.topic{color:#666;font-size:10pt;font-style:italic}"""
open('phone.html','w').write(f'<html><head><meta charset="utf-8"><style>{css}</style></head><body>{"".join(P)}</body></html>')
# overflow check helper list
open('_counts.txt','w').write(str(len(P)))
