#!/usr/bin/env python3
"""КПК 105 v2b: D-balance compensation
After v2 fixes: A=74, B=76, C=73, D=77
Need: A+1, B-1, C+2, D-2

Swap plan:
- B2v3: A<->B swap (B->A: A+1, B-1)
- B12v14: C<->D swap (D->C: C+1, D-1)
- B15v8: C<->D swap (D->C: C+1, D-1)
Result: 75/75/75/75
"""
import re, docx

FOLDER = 'Оператор автоматических и полуавтоматических станков и линий станков/повышения квалификации 15474 Оператор станков'
FNAME = '105. ПО_КПК_ФОС_Оператор станков_4_разр.docx'
FPATH = f'{FOLDER}/{FNAME}'

TICKET_RE = re.compile(r'^Билет\s+№?\s*(\d+)')
Q_RE = re.compile(r'^(\d+)\s*[.)]\s*(.*)')
OPT_RE = re.compile(r'^([ABCD])\)\s*(.*)')
ANS_RE = re.compile(r'^Правильный ответ под номером:\s*([ABCD])')

doc = docx.Document(FPATH)

# Collect all copies
copies = {}
cur_ticket = None
cur_qnum = None
cur_q_start = None
cur_opts = []
cur_ans_idx = None

def save_q():
    global cur_ticket, cur_qnum, cur_q_start, cur_opts, cur_ans_idx
    if cur_ticket is not None and cur_qnum is not None:
        copies.setdefault(cur_ticket, {}).setdefault(cur_qnum, []).append({
            'q_idx': cur_q_start, 'opts': cur_opts[:], 'ans_idx': cur_ans_idx
        })

for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if not text: continue
    m_t = TICKET_RE.match(text)
    if m_t:
        save_q()
        cur_ticket = int(m_t.group(1)); cur_qnum = None; cur_opts = []; cur_ans_idx = None
        continue
    m_q = Q_RE.match(text)
    if m_q and cur_ticket is not None:
        save_q()
        cur_qnum = int(m_q.group(1)); cur_q_start = i; cur_opts = []; cur_ans_idx = None
        continue
    m_o = OPT_RE.match(text)
    if m_o and cur_qnum is not None:
        cur_opts.append({'idx': i, 'letter': m_o.group(1), 'text': m_o.group(2)})
        continue
    m_a = ANS_RE.match(text)
    if m_a and cur_qnum is not None:
        cur_ans_idx = i
save_q()

print(f'Копий: {sum(len(v2) for v1 in copies.values() for v2 in v1.values())}')

# === SWAP 1: B2v3 A<->B (answer B->A) ===
print('\n=== SWAP B2v3 A<->B ===')
for cp in copies.get(2, {}).get(3, []):
    a_idx = None; b_idx = None
    for opt in cp['opts']:
        if opt['letter'] == 'A': a_idx = opt['idx']
        if opt['letter'] == 'B': b_idx = opt['idx']
    if a_idx is not None and b_idx is not None:
        a_text = doc.paragraphs[a_idx].runs[0].text[2:]  # strip 'A)'
        b_text = doc.paragraphs[b_idx].runs[0].text[2:]  # strip 'B)'
        doc.paragraphs[a_idx].runs[0].text = f'A){b_text}'
        doc.paragraphs[b_idx].runs[0].text = f'B){a_text}'
        print(f'  A={b_text[:40]} <-> B={a_text[:40]}')
    if cp['ans_idx'] is not None:
        for run in doc.paragraphs[cp['ans_idx']].runs:
            if 'Правильный ответ под номером: B' in run.text:
                run.text = run.text.replace('Правильный ответ под номером: B', 'Правильный ответ под номером: A')
                print(f'  ответ B->A')

# === SWAP 2: B12v14 C<->D (answer D->C) ===
print('\n=== SWAP B12v14 C<->D ===')
for cp in copies.get(12, {}).get(14, []):
    c_idx = None; d_idx = None
    for opt in cp['opts']:
        if opt['letter'] == 'C': c_idx = opt['idx']
        if opt['letter'] == 'D': d_idx = opt['idx']
    if c_idx is not None and d_idx is not None:
        c_text = doc.paragraphs[c_idx].runs[0].text[2:]
        d_text = doc.paragraphs[d_idx].runs[0].text[2:]
        doc.paragraphs[c_idx].runs[0].text = f'C){d_text}'
        doc.paragraphs[d_idx].runs[0].text = f'D){c_text}'
        print(f'  C={d_text[:40]} <-> D={c_text[:40]}')
    if cp['ans_idx'] is not None:
        for run in doc.paragraphs[cp['ans_idx']].runs:
            if 'Правильный ответ под номером: D' in run.text:
                run.text = run.text.replace('Правильный ответ под номером: D', 'Правильный ответ под номером: C')
                print(f'  ответ D->C')

# === SWAP 3: B15v8 C<->D (answer D->C) ===
print('\n=== SWAP B15v8 C<->D ===')
for cp in copies.get(15, {}).get(8, []):
    c_idx = None; d_idx = None
    for opt in cp['opts']:
        if opt['letter'] == 'C': c_idx = opt['idx']
        if opt['letter'] == 'D': d_idx = opt['idx']
    if c_idx is not None and d_idx is not None:
        c_text = doc.paragraphs[c_idx].runs[0].text[2:]
        d_text = doc.paragraphs[d_idx].runs[0].text[2:]
        doc.paragraphs[c_idx].runs[0].text = f'C){d_text}'
        doc.paragraphs[d_idx].runs[0].text = f'D){c_text}'
        print(f'  C={d_text[:40]} <-> D={c_text[:40]}')
    if cp['ans_idx'] is not None:
        for run in doc.paragraphs[cp['ans_idx']].runs:
            if 'Правильный ответ под номером: D' in run.text:
                run.text = run.text.replace('Правильный ответ под номером: D', 'Правильный ответ под номером: C')
                print(f'  ответ D->C')

doc.save(FPATH)
print(f'\nСохранено: {FNAME}')

# Verify
print('\n=== БАЛАНС ===')
doc2 = docx.Document(FPATH)
qs2 = []; cq = None; ct = None
for p in doc2.paragraphs:
    t = p.text.strip()
    if not t: continue
    ma = ANS_RE.match(t)
    if ma:
        if cq: cq['a'] = ma.group(1)
        continue
    mo = OPT_RE.match(t)
    if mo and cq: cq['o'][mo.group(1)] = mo.group(2); continue
    mq = Q_RE.match(t)
    if mq:
        if cq: qs2.append(cq)
        cq = {'t': ct, 'n': int(mq.group(1)), 'o': {}, 'a': None}
        continue
    mt = TICKET_RE.match(t)
    if mt:
        if cq: qs2.append(cq); cq = None
        ct = int(mt.group(1))
if cq: qs2.append(cq)

from collections import Counter
dt = Counter(q['a'] for q in qs2 if q.get('a'))
print(f'Общий: {dict(dt)}')
for t in sorted(set(q['t'] for q in qs2 if q['t'] is not None)):
    tq = [q for q in qs2 if q['t'] == t and q.get('a')]
    td = Counter(q['a'] for q in tq)
    d = td.get('D', 0)
    fl = '' if 4 <= d <= 6 else ' FAIL'
    print(f'  B{t}: {dict(td)} D={d}{fl}')
