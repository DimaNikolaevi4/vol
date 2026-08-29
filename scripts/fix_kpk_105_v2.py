#!/usr/bin/env python3
"""КПК 105 v2: 4 fixes from external review
1. Б6в7: аттестация->СОУТ, ответ C->B
2. Б12в9: Сорбит закалки->Сорбит отпуска (опция D текст)
3. Б12в12: ПОТ ЭЭ ответ C->D, опция D текст fix
4. Б15в5: замена вопроса целиком
+ D-balance compensation
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

# === PASS 1: Collect all paragraph indices per ticket/question ===
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
            'q_idx': cur_q_start,
            'opts': cur_opts[:],
            'ans_idx': cur_ans_idx
        })

for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    
    m_t = TICKET_RE.match(text)
    if m_t:
        save_q()
        cur_ticket = int(m_t.group(1))
        cur_qnum = None
        cur_opts = []
        cur_ans_idx = None
        continue
    
    m_q = Q_RE.match(text)
    if m_q and cur_ticket is not None:
        save_q()
        cur_qnum = int(m_q.group(1))
        cur_q_start = i
        cur_opts = []
        cur_ans_idx = None
        continue
    
    m_opt = OPT_RE.match(text)
    if m_opt and cur_qnum is not None:
        cur_opts.append({'idx': i, 'letter': m_opt.group(1), 'text': m_opt.group(2)})
        continue
    
    m_ans = ANS_RE.match(text)
    if m_ans and cur_qnum is not None:
        cur_ans_idx = i

save_q()

print(f'Собрано копий: {sum(len(v2) for v1 in copies.values() for v2 in v1.values())}')

# === FIX 1: Б6в7 - аттестация->СОУТ, answer C->B ===
print('\n=== FIX 1: Б6в7 ===')
for cp in copies.get(6, {}).get(7, []):
    p = doc.paragraphs[cp['q_idx']]
    for run in p.runs:
        if 'аттестации рабочего места' in run.text:
            run.text = run.text.replace('аттестации рабочего места', 'проведения специальной оценки условий труда (СОУТ)')
            print(f'  p{cp["q_idx"]}: вопрос текст')
    if cp['ans_idx'] is not None:
        p_ans = doc.paragraphs[cp['ans_idx']]
        for run in p_ans.runs:
            if 'Правильный ответ под номером: C' in run.text:
                run.text = run.text.replace('Правильный ответ под номером: C', 'Правильный ответ под номером: B')
                print(f'  p{cp["ans_idx"]}: ответ C->B')

# === FIX 2: Б12в9 - Сорбит закалки -> Сорбит отпуска ===
print('\n=== FIX 2: Б12в9 ===')
for cp in copies.get(12, {}).get(9, []):
    for opt in cp['opts']:
        if opt['letter'] == 'D':
            p = doc.paragraphs[opt['idx']]
            for run in p.runs:
                if 'Сорбит (троостит) закалки' in run.text:
                    run.text = run.text.replace('Сорбит (троостит) закалки', 'Сорбит отпуска')
                    print(f'  p{opt["idx"]}: D текст')
                elif 'закалки' in run.text and 'Сорбит' in run.text:
                    run.text = run.text.replace('закалки', 'отпуска')
                    print(f'  p{opt["idx"]}: D текст (alt)')

# === FIX 3: Б12в12 - ПОТ ЭЭ answer C->D, fix D text ===
print('\n=== FIX 3: Б12в12 ===')
for cp in copies.get(12, {}).get(12, []):
    for opt in cp['opts']:
        if opt['letter'] == 'D':
            p = doc.paragraphs[opt['idx']]
            for run in p.runs:
                if 'Не нормируется, если руки сухие' in run.text:
                    run.text = run.text.replace('Не нормируется, если руки сухие', 'Не нормируется (без прикосновение)')
                    print(f'  p{opt["idx"]}: D текст')
    if cp['ans_idx'] is not None:
        p_ans = doc.paragraphs[cp['ans_idx']]
        for run in p_ans.runs:
            if 'Правильный ответ под номером: C' in run.text:
                run.text = run.text.replace('Правильный ответ под номером: C', 'Правильный ответ под номером: D')
                print(f'  p{cp["ans_idx"]}: ответ C->D')

# === FIX 4: Б15в5 - replace entire question ===
print('\n=== FIX 4: Б15в5 ===')
NEW_Q = 'Согласно ПОТ ЭЭ, каково минимальное расстояние от токоведущих частей до человека без применения электрозащитных средств в электроустановках до 1000 В?'
NEW_OPTS = {'A': '0,2 м', 'B': '0,6 м', 'C': '0,8 м', 'D': 'Не нормируется (без прикосновение)'}
NEW_ANS = 'D'

for cp in copies.get(15, {}).get(5, []):
    p = doc.paragraphs[cp['q_idx']]
    for run in p.runs:
        run.text = ''
    p.runs[0].text = f'5. {NEW_Q}'
    print(f'  p{cp["q_idx"]}: вопрос заменён')
    
    for opt in cp['opts']:
        p_opt = doc.paragraphs[opt['idx']]
        new_text = f"{opt['letter']}){NEW_OPTS[opt['letter']]}"
        for run in p_opt.runs:
            run.text = ''
        p_opt.runs[0].text = new_text
        print(f'  p{opt["idx"]}: {opt["letter"]} опция')
    
    if cp['ans_idx'] is not None:
        p_ans = doc.paragraphs[cp['ans_idx']]
        for run in p_ans.runs:
            if 'Правильный ответ под номером:' in run.text:
                run.text = f'Правильный ответ под номером: {NEW_ANS}'
                print(f'  p{cp["ans_idx"]}: ответ -> {NEW_ANS}')

doc.save(FPATH)
print(f'\nСохранено: {FNAME}')

# === Verify ===
print('\n=== БАЛАНС ===')
doc2 = docx.Document(FPATH)
qs = []
cur_q2 = None
cur_t2 = None

for p in doc2.paragraphs:
    text = p.text.strip()
    if not text:
        continue
    m_a = ANS_RE.match(text)
    if m_a:
        if cur_q2: cur_q2['answer'] = m_a.group(1)
        continue
    m_o = OPT_RE.match(text)
    if m_o and cur_q2: cur_q2['options'][m_o.group(1)] = m_o.group(2); continue
    m_q = Q_RE.match(text)
    if m_q:
        if cur_q2: qs.append(cur_q2)
        cur_q2 = {'ticket': cur_t2, 'num': int(m_q.group(1)), 'text': m_q.group(2), 'options': {}, 'answer': None}
        continue
    m_t = TICKET_RE.match(text)
    if m_t:
        if cur_q2: qs.append(cur_q2); cur_q2 = None
        cur_t2 = int(m_t.group(1))
if cur_q2: qs.append(cur_q2)

from collections import Counter
dt = Counter(q.get('answer') for q in qs if q.get('answer'))
print(f'Общий: {dict(dt)}')

for t in sorted(set(q['ticket'] for q in qs if q['ticket'] is not None)):
    tq = [q for q in qs if q['ticket'] == t and q.get('answer')]
    td = Counter(q['answer'] for q in tq)
    d = td.get('D', 0)
    fl = '' if 4 <= d <= 6 else ' FAIL'
    print(f'  Б{t}: {dict(td)} D={d}{fl}')