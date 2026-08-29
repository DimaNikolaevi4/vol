#!/usr/bin/env python3
"""КПК 105: ГОСТ 25347 -> 25346 в текстах вопросов (5 шт.)"""
import re, docx

FOLDER = 'Оператор автоматических и полуавтоматических станков и линий станков/повышения квалификации 15474 Оператор станков'
FNAME = '105. ПО_КПК_ФОС_Оператор станков_4_разр.docx'
FPATH = f'{FOLDER}/{FNAME}'

TICKET_RE = re.compile(r'^Билет\s+№?\s*(\d+)')
Q_RE = re.compile(r'^(\d+)\s*[.)]\s*(.*)')
OPT_RE = re.compile(r'^([ABCD])\)\s*(.*)')
ANS_RE = re.compile(r'^Правильный ответ под номером:\s*([ABCD])')

TARGETS = {(2, 11), (5, 16), (9, 4), (11, 10), (15, 9)}

# Collect paragraphs that need fix (all 3 copies per question)
fix_indices = []

doc = docx.Document(FPATH)
cur_ticket = None
cur_qnum = None
cur_q_start = None  # paragraph index where question text starts

for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    
    m_t = TICKET_RE.match(text)
    if m_t:
        cur_ticket = int(m_t.group(1))
        cur_qnum = None
        continue
    
    m_q = Q_RE.match(text)
    if m_q:
        cur_qnum = int(m_q.group(1))
        cur_q_start = i
        if cur_ticket is not None and (cur_ticket, cur_qnum) in TARGETS:
            fix_indices.append(i)
        continue

deduped = list(dict.fromkeys(fix_indices))  # remove duplicates preserving order
print(f'Параграфов к замене: {len(deduped)} (ожидаем 5*3=15 если 3 копии)')

for idx in deduped:
    p = doc.paragraphs[idx]
    old_text = p.text
    new_text = old_text.replace('ГОСТ 25347', 'ГОСТ 25346')
    if old_text != new_text:
        # Clear and rewrite
        for run in p.runs:
            if '25347' in run.text:
                run.text = run.text.replace('25347', '25346')
                print(f'  p{idx}: заменено в run')
            elif '25346' in run.text:
                pass  # already fixed
    else:
        print(f'  p{idx}: уже исправлено или нет совпадения')

doc.save(FPATH)
print(f'Сохранено: {FNAME}')
