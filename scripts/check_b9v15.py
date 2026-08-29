#!/usr/bin/env python3
"""Check actual text in Б9в15 to debug 'Ш pressova' fix."""

import re
from docx import Document

BASE = '/home/z/my-project/vol/Оператор автоматических и полуавтоматических станков и линий станков'
FOLDER = 'профессиональной подготовки 15474 Оператор станков'
FILEPATH = BASE + '/' + FOLDER + '/105. ПО_П_ФОС_Оператор станков_2-3_разр.docx'

Q_RE = re.compile(r'^(\d+)\s*[.)]\s*(.*)')
OPT_RE = re.compile(r'^([ABCD])\)\s*(.*)')
ANS_RE = re.compile(r'Правильный ответ под номером:\s*([ABCD])')
TICKET_RE = re.compile(r'Билет\s+№?\s*(\d+)')

doc = Document(FILEPATH)
cur_ticket = None
cur_q = None

for p in doc.paragraphs:
    txt_s = p.text.strip()
    tm = TICKET_RE.match(txt_s)
    if tm:
        cur_ticket = int(tm.group(1))
        cur_q = None
        continue
    if cur_ticket != 9:
        continue
    m = Q_RE.match(txt_s)
    if m:
        cur_q = int(m.group(1))
        if cur_q == 15:
            print('Q text:', repr(m.group(2)))
        continue
    m2 = OPT_RE.match(txt_s)
    if m2 and cur_q == 15:
        print(f'Opt {m2.group(1)}:', repr(m2.group(2)))
        continue
