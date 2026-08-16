#!/usr/bin/env python3
"""Extract unique questions from ПП 105."""
import os, re, json
from collections import Counter
from docx import Document

FOLDER = "/home/z/my-project/vol/Оператор автоматических и полуавтоматических станков и линий станков/профессиональной переподготовки 15474 Оператор станков"
path = os.path.join(FOLDER, [f for f in os.listdir(FOLDER) if f.startswith('105.')][0])

doc = Document(path)
text = '\n'.join(p.text for p in doc.paragraphs)

lines = text.split('\n')
questions = []
i = 0
while i < len(lines):
    line = lines[i].strip()
    m = re.match(r'^(\d+)\.\s+(.+)', line)
    if m:
        q_text = m.group(2)
        options = {}
        j = i + 1
        while j < len(lines):
            ln = lines[j].strip()
            opt_m = re.match(r'^([A-Da-dа-г])[).)]\s+(.+?)[.;]?$', ln)
            if opt_m:
                ltr = opt_m.group(1).upper()
                if ltr in 'ABCD':
                    options[ltr] = opt_m.group(2)
                j += 1
                continue
            ans_m = re.match(r'^Правильный ответ под номером:\s*([ABCDabcdа-г])', ln)
            if ans_m and len(options) == 4:
                ans = ans_m.group(1).upper()
                if ans in 'ABCD':
                    questions.append({
                        'question': q_text,
                        'A': options.get('A',''), 'B': options.get('B',''),
                        'C': options.get('C',''), 'D': options.get('D',''),
                        'ans': ans
                    })
                i = j
                break
            if ln and not re.match(r'^[A-Da-dа-г][).)]', ln) and 'Правильный ответ' not in ln:
                q_text += ' ' + ln
            j += 1
        else:
            i += 1
        continue
    i += 1

print(f'Всего распарсено вопросов: {len(questions)}')

# Deduplicate
seen = set()
unique = []
for q in questions:
    if q['question'] not in seen:
        seen.add(q['question'])
        unique.append(q)

print(f'Уникальных: {len(unique)}')
print(f'Дубликатов удалено: {len(questions) - len(unique)}')

d = Counter(q['ans'] for q in unique)
print(f'D-распределение уникальных: {dict(sorted(d.items()))}')

with open('/home/z/my-project/vol/scripts/pp_existing_unique.json', 'w', encoding='utf-8') as f:
    json.dump(unique, f, ensure_ascii=False, indent=2)
print('Сохранено: pp_existing_unique.json')
