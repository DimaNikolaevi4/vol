import re, json

with open('/home/z/my-project/vol/scripts/105_current_text.txt', 'r', encoding='utf-8') as f:
    text = f.read()

lines = text.split('\n')
questions = []
i = 0
while i < len(lines):
    line = lines[i].strip()
    # Match question start: "1. ", "2. ", etc. at beginning of line
    m = re.match(r'^(\d+)\.\s+(.+)', line)
    if m:
        q_num = int(m.group(1))
        q_text = m.group(2)
        # Collect next lines until we find options or answer
        options = {}
        j = i + 1
        while j < len(lines):
            ln = lines[j].strip()
            # Check for option line: A) ... or B) ... etc
            opt_m = re.match(r'^([ABCD])\)\s+(.+?)[.;]?$', ln)
            if opt_m:
                options[opt_m.group(1)] = opt_m.group(2)
                j += 1
                continue
            # Check for answer line
            ans_m = re.match(r'^Правильный ответ под номером:\s*([ABCD])', ln)
            if ans_m and len(options) == 4:
                ans = ans_m.group(1)
                questions.append({
                    'num': q_num,
                    'question': q_text,
                    'A': options.get('A',''),
                    'B': options.get('B',''),
                    'C': options.get('C',''),
                    'D': options.get('D',''),
                    'ans': ans
                })
                i = j
                break
            # Check if it's a continuation of the question text (no option letter prefix)
            if not re.match(r'^[ABCD]\)', ln) and 'Правильный ответ' not in ln and ln:
                # Could be continuation of question or a malformed option
                # If it starts with whitespace and has no pattern, it might be question continuation
                q_text += ' ' + ln
                j += 1
                continue
            j += 1
        else:
            i += 1
        continue
    i += 1

print(f'Parsed {len(questions)} questions')
if len(questions) != 150:
    # Find missing question numbers
    found_nums = {q['num'] for q in questions}
    # Questions are numbered 1-10 per ticket, so nums repeat
    # Instead check total count
    print(f'Expected 150, got {len(questions)}')
    # Show ticket distribution
    from collections import Counter
    # Count questions per ticket (approximate by position)
    print('Last 5 parsed questions:')
    for q in questions[-5:]:
        print(f"  #{len(questions)-questions.index(q)}: {q['question'][:60]}...")
else:
    print('All 150 questions parsed successfully!')

from collections import Counter
ans_dist = Counter(q['ans'] for q in questions)
print(f'Answer distribution: {dict(sorted(ans_dist.items()))}')

# Save
with open('/home/z/my-project/vol/scripts/existing_150_questions.json', 'w', encoding='utf-8') as f:
    json.dump(questions, f, ensure_ascii=False, indent=2)

print('Saved: existing_150_questions.json')
