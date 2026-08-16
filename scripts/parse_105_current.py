#!/usr/bin/env python3
"""Parse current 105 (150 questions) and save as structured data for reuse."""
import re, json

with open('/home/z/my-project/vol/scripts/105_current_text.txt', 'r', encoding='utf-8') as f:
    text = f.read()

# Pattern: numbered question, 4 options (A/B/C/D), answer line
q_pattern = re.compile(
    r'(\d+)\. (.+?)\n'
    r'\s*A\) (.+?);\n'
    r'\s*B\) (.+?);\n'
    r'\s*C\) (.+?);\n'
    r'\s*D\) (.+?)\.(?:\n|$)'
    r'\s*Правильный ответ под номером: (\w)',
    re.MULTILINE
)

matches = q_pattern.findall(text)
print(f"Parsed {len(matches)} questions")

questions = []
for m in matches:
    q_num, q_text, a, b, c, d, ans = m
    questions.append({
        'num': int(q_num),
        'question': q_text.strip(),
        'A': a.strip(),
        'B': b.strip(),
        'C': c.strip(),
        'D': d.strip(),
        'ans': ans.strip()
    })

# Save as JSON for reuse
with open('/home/z/my-project/vol/scripts/existing_150_questions.json', 'w', encoding='utf-8') as f:
    json.dump(questions, f, ensure_ascii=False, indent=2)

# Print summary
from collections import Counter
ans_dist = Counter(q['ans'] for q in questions)
print(f"Answer distribution: {dict(sorted(ans_dist.items()))}")
print(f"Total: {len(questions)}")

# Print first 5 for verification
for q in questions[:3]:
    print(f"\nQ{q['num']}: {q['question'][:80]}...")
    print(f"  A) {q['A'][:50]}  B) {q['B'][:50]}")
    print(f"  C) {q['C'][:50]}  D) {q['D'][:50]}")
    print(f"  Ans: {q['ans']}")

# Save as text format for reference
with open('/home/z/my-project/vol/scripts/existing_150.txt', 'w', encoding='utf-8') as f:
    for i, q in enumerate(questions, 1):
        f.write(f"--- Q{i} (ans={q['ans']}) ---\n")
        f.write(f"{q['question']}\n")
        f.write(f"A) {q['A']}\n")
        f.write(f"B) {q['B']}\n")
        f.write(f"C) {q['C']}\n")
        f.write(f"D) {q['D']}\n\n")

print("\nSaved: existing_150_questions.json, existing_150.txt")
