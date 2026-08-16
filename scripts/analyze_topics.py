#!/usr/bin/env python3
import json

with open('/home/z/my-project/vol/scripts/existing_150_questions.json','r',encoding='utf-8') as f:
    qs = json.load(f)

for i,q in enumerate(qs,1):
    print(f'{i}. [{q["ans"]}] {q["question"][:100]}')
