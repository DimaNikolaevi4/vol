# -*- coding: utf-8 -*-
"""Проверка мега-документа 00: секции, вписывание таблиц в полосу, уникальность нумерации.

Проверяет:
  - ориентацию и полосу набора каждой секции;
  - ширину каждой таблицы против полосы СВОЕЙ секции (плюс нулевой отступ);
  - дубли номеров подразделов (записи оглавления исключаются);
  - наличие поля TOC, завершающей sectPr.

Запуск:
    python -X utf8 verify_kpk_00_program.py --src out.docx [--report report.txt]
"""
import argparse
import io
import re

from docx import Document
from docx.oxml.ns import qn

TW = 20.0


def main():
    ap = argparse.ArgumentParser(description='Проверка вёрстки мега-документа 00')
    ap.add_argument('--src', required=True, help='проверяемый docx')
    ap.add_argument('--report', default=None, help='файл отчёта (по умолчанию <src>.verify.txt)')
    args = ap.parse_args()
    path = args.src
    out_path = args.report or (path + '.verify.txt')

    doc = Document(path)
    out = []
    add = out.append

    add('=== секции (%d)' % len(doc.sections))
    widths = []
    for i, s in enumerate(doc.sections):
        w_pt = s.page_width / 12700.0
        h_pt = s.page_height / 12700.0
        text_w = w_pt - s.left_margin / 12700.0 - s.right_margin / 12700.0
        widths.append(text_w)
        add('S%d %s %.0fx%.0f pt, полоса %.1f pt' % (i, s.orientation, w_pt, h_pt, text_w))

    sec_of = []
    sec = 0
    body = doc.element.body
    for el in body.iterchildren():
        tag = el.tag.split('}')[1]
        if tag == 'p':
            pPr = el.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                sec += 1
        elif tag == 'tbl':
            sec_of.append(min(sec, len(widths) - 1))

    add('')
    add('=== таблицы')
    over = 0
    for ti, t in enumerate(doc.tables):
        grid = t._tbl.find(qn('w:tblGrid'))
        w = sum(int(gc.get(qn('w:w'))) for gc in grid) / TW
        ind = t._tbl.tblPr.find(qn('w:tblInd'))
        ind_w = ind.get(qn('w:w')) if ind is not None else None
        ind_v = int(ind_w) / TW if ind_w else 0.0
        s = sec_of[ti]
        flag = ''
        if w > widths[s] + 0.6 or ind_v != 0:
            flag = ' <<< ВЫХОД ЗА ПОЛОСУ'
            over += 1
        add('T%-2d S%d w=%6.1f ind=%5.1f полоса=%6.1f%s' % (ti, s, w, ind_v, widths[s], flag))
    add('таблиц с выходом за полосу: %d' % over)

    add('')
    add('=== нумерация подразделов (без записей оглавления)')
    num_re = re.compile(r'^(\d+(?:\.\d+){1,3})\.\s')
    seen = {}
    for p in doc.paragraphs:
        if p.style.name.lower().startswith('toc'):
            continue
        t = ' '.join(p.text.split())
        m = num_re.match(t)
        if m:
            seen.setdefault(m.group(1), []).append(t[:60])
    dups = {k: v for k, v in seen.items() if len(v) > 1}
    add('уникальных номеров: %d, дублей: %d' % (len(seen), len(dups)))
    for k, v in sorted(dups.items()):
        add('  ДУБЛЬ %s x%d: %r' % (k, len(v), v[0]))

    add('')
    add('=== служебное')
    txt = '\n'.join(p.text for p in doc.paragraphs)
    add('поле TOC: %s' % ('есть' if 'TOC \\o' in body.xml else 'НЕТ'))
    add('завершающая sectPr: %s' % ('есть' if body.find(qn('w:sectPr')) is not None else 'НЕТ'))
    add('абзацев: %d, таблиц: %d' % (len(doc.paragraphs), len(doc.tables)))
    add('итог: %s' % ('OK' if over == 0 and not dups else 'ЕСТЬ ЗАМЕЧАНИЯ'))

    io.open(out_path, 'w', encoding='utf-8').write('\n'.join(out))
    print('over=%d dups=%d -> %s' % (over, len(dups), out_path))


if __name__ == '__main__':
    main()
