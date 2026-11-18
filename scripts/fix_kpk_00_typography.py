# -*- coding: utf-8 -*-
"""Приведение типографики к требованию: Times New Roman, текст 14 pt, таблицы 12 pt.

  - docDefaults: Times New Roman, 14 pt
  - все абзацы вне таблиц (включая заголовки): 14 pt
  - все ячейки таблиц: 12 pt
  - в матрице ответов шапка «Вопрос» -> «№», иначе 20 колонок не влезают в полосу

Запуск: python -X utf8 fix_typography.py <in.docx> <out.docx>
"""
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

FONT = 'Times New Roman'
TEXT_PT = 14
TABLE_PT = 12


def set_run(run, pt):
    run.font.name = FONT
    run.font.size = Pt(pt)
    rPr = run._r.get_or_add_rPr()
    fonts = rPr.find(qn('w:rFonts'))
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        rPr.insert(0, fonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        fonts.set(qn(attr), FONT)
    for tag in ('w:sz', 'w:szCs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn('w:val'), str(int(pt * 2)))


def main(src, dst):
    doc = Document(src)

    # 1. docDefaults
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpr_def = dd.find(qn('w:rPrDefault'))
        if rpr_def is not None:
            rPr = rpr_def.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                rpr_def.append(rPr)
            fonts = rPr.find(qn('w:rFonts'))
            if fonts is None:
                fonts = OxmlElement('w:rFonts')
                rPr.insert(0, fonts)
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                fonts.set(qn(attr), FONT)
            sz = rPr.find(qn('w:sz'))
            if sz is None:
                sz = OxmlElement('w:sz')
                rPr.append(sz)
            sz.set(qn('w:val'), str(TEXT_PT * 2))

    # 2. текст вне таблиц
    n_text = 0
    for par in doc.paragraphs:
        for run in par.runs:
            if run.text.strip():
                set_run(run, TEXT_PT)
                n_text += 1

    # 3. таблицы — обход по XML, включая вложенные таблицы и объединённые ячейки
    n_tbl = 0
    body = doc.element.body
    for tbl in body.iter(qn('w:tbl')):
        for r in tbl.iter(qn('w:r')):
            texts = [t.text for t in r.findall(qn('w:t')) if t.text and t.text.strip()]
            if not texts:
                continue
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)
            fonts = rPr.find(qn('w:rFonts'))
            if fonts is None:
                fonts = OxmlElement('w:rFonts')
                rPr.insert(0, fonts)
            for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                fonts.set(qn(attr), FONT)
            for tag in ('w:sz', 'w:szCs'):
                el = rPr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    rPr.append(el)
                el.set(qn('w:val'), str(TABLE_PT * 2))
            n_tbl += 1

    # 4. матрица ответов: шапка «Вопрос» -> «№»
    changed_head = 0
    for t in doc.tables:
        if len(t.rows) == 17 and len(t.columns) == 21:
            for cell in t.rows[0].cells:
                for par in cell.paragraphs:
                    for run in par.runs:
                        if run.text.strip() == 'Вопрос':
                            run.text = '№'
                            changed_head += 1
    print('runs: текст %d, таблицы %d, шапок матрицы %d' % (n_text, n_tbl, changed_head))
    doc.save(dst)
    print('saved', dst)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
