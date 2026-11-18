# -*- coding: utf-8 -*-
"""Аудит шрифтов и кеглей документа: текст, таблицы, заголовки.

Показывает эффективные значения (прямое форматирование -> стиль -> docDefaults)
и распределение по категориям.

Запуск: python -X utf8 audit_fonts.py <docx> <отчёт.txt>
"""
import io
import sys
from collections import Counter

from docx import Document
from docx.oxml.ns import qn


def style_chain(par):
    """Цепочка стилей: сам стиль и его basedOn."""
    res = []
    st = par.style
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        res.append(st)
        st = st.base_style
    return res


def eff_value(par, run, tag):
    """Эффективное значение свойства: run -> стили абзаца -> None."""
    rPr = run._r.find(qn('w:rPr'))
    if rPr is not None:
        el = rPr.find(qn(tag))
        if el is not None and el.get(qn('w:val')):
            return el.get(qn('w:val'))
    for st in style_chain(par):
        rPr = st.element.find(qn('w:rPr'))
        if rPr is not None:
            el = rPr.find(qn(tag))
            if el is not None and el.get(qn('w:val')):
                return el.get(qn('w:val'))
    return None


def audit(path, out_path):
    doc = Document(path)

    # docDefaults
    dd = doc.styles.element.find(qn('w:docDefaults'))
    def_sz = None
    def_font = None
    if dd is not None:
        rpr = dd.find(qn('w:rPrDefault'))
        if rpr is not None:
            r = rpr.find(qn('w:rPr'))
            if r is not None:
                sz = r.find(qn('w:sz'))
                if sz is not None:
                    def_sz = int(sz.get(qn('w:val'))) / 2.0
                f = r.find(qn('w:rFonts'))
                if f is not None:
                    def_font = f.get(qn('w:ascii'))

    text_sizes = Counter()
    text_fonts = Counter()
    tbl_sizes = Counter()
    tbl_fonts = Counter()
    head_sizes = Counter()

    def collect(par, cat):
        for run in par.runs:
            if not run.text.strip():
                continue
            sz = eff_value(par, run, 'w:sz')
            sz = int(sz) / 2.0 if sz else def_sz
            fnt = None
            rPr = run._r.find(qn('w:rPr'))
            if rPr is not None:
                f = rPr.find(qn('w:rFonts'))
                if f is not None:
                    fnt = f.get(qn('w:ascii'))
            if fnt is None:
                for st in style_chain(par):
                    r = st.element.find(qn('w:rPr'))
                    if r is not None:
                        f = r.find(qn('w:rFonts'))
                        if f is not None and f.get(qn('w:ascii')):
                            fnt = f.get(qn('w:ascii'))
                            break
            fnt = fnt or def_font
            if cat == 'text':
                text_sizes[sz] += len(run.text)
                text_fonts[fnt] += len(run.text)
            elif cat == 'table':
                tbl_sizes[sz] += len(run.text)
                tbl_fonts[fnt] += len(run.text)
            else:
                head_sizes[sz] += len(run.text)

    for par in doc.paragraphs:
        cat = 'head' if par.style.name.startswith('Heading') else 'text'
        collect(par, cat)
    seen_cells = set()
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if id(cell._tc) in seen_cells:
                    continue
                seen_cells.add(id(cell._tc))
                for par in cell.paragraphs:
                    collect(par, 'table')

    out = []
    add = out.append
    add('docDefaults: шрифт=%s, кегль=%s' % (def_font, def_sz))
    add('')
    add('=== ОСНОВНОЙ ТЕКСТ (вне таблиц и заголовков)')
    add('  кегли: %s' % [(k, v) for k, v in text_sizes.most_common()])
    add('  шрифты: %s' % [(k, v) for k, v in text_fonts.most_common()])
    add('')
    add('=== ТАБЛИЦЫ')
    add('  кегли: %s' % [(k, v) for k, v in tbl_sizes.most_common()])
    add('  шрифты: %s' % [(k, v) for k, v in tbl_fonts.most_common()])
    add('')
    add('=== ЗАГОЛОВКИ')
    add('  кегли: %s' % [(k, v) for k, v in head_sizes.most_common()])
    add('')
    total_text = sum(text_sizes.values())
    main_ok = sum(v for k, v in text_sizes.items() if k == 14.0)
    add('доля текста кеглем 14 pt: %.1f%% (%d из %d знаков)' % (100.0 * main_ok / total_text if total_text else 0, main_ok, total_text))
    tbl_ok = sum(v for k, v in tbl_sizes.items() if k in (10.0, 12.0))
    total_tbl = sum(tbl_sizes.values())
    add('доля таблиц кеглем 10/12 pt: %.1f%%' % (100.0 * tbl_ok / total_tbl if total_tbl else 0))
    io.open(out_path, 'w', encoding='utf-8').write('\n'.join(out))
    print('ok')


if __name__ == '__main__':
    audit(sys.argv[1], sys.argv[2])
