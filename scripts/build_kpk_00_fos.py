# -*- coding: utf-8 -*-
"""Перенос ФОС (03, 04, 05) в типовые задания программы 00.

  5.1.5.1 <- 03 (20 вопросов + ключи)
  5.2.5.1 <- 04 (20 вопросов + ключи)
  5.4.5.1 <- 05 (15 билетов x 20 вопросов + матрица ответов)

Форматирование берётся из самой программы: абзац вопроса, полужирный заголовок
и таблица ключей клонируются из существующих элементов, поэтому вид не меняется.

Запуск: python -X utf8 build_fos.py <00.docx> <03.docx> <04.docx> <05.docx> <out.docx>
"""
import copy
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

Q_RE = re.compile(r'^\d+\.\s')


def get_block_pars(doc):
    return [p for p in doc.paragraphs]


def fos_questions(path, limit=None):
    doc = Document(path)
    out = []
    for p in doc.paragraphs:
        t = ' '.join(p.text.split())
        if Q_RE.match(t):
            out.append(t)
    return out[:limit] if limit else out


def fos_keys(path, ti=1):
    doc = Document(path)
    t = doc.tables[ti]
    keys = []
    for r in t.rows[1:]:
        cells = [' '.join(c.text.split()) for c in r.cells]
        if len(cells) >= 2 and cells[0].isdigit():
            keys.append((cells[0], cells[1]))
    return keys


def fos_tickets(path):
    """Список (номер, [20 вопросов]) и матрица ответов 15x20."""
    doc = Document(path)
    tickets = []
    cur = None
    for p in doc.paragraphs:
        t = ' '.join(p.text.split())
        m = re.match(r'^Билет\s*№\s*(\d+)$', t)
        if m:
            cur = (int(m.group(1)), [])
            tickets.append(cur)
            continue
        if cur is not None and Q_RE.match(t):
            cur[1].append(t)
    matrix = []
    t = doc.tables[1]
    for r in t.rows[2:]:
        cells = [' '.join(c.text.split()) for c in r.cells]
        if cells[0].isdigit():
            matrix.append(cells[1:21])
    return tickets, matrix


def set_cell_text(cell, text):
    par = cell.paragraphs[0]
    runs = par.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ''
    else:
        par.add_run(text)


def clone_par(template_par, text, bold=None):
    """Копия абзаца-образца с новым текстом."""
    new_p = copy.deepcopy(template_par._p)
    runs = new_p.findall(qn('w:r'))
    if not runs:
        return new_p
    first = runs[0]
    for r in runs[1:]:
        new_p.remove(r)
    for t in first.findall(qn('w:t')):
        first.remove(t)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    first.append(t)
    if bold is not None:
        rPr = first.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            first.insert(0, rPr)
        b = rPr.find(qn('w:b'))
        if bold and b is None:
            rPr.append(OxmlElement('w:b'))
        elif not bold and b is not None:
            rPr.remove(b)
    return new_p


def clear_block(heading_p):
    """Удаляет всё между заголовком и следующим заголовком; возвращает якорь."""
    el = heading_p._p
    removed = 0
    nxt = el.getnext()
    while nxt is not None:
        tag = nxt.tag.split('}')[1]
        if tag == 'sectPr':
            break          # завершающую sectPr удалять нельзя
        is_heading = False
        if tag == 'p':
            pPr = nxt.find(qn('w:pPr'))
            st = pPr.find(qn('w:pStyle')) if pPr is not None else None
            style = st.get(qn('w:val')) if st is not None else ''
            txt = ''.join(x.text or '' for x in nxt.iter(qn('w:t')))
            if style in ('1', '2', 'ae') or re.match(r'^5\.\d', txt.strip()):
                is_heading = True
        if is_heading:
            break
        after = nxt.getnext()
        nxt.getparent().remove(nxt)
        removed += 1
        nxt = after
    return removed


def find_heading(doc, prefix):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(prefix) and p.style.name.startswith('Heading'):
            return i, p
    raise KeyError('не найден заголовок ' + prefix)


def insert_after(anchor, element):
    anchor.addnext(element)
    return element


def build_key_table(template_tbl, keys, rows_needed=None):
    tbl = copy.deepcopy(template_tbl)
    trs = tbl.findall(qn('w:tr'))
    # обрезать/дополнить строки не нужно: шаблон 21x2 под 20 вопросов
    doc_rows = len(trs)
    for i in range(1, min(doc_rows, len(keys) + 1)):
        num, ans = keys[i - 1]
        tcs = trs[i].findall(qn('w:tc'))
        if len(tcs) >= 2:
            # номер
            p0 = tcs[0].find(qn('w:p'))
            r0 = p0.find(qn('w:r')) if p0 is not None else None
            if r0 is not None:
                for t in r0.findall(qn('w:t')):
                    r0.remove(t)
                t = OxmlElement('w:t')
                t.text = num
                r0.append(t)
            p1 = tcs[1].find(qn('w:p'))
            r1 = p1.find(qn('w:r')) if p1 is not None else None
            if r1 is not None:
                for t in r1.findall(qn('w:t')):
                    r1.remove(t)
                t = OxmlElement('w:t')
                t.text = ans
                r1.append(t)
    return tbl


def set_grid_tbl(tbl, widths_tw):
    """Задать сетку таблицы в twips и убрать отступ."""
    grid = tbl.find(qn('w:tblGrid'))
    cols = list(grid)
    if len(cols) != len(widths_tw):
        raise ValueError('grid mismatch: %d cols vs %d widths' % (len(cols), len(widths_tw)))
    for gc, w in zip(cols, widths_tw):
        gc.set(qn('w:w'), str(w))
    tblPr = tbl.find(qn('w:tblPr'))
    tw = tblPr.find(qn('w:tblW'))
    if tw is None:
        tw = OxmlElement('w:tblW')
        tblPr.append(tw)
    tw.set(qn('w:type'), 'dxa')
    tw.set(qn('w:w'), str(sum(widths_tw)))
    ind = tblPr.find(qn('w:tblInd'))
    if ind is None:
        ind = OxmlElement('w:tblInd')
        tblPr.append(ind)
    ind.set(qn('w:type'), 'dxa')
    ind.set(qn('w:w'), '0')
    for tr in tbl.findall(qn('w:tr')):
        ci = 0
        for tc in tr.findall(qn('w:tc')):
            span = 1
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    span = int(gs.get(qn('w:val')))
            w = sum(widths_tw[ci:ci + span])
            if tcPr is not None:
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    tcW.set(qn('w:type'), 'dxa')
                    tcW.set(qn('w:w'), str(w))
            ci += span


def main(src00, src03, src04, src05, dst):
    doc = Document(src00)
    q03 = fos_questions(src03)
    q04 = fos_questions(src04)
    k03 = fos_keys(src03)
    k04 = fos_keys(src04)
    tickets, matrix = fos_tickets(src05)
    print('03: %d вопросов, %d ключей' % (len(q03), len(k03)))
    print('04: %d вопросов, %d ключей' % (len(q04), len(k04)))
    print('05: билетов %d, матрица %dx%d' % (len(tickets), len(matrix), len(matrix[0]) if matrix else 0))

    # образцы форматирования
    h151_i, h151 = find_heading(doc, '5.1.5.1.')
    tmpl_q = doc.paragraphs[h151_i + 1]
    tmpl_bold = None
    for p in doc.paragraphs:
        if p.text.strip() == 'Ключи ответов':
            tmpl_bold = p
            break
    tmpl_key_tbl = None
    for t in doc.tables:
        if t.rows and t.rows[0].cells[0].text.strip().startswith('№ вопроса') and len(t.rows) == 21:
            tmpl_key_tbl = t._tbl
            break
    assert tmpl_bold is not None and tmpl_key_tbl is not None

    # --- 5.1.5.1
    removed = clear_block(h151)
    anchor = h151._p
    for q in q03:
        anchor = insert_after(anchor, clone_par(tmpl_q, q))
    anchor = insert_after(anchor, clone_par(tmpl_bold, 'Ключи ответов', bold=True))
    anchor = insert_after(anchor, build_key_table(tmpl_key_tbl, k03))
    print('5.1.5.1: удалено %d элементов, вставлено %d вопросов' % (removed, len(q03)))

    # --- 5.2.5.1
    h251_i, h251 = find_heading(doc, '5.2.5.1.')
    removed = clear_block(h251)
    anchor = h251._p
    for q in q04:
        anchor = insert_after(anchor, clone_par(tmpl_q, q))
    anchor = insert_after(anchor, clone_par(tmpl_bold, 'Ключи ответов', bold=True))
    anchor = insert_after(anchor, build_key_table(tmpl_key_tbl, k04))
    print('5.2.5.1: удалено %d элементов, вставлено %d вопросов' % (removed, len(q04)))

    # --- 5.4.5.1
    h451_i, h451 = find_heading(doc, '5.4.5.1.')
    removed = clear_block(h451)
    anchor = h451._p
    for num, qs in tickets:
        anchor = insert_after(anchor, clone_par(tmpl_bold, 'Билет №%d' % num, bold=True))
        for q in qs:
            anchor = insert_after(anchor, clone_par(tmpl_q, q))
    anchor = insert_after(anchor, clone_par(tmpl_bold, 'Ключи ответов', bold=True))
    # матрица: копия таблицы из 05 с явными границами
    src_doc05 = Document(src05)
    mtbl = copy.deepcopy(src_doc05.tables[1]._tbl)
    tblPr = mtbl.find(qn('w:tblPr'))
    if tblPr is not None and tblPr.find(qn('w:tblBorders')) is None:
        borders = OxmlElement('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            el = OxmlElement('w:' + side)
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '4')
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), 'auto')
            borders.append(el)
        tblPr.append(borders)
    set_grid_tbl(mtbl, [700] + [432] * 20)   # 467 pt под полосу программы
    # фиксированная разметка + мелкий кегль в шапке, иначе Word расширяет колонки под слово «Вопрос»
    mtblPr = mtbl.find(qn('w:tblPr'))
    lay = mtblPr.find(qn('w:tblLayout'))
    if lay is None:
        lay = OxmlElement('w:tblLayout')
        mtblPr.append(lay)
    lay.set(qn('w:type'), 'fixed')
    head = mtbl.findall(qn('w:tr'))[0]
    for tc in head.findall(qn('w:tc')):
        for r in tc.iter(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)
            for tag, val in (('w:sz', '15'), ('w:szCs', '15')):
                el = rPr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    rPr.append(el)
                el.set(qn('w:val'), val)
    anchor = insert_after(anchor, mtbl)
    print('5.4.5.1: удалено %d элементов, вставлено %d билетов' % (removed, len(tickets)))

    doc.save(dst)
    print('saved', dst)


if __name__ == '__main__':
    main(*sys.argv[1:6])
