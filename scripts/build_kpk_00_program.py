# -*- coding: utf-8 -*-
"""Сборка мега-документа 00 (Программа) — правка вёрстки и нумерации.

Что делает:
  1. Переводит секции «Учебно-тематический план» и «Календарный учебный график»
     в альбомную ориентацию A4 (поля сохраняются, полоса набора 714 pt).
  2. Пересчитывает ширины таблиц под полосу своей секции: 467 pt — портрет,
     714 pt — альбом; для широких таблиц приложений задаёт осмысленные колонки.
  3. Делает сквозную нумерацию подразделов внутри рабочих программ разделов:
     1 -> 5.N.1, 1.1 -> 5.N.1.1, 3.1 -> 5.N.3.1, 5.1 -> 5.N.5.1.
     Запись оглавления для программы раздела берётся из абзаца
     «5.N. Рабочая программа раздела …», дублирующий титульный заголовок
     выводится из структуры (внешний вид не меняется).
  4. Добавляет завершающую sectPr, если её нет (иначе хвост документа
     оказывается без параметров страницы), убирает пустые <w:u w:val="none"/>.

После сборки поля и оглавление нужно обновить в Word — для этого
рядом лежит update_fields_word.ps1 (он же сохраняет итоговый docx).

Запуск:
    python -X utf8 build_kpk_00_program.py --src in.docx --out out.docx [--log log.txt]
"""
import argparse
import copy
import io
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

A4_W, A4_H = 11900, 16840          # twips, портрет
TEXT_W_PORTRAIT = 467              # pt
TEXT_W_LANDSCAPE = 714             # pt
TW = 20.0
LANDSCAPE_SECTIONS = (2, 3)        # индексы секций: УТП и календарный график

log = []


def say(msg):
    log.append(msg)
    print(msg.encode('ascii', 'replace').decode('ascii'))


# ---------------------------------------------------------------- таблицы

def set_grid(t, widths_tw):
    tbl = t._tbl
    grid = tbl.find(qn('w:tblGrid'))
    cols = list(grid)
    if len(cols) != len(widths_tw):
        raise ValueError('grid mismatch: %d cols vs %d widths' % (len(cols), len(widths_tw)))
    total = sum(widths_tw)
    for gc, w in zip(cols, widths_tw):
        gc.set(qn('w:w'), str(w))
    tblPr = tbl.tblPr
    tw = tblPr.find(qn('w:tblW'))
    if tw is None:
        tw = OxmlElement('w:tblW')
        tblPr.append(tw)
    tw.set(qn('w:type'), 'dxa')
    tw.set(qn('w:w'), str(total))
    ind = tblPr.find(qn('w:tblInd'))
    if ind is None:
        ind = OxmlElement('w:tblInd')
        tblPr.append(ind)
    ind.set(qn('w:type'), 'dxa')
    ind.set(qn('w:w'), '0')
    for tr in tbl.tr_lst:
        ci = 0
        for tc in tr.tc_lst:
            span = 1
            if tc.tcPr is not None:
                gs = tc.tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    span = int(gs.get(qn('w:val')))
            w = sum(widths_tw[ci:ci + span])
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW')
                tcPr.insert(0, tcW)
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(w))
            ci += span


def grid_tw(t):
    grid = t._tbl.find(qn('w:tblGrid'))
    return [int(gc.get(qn('w:w'))) for gc in grid]


def scale_to(t, target_pt):
    cur = grid_tw(t)
    k = target_pt * TW / sum(cur)
    new = [max(1, int(round(w * k))) for w in cur]
    new[-1] += int(target_pt * TW) - sum(new)
    set_grid(t, new)


def semantic_widths(t, target_pt):
    """Осмысленные ширины для таблиц, скопированных из альбомного шаблона (714 pt)."""
    cur = grid_tw(t)
    total = sum(cur)
    target_tw = int(target_pt * TW)
    n = len(cur)
    if abs(total - 14280) > 400:
        return None
    if n == 6:                      # № | наименование | всего | теор. | практ. | форма
        base = [800, 3400, 1200, 1300, 1300, 1340]
    elif n == 2:                    # № вопроса | ответ
        base = [1000, 8340]
    elif n == 3:                    # знать | уметь | владеть
        base = [3113, 3113, 3114]
    else:
        return None
    scale = target_tw / float(sum(base))
    out = [int(round(w * scale)) for w in base]
    out[-1] += target_tw - sum(out)
    return out


def table_section_map(doc):
    """Индекс секции для каждой таблицы в порядке следования."""
    res = []
    sec = 0
    for el in doc.element.body.iterchildren():
        tag = el.tag.split('}')[1]
        if tag == 'p':
            pPr = el.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                sec += 1
        elif tag == 'tbl':
            res.append(sec)
    return res


def resize_tables(doc):
    secs = table_section_map(doc)
    changed = 0
    for ti, t in enumerate(doc.tables):
        tblPr = t._tbl.tblPr
        ind = tblPr.find(qn('w:tblInd'))
        if ind is None:
            ind = OxmlElement('w:tblInd')
            tblPr.append(ind)
        if ind.get(qn('w:w')) not in (None, '0'):
            ind.set(qn('w:type'), 'dxa')
            ind.set(qn('w:w'), '0')
            say('  T%-2d отступ таблицы сброшен' % ti)
        sec = secs[ti]
        target = TEXT_W_LANDSCAPE if sec in LANDSCAPE_SECTIONS else TEXT_W_PORTRAIT
        cur_pt = sum(grid_tw(t)) / TW
        if abs(cur_pt - target) < 1.0:
            continue
        sem = semantic_widths(t, target) if target == TEXT_W_PORTRAIT else None
        if sem:
            set_grid(t, sem)
        else:
            scale_to(t, target)
        changed += 1
        say('  T%-2d S%d %6.1f -> %6.1f pt %s' % (ti, sec, cur_pt, target, '(semantic)' if sem else ''))
    say('таблиц перестроено: %d' % changed)


# ---------------------------------------------------------------- секции

def make_landscape(section):
    pgSz = section._sectPr.find(qn('w:pgSz'))
    if pgSz is None:
        pgSz = OxmlElement('w:pgSz')
        section._sectPr.append(pgSz)
    pgSz.set(qn('w:w'), str(A4_H))
    pgSz.set(qn('w:h'), str(A4_W))
    pgSz.set(qn('w:orient'), 'landscape')
    typ = section._sectPr.find(qn('w:type'))
    if typ is not None and typ.get(qn('w:val')) == 'continuous':
        typ.getparent().remove(typ)


def add_final_sectpr(doc, template_sect):
    body = doc.element.body
    if body.find(qn('w:sectPr')) is not None:
        say('завершающая sectPr уже есть — не добавляем')
        return
    new = copy.deepcopy(template_sect)
    typ = new.find(qn('w:type'))
    if typ is not None:
        new.remove(typ)
    pgn = new.find(qn('w:pgNumType'))
    if pgn is not None:
        new.remove(pgn)
    pgSz = new.find(qn('w:pgSz'))
    if pgSz is not None:
        pgSz.set(qn('w:w'), str(A4_W))
        pgSz.set(qn('w:h'), str(A4_H))
        if pgSz.get(qn('w:orient')):
            del pgSz.attrib[qn('w:orient')]
    body.append(new)
    say('добавлена завершающая портретная sectPr')


# ---------------------------------------------------------------- нумерация

BLOCK_RE = re.compile(r'^5\.(\d)\.\s*Рабочая программа раздела')
NUM_RE = re.compile(r'^(\d(?:\.\d){0,2})\.\s+')
INNER_OK = {
    '1.1': ('Цель', 'Задачи', 'Рабочая программа раздела', 'Характеристика',
            'Производственная практика', 'Итоговая аттестация', 'К итоговой аттестации'),
    '1.2': ('Задачи', 'Рабочая программа раздела', 'Производственная практика', 'Итоговая аттестация'),
    '1.3': ('Рабочая программа раздела', 'Характеристика', 'К итоговой аттестации'),
    '1.4': ('Характеристика',),
    '3.1': ('Учебно-тематический план',),
    '3.2': ('Содержание',),
    '5.1': ('Типовые задания',),
}


def set_outline(par, level):
    """Уровень в структуре документа (1 = первый)."""
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn('w:outlineLvl'))
    if el is None:
        el = OxmlElement('w:outlineLvl')
        pPr.append(el)
    el.set(qn('w:val'), str(level - 1))


def set_par_text(par, text):
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ''


def demote_title(par):
    """Титульный заголовок программы визуально остаётся заголовком, но выходит из структуры."""
    par.style = par.part.document.styles['Normal']
    pf = par.paragraph_format
    pf.keep_with_next = True
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(24)
    pf.space_after = Pt(6)
    for r in par.runs:
        r.font.bold = True
        r.font.size = Pt(14)


def toc_sources(doc):
    """Запись оглавления для программы раздела — абзац «5.N. Рабочая программа раздела …»."""
    pars = doc.paragraphs
    for i, p in enumerate(pars):
        if not BLOCK_RE.match(p.text.strip()):
            continue
        set_outline(p, 2)
        for j in range(i + 1, min(i + 12, len(pars))):
            t = pars[j].text.strip()
            if t.startswith('«') and t.endswith('»'):
                demote_title(pars[j])
                say('  запись оглавления: p%d вместо p%d (%s)' % (i, j, t[:40]))
                break


def renumber(doc):
    pars = doc.paragraphs
    starts = [i for i, p in enumerate(pars) if BLOCK_RE.match(p.text.strip())]
    if len(starts) != 4:
        say('ВНИМАНИЕ: найдено блоков программ разделов: %d' % len(starts))
    for bi, start in enumerate(starts):
        end = starts[bi + 1] if bi + 1 < len(starts) else len(pars)
        n = BLOCK_RE.match(pars[start].text.strip()).group(1)
        say('--- блок 5.%s (абзацы %d..%d)' % (n, start, end - 1))
        for i in range(start + 1, end):
            p = pars[i]
            txt = p.text.strip()
            m = NUM_RE.match(txt)
            if not m:
                continue
            prefix = m.group(1)
            rest = txt[m.end():]
            is_heading = p.style.name.startswith('Heading')
            keep = False
            if is_heading and '.' not in prefix:
                keep = True
            elif prefix in INNER_OK and rest.startswith(INNER_OK[prefix]):
                keep = True
            if not keep:
                continue
            new_text = '5.%s.%s. %s' % (n, prefix, rest)
            say('  p%-4d %-45s -> %s' % (i, txt[:45], new_text[:60]))
            set_par_text(p, new_text)
            if prefix.count('.') == 1 and is_heading:
                set_outline(p, 3)


# ---------------------------------------------------------------- мелочи

def keep_toc_field(doc):
    """Абзац с подсказкой — это кэш поля TOC; удалять его нельзя, иначе Word нечего обновлять."""
    for p in doc.paragraphs:
        if p.text.strip().startswith('Оглавление обновляется при открытии'):
            fld = list(p._p.iter(qn('w:fldChar')))
            instr = [t.text for t in p._p.iter(qn('w:instrText'))]
            say('абзац поля TOC сохранён (fldChar=%s, instr=%s)' % (len(fld), instr))
            return


def clean_u_none(doc):
    removed = 0
    for el in doc.element.body.iter(qn('w:u')):
        if el.get(qn('w:val')) == 'none':
            el.getparent().remove(el)
            removed += 1
    say('удалено тегов w:u val="none": %d' % removed)


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description='Правка вёрстки и нумерации мега-документа 00')
    ap.add_argument('--src', required=True, help='исходный docx')
    ap.add_argument('--out', required=True, help='куда сохранить результат')
    ap.add_argument('--log', default=None, help='файл отчёта (по умолчанию <out>.log)')
    args = ap.parse_args()
    log_path = args.log or (args.out + '.log')

    doc = Document(args.src)
    say('секций: %d' % len(doc.sections))
    portrait_template = copy.deepcopy(doc.sections[1]._sectPr)

    for idx in LANDSCAPE_SECTIONS:
        make_landscape(doc.sections[idx])
        say('S%d -> landscape %.0fx%.0f pt' % (idx, doc.sections[idx].page_width / 12700.0,
                                               doc.sections[idx].page_height / 12700.0))
    add_final_sectpr(doc, portrait_template)

    say('--- таблицы')
    resize_tables(doc)

    say('--- нумерация')
    renumber(doc)
    toc_sources(doc)

    say('--- прочее')
    keep_toc_field(doc)
    clean_u_none(doc)

    doc.save(args.out)
    say('saved %s' % args.out)
    io.open(log_path, 'w', encoding='utf-8').write('\n'.join(log))


if __name__ == '__main__':
    main()
