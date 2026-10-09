# -*- coding: utf-8 -*-
"""Автопроверка docx: запрещённые слова, сквозная нумерация, ссылки на объекты, список [УТОЧНИТЬ]."""
import re, sys
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

FORBID = [r"исследован", r"исследовани", r"\bНИР\b", r"\bНИОКР\b", r"\bTRL\b", r"\bУГТ\b",
          r"\bмы\b", r"\bнами\b", r"\bнаш(а|е|и|у|ей|его|им|их|ему)?\b", r"\bя\b", r"уровн\w* готовности"]
URL = [r"https?://", r"www\."]

def iter_texts(doc):
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split('}')[1]
        if tag == 'p':
            yield ('p', Paragraph(child, doc).text)
        elif tag == 'tbl':
            t = Table(child, doc)
            for row in t.rows:
                seen=set()
                for c in row.cells:
                    if id(c._tc) in seen: continue
                    seen.add(id(c._tc))
                    yield ('c', c.text)

def check(path):
    doc = Document(path)
    items = list(iter_texts(doc))
    print(f"=== {path}: {len(items)} блоков")
    problems = []
    in_sources = False
    tabs, figs = [], []
    refs_t, refs_f = {}, {}
    utoch = []
    for idx, (kind, t) in enumerate(items):
        if t.strip().upper().startswith("ПЕРЕЧЕНЬ ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
            in_sources = True
        for pat in FORBID:
            for m in re.finditer(pat, t, flags=re.IGNORECASE):
                problems.append((idx, pat, t[max(0,m.start()-50):m.end()+50]))
        if not in_sources:
            for pat in URL:
                for m in re.finditer(pat, t, flags=re.IGNORECASE):
                    problems.append((idx, 'URL', t[max(0,m.start()-50):m.end()+50]))
        m = re.match(r"^Таблица (\d+) – (.+)$", t.strip())
        if m and kind == 'p':
            tabs.append((int(m.group(1)), idx, m.group(2)))
        m = re.match(r"^Рисунок (\d+) – (.+)$", t.strip())
        if m and kind == 'p':
            figs.append((int(m.group(1)), idx, m.group(2)))
        for m in re.finditer(r"таблиц[аеуы] (\d+)", t, flags=re.IGNORECASE):
            refs_t.setdefault(int(m.group(1)), []).append(idx)
        for m in re.finditer(r"рисун(?:ок|ке|ка) (\d+)", t, flags=re.IGNORECASE):
            refs_f.setdefault(int(m.group(1)), []).append(idx)
        for m in re.finditer(r"\[УТОЧНИТЬ[^\]]*\]", t):
            utoch.append((idx, m.group(0), t[max(0,m.start()-60):m.start()].strip()))
    # нумерация
    def seq_check(name, lst):
        nums = [n for n,_,_ in lst]
        exp = list(range(1, len(nums)+1))
        if nums != exp:
            problems.append((-1, f'{name}: нумерация', f'{nums}'))
    seq_check('Таблицы', tabs); seq_check('Рисунки', figs)
    for n, idx, title in tabs:
        r = [i for i in refs_t.get(n, []) if i < idx]
        if not r:
            problems.append((idx, 'нет ссылки до объекта', f'Таблица {n} – {title}'))
    for n, idx, title in figs:
        r = [i for i in refs_f.get(n, []) if i < idx]
        if not r:
            problems.append((idx, 'нет ссылки до объекта', f'Рисунок {n} – {title}'))
    # ссылки на несуществующие объекты
    for n in refs_t:
        if n not in [x[0] for x in tabs]:
            problems.append((-1, 'ссылка на несуществующую таблицу', str(n)))
    for n in refs_f:
        if n not in [x[0] for x in figs]:
            problems.append((-1, 'ссылка на несуществующий рисунок', str(n)))
    print(f"таблиц: {len(tabs)}, рисунков: {len(figs)}, [УТОЧНИТЬ]: {len(utoch)}")
    for p in problems:
        print("  ПРОБЛЕМА:", p)
    if not problems:
        print("  проблем не найдено")
    return utoch

if __name__ == '__main__':
    allu = {}
    for p in sys.argv[1:]:
        allu[p] = check(p)
    with open('utochnit.txt','w',encoding='utf-8') as f:
        for p, lst in allu.items():
            f.write(f"\n### {p}\n")
            for idx, tag, ctx in lst:
                f.write(f"- [{idx}] {tag}  (контекст: …{ctx[-60:]})\n")
    print("список [УТОЧНИТЬ] записан в utochnit.txt")
