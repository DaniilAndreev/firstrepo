# -*- coding: utf-8 -*-
"""Список меток [УТОЧНИТЬ] с контекстом по каждому docx: python3 -I list_utochnit.py <папка> > список.txt

Для ячейки таблицы контекст — номер таблицы, первая ячейка строки и заголовок столбца.
"""
import os
import re
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

MARK = re.compile(r"\[УТОЧНИТЬ[^\]]*\]")


def clean(t):
    return " ".join(t.split())


def fragments(doc):
    """(текст, подпись-контекст) для абзацев и ячеек таблиц в порядке документа."""
    caption = ""
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            text = clean(Paragraph(child, doc).text)
            m = re.match(r"^(Таблица [\dА-Я.]+)", text)
            if m:
                caption = m.group(1)
            yield text, ""
        elif tag == "tbl":
            rows = Table(child, doc).rows
            header = [clean(c.text) for c in rows[0].cells]
            for row in rows[1:]:
                seen = set()
                cells = row.cells
                for ci, c in enumerate(cells):
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    head = header[ci] if ci < len(header) else ""
                    key = clean(cells[0].text)
                    if key.isdigit() and len(cells) > 1:
                        key = f"{key}. {clean(cells[1].text)}"
                    label = f"{caption}: «{key[:50]}» / {head}" if ci else f"{caption}: {head}"
                    yield clean(c.text), label


def main(root):
    files = []
    for dp, _, fs in os.walk(root):
        files += [os.path.join(dp, f) for f in fs if f.endswith(".docx") and not f.startswith("~$")]
    order = ["Отчет_о_выполнении", "Бизнес-план", "Отчет_о_развитии", "Приложение"]
    files.sort(key=lambda p: (next((i for i, k in enumerate(order) if k in os.path.basename(p)), 9),
                              os.path.basename(p)))
    total = 0
    for path in files:
        items = []
        for text, label in fragments(Document(path)):
            for m in MARK.finditer(text):
                before = text[max(0, m.start() - 70):m.start()].strip()
                if m.start() > 70:
                    before = "…" + before
                ctx = f"{label} — {before}" if label and before else (label or before)
                items.append((m.group(0), ctx))
        total += len(items)
        print(f"== {os.path.basename(path)} — {len(items)}")
        for i, (mark, ctx) in enumerate(items, 1):
            print(f"{i:3}. {mark}  ← {ctx}")
        print()
    print(f"Всего меток: {total}")


if __name__ == "__main__":
    main(sys.argv[1])
