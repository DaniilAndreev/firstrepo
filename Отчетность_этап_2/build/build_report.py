# -*- coding: utf-8 -*-
"""Сборка отчёта о выполнении работ (этап 2): docx → PDF → страницы содержания → docx → PDF.

Запуск: python3 -I build_report.py  (из папки build)
"""
import os
import sys
import subprocess
import shutil
import pdfplumber

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_helpers import *
import docx_helpers
import facts as F
import content_a as CA
import content_b as CB
import content_c as CC

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "..", "out")
PH = os.path.join(HERE, "placeholders")
DIAG = os.environ.get("DIAG_DIR", os.path.join(HERE, "..", "diagrams"))
os.makedirs(OUT, exist_ok=True)
os.makedirs(PH, exist_ok=True)

FIG_PATHS = {f"fig{i}": os.path.join(DIAG, f"fig{i}.png") for i in range(1, 7)}

DOCX_NAME = "Отчет_о_выполнении_работ_этап_2.docx"

TOC_ITEMS = [
    ("РЕФЕРАТ", 1),
    ("ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ", 1),
    ("ВВЕДЕНИЕ", 1),
    ("1 ПРОВЕДЁННЫЕ РАБОТЫ ПО ВТОРОМУ ЭТАПУ ГРАНТА", 1),
    (f"1.1 {F.PLAN[0]}", 2),
    (f"1.2 {F.PLAN[1]}", 2),
    (f"1.3 {F.PLAN[2]}", 2),
    (f"1.4 {F.PLAN[3]}", 2),
    (f"1.5 {F.PLAN[4]}", 2),
    ("ЗАКЛЮЧЕНИЕ", 1),
    ("ПЕРЕЧЕНЬ ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", 1),
]


def build(pages_map, counters):
    doc = new_document()
    remove_leading_empty_paragraph(doc)
    R, N, nfig, ntab = CA.make_resolver()
    docx_helpers.TEXT_FILTER = R
    CA.add_title(doc)
    CA.add_abstract(doc, counters)
    heading_struct(doc, "Содержание")
    for text, level in TOC_ITEMS:
        toc_entry(doc, text, pages_map.get(text, "0"), level=level)
    CA.add_abbrev(doc)
    CA.add_intro(doc, R, FIG_PATHS, PH)
    CA.add_1_1(doc, R, FIG_PATHS, PH)
    CB.add_1_2(doc, R, FIG_PATHS, PH)
    CB.add_1_3(doc, R, FIG_PATHS, PH)
    CC.add_1_4(doc, R, FIG_PATHS, PH)
    CC.add_1_5(doc, R, FIG_PATHS, PH)
    CC.add_conclusion(doc, R)
    CC.add_sources(doc)
    return doc, nfig, ntab


def to_pdf(docx_path):
    outdir = os.path.dirname(docx_path)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
    return docx_path[:-5] + ".pdf"


def page_map(pdf_path):
    """Номер страницы (печатный) для каждого заголовка из TOC_ITEMS. Титул = стр. 1."""
    found = {}
    with pdfplumber.open(pdf_path) as pdf:
        texts = [(p.extract_text() or "") for p in pdf.pages]
    n = len(texts)

    def is_toc(t):
        return t.strip().startswith("СОДЕРЖАНИЕ") or t.count("....") > 3

    for text, level in TOC_ITEMS:
        key = text if level == 1 else text[:28]
        for i in range(1, n):
            if is_toc(texts[i]):
                continue
            lines = [ln.strip() for ln in texts[i].splitlines()]
            if level == 1:
                if any(ln == text for ln in lines):
                    found[text] = i + 1
                    break
            else:
                if any(ln.startswith(key) for ln in lines):
                    found[text] = i + 1
                    break
    return found, n


def main():
    docx_path = os.path.join(OUT, DOCX_NAME)
    # проход 1
    doc, nfig, ntab = build({}, ("00", "0", "0", "0", "0"))
    doc.save(docx_path)
    pdf = to_pdf(docx_path)
    pm, npages = page_map(pdf)
    print("pass1 pages:", npages, pm)
    counters = (str(npages), str(nfig), str(ntab), str(CC.N_SOURCES), str(CC.N_APPS))
    # проход 2
    doc, nfig, ntab = build(pm, counters)
    doc.save(docx_path)
    pdf = to_pdf(docx_path)
    pm2, npages2 = page_map(pdf)
    print("pass2 pages:", npages2, pm2)
    if pm2 != pm or npages2 != npages:
        counters = (str(npages2), str(nfig), str(ntab), str(CC.N_SOURCES), str(CC.N_APPS))
        doc, nfig, ntab = build(pm2, counters)
        doc.save(docx_path)
        pdf = to_pdf(docx_path)
        pm3, npages3 = page_map(pdf)
        print("pass3 pages:", npages3, pm3)
        assert pm3 == pm2 and npages3 == npages2, "содержание не стабилизировалось"
    print("figures:", nfig, "tables:", ntab, "sources:", CC.N_SOURCES, "apps:", CC.N_APPS)
    print("OK:", docx_path, pdf)


if __name__ == "__main__":
    main()
