# -*- coding: utf-8 -*-
"""Проверка docx: запрещенные слова, URL вне перечня источников, нумерация и ссылки на объекты,
цитирование источников, порядок литер приложений, порядок элементов OOXML, размеры шрифта, метки [УТОЧНИТЬ]."""
import re
import sys
import zipfile
import collections
from lxml import etree
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

FORBID = [r"исследован", r"\bНИР\b", r"\bНИОКР\b", r"\bTRL\b", r"\bУГТ\b", r"\bмы\b", r"\bнами\b",
          r"\bнаш(а|е|и|у|ей|его|им|их|ему)?\b", r"\bя\b", r"уровн\w* готовности"]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SEQ = {
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
              "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
             "tcFitText", "vAlign", "hideMark"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
               "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
               "textDirection", "bidi", "rtlGutter", "docGrid"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
            "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
            "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
            "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
            "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
            "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
}


def blocks(doc):
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            yield "p", Paragraph(child, doc).text
        elif tag == "tbl":
            t = Table(child, doc)
            for row in t.rows:
                seen = set()
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    yield "c", c.text


def xml_order(path):
    z = zipfile.ZipFile(path)
    bad = collections.Counter()
    for part in [n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")]:
        root = etree.fromstring(z.read(part))
        for tag, seq in SEQ.items():
            for el in root.iter(W + tag):
                kids = [c.tag.replace(W, "") for c in el if isinstance(c.tag, str) and c.tag.startswith(W)]
                idx = [seq.index(k) for k in kids if k in seq]
                dup = tag in ("tblPr", "tcPr") and len(kids) != len(set(kids))
                if idx != sorted(idx) or dup:
                    bad[tag + ": " + " > ".join(kids)] += 1
    sizes = collections.Counter()
    root = etree.fromstring(z.read("word/document.xml"))
    for r in root.iter(W + "r"):
        if r.find(W + "t") is None:
            continue
        sz = r.find(W + "rPr/" + W + "sz")
        sizes[int(sz.get(W + "val")) / 2 if sz is not None else None] += 1
    return bad, sizes


def check(path, n_sources=None, sources_heading="ПЕРЕЧЕНЬ ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", allow_url=False):
    doc = Document(path)
    items = list(blocks(doc))
    probs = []
    in_src = False
    tabs, figs = [], []
    rt, rf = collections.defaultdict(list), collections.defaultdict(list)
    cites = set()
    utoch = []
    apps_first = []
    for i, (k, t) in enumerate(items):
        if t.strip().upper() == sources_heading:
            in_src = True
        for pat in FORBID:
            if allow_url and "доведено до уровня TRL 3" in t:
                break  # формулировка строки утвержденной формы отчета о развитии
            for m in re.finditer(pat, t, flags=re.IGNORECASE):
                probs.append(f"запрещенное «{m.group(0)}»: …{t[max(0, m.start() - 40):m.end() + 40]}…")
        if not in_src and not allow_url and re.search(r"https?://|www\.", t):
            probs.append(f"URL вне перечня источников: {t[:90]}")
        m = re.match(r"^Таблица (\d+) – ", t.strip())
        if m and k == "p":
            tabs.append((int(m.group(1)), i))
        m = re.match(r"^Рисунок (\d+) – ", t.strip())
        if m and k == "p":
            figs.append((int(m.group(1)), i))
        if not (m and k == "p"):
            for mm in re.finditer(r"таблиц(?:а|е|у|ы|и|ах)\s+(\d+)(?:\s+и\s+(\d+))?", t, flags=re.IGNORECASE):
                for g in mm.groups():
                    if g:
                        rt[int(g)].append(i)
            for mm in re.finditer(r"рисун(?:ок|ке|ка|ки|ках)\s+(\d+)", t, flags=re.IGNORECASE):
                rf[int(mm.group(1))].append(i)
        if not in_src:
            for mm in re.finditer(r"\[(\d+(?:\s*[,–-]\s*\d+)*)\]", t):
                for part in re.split(r"\s*,\s*", mm.group(1)):
                    if re.search(r"[–-]", part):
                        a, b = re.split(r"\s*[–-]\s*", part)
                        cites.update(range(int(a), int(b) + 1))
                    else:
                        cites.add(int(part))
            for mm in re.finditer(r"Приложени(?:и|е|я|ях)\s+([А-Я])\b", t):
                if mm.group(1) not in apps_first:
                    apps_first.append(mm.group(1))
            for mm in re.finditer(r"приложени(?:и|е|я)\s+([А-Я])\b", t):
                if mm.group(1) not in apps_first:
                    apps_first.append(mm.group(1))
        for mm in re.finditer(r"\[УТОЧНИТЬ[^\]]*\]", t):
            utoch.append(mm.group(0))
    for name, lst, refs in (("Таблица", tabs, rt), ("Рисунок", figs, rf)):
        nums = [n for n, _ in lst]
        if nums != list(range(1, len(nums) + 1)):
            probs.append(f"{name}: нумерация {nums}")
        for n, idx in lst:
            if allow_url:
                break  # отчет о развитии — утвержденная форма из одной таблицы без текста
            if not [j for j in refs.get(n, []) if j < idx]:
                probs.append(f"{name} {n}: нет ссылки в тексте до объекта")
        for n in refs:
            if n not in nums:
                probs.append(f"{name} {n}: ссылка на несуществующий объект")
    if n_sources:
        miss = [i for i in range(1, n_sources + 1) if i not in cites]
        if miss:
            probs.append(f"не процитированы источники: {miss}")
        extra = [c for c in cites if c > n_sources]
        if extra:
            probs.append(f"ссылки на несуществующие источники: {extra}")
    bad, sizes = xml_order(path)
    for k, v in bad.items():
        probs.append(f"порядок XML: {v} x {k[:120]}")
    print(f"== {path.split('/')[-1]}: блоков {len(items)}, таблиц {len(tabs)}, рисунков {len(figs)}, "
          f"[УТОЧНИТЬ] {len(utoch)}, приложения по порядку упоминания: {''.join(apps_first) or '-'}")
    print("   размеры шрифта (пт: кол-во прогонов):", dict(sorted(((k or 0), v) for k, v in sizes.items())))
    for p in probs:
        print("   ПРОБЛЕМА:", p)
    if not probs:
        print("   проблем нет")
    return utoch


if __name__ == "__main__":
    args = sys.argv[1:]
    for a in args:
        path, _, ns = a.partition("::")
        check(path, int(ns) if ns else None, allow_url="развитии" in path)
