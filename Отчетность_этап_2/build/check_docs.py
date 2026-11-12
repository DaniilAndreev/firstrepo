# -*- coding: utf-8 -*-
"""Проверка документов отчетности.

Общие проверки (любой docx): недопустимые формулировки, нумерация таблиц и рисунков и ссылки на них,
цитирование источников, порядок элементов OOXML, метки [УТОЧНИТЬ].

Проверки отчета о выполнении работ и приложений по правилам оформления Фонда (замечания от 05.11.2026):
поля 30/15/20/20 мм; абзацный отступ 1,25 см; интервал 1,15; Times New Roman 12 единым кеглем; без
полужирного и курсива; текст по ширине, заголовки по центру; номер страницы внизу по центру, первая страница
файла — вторая (реферат); объем не более 20 страниц с титульным листом; нет подзаголовков вне содержания;
тема и работы календарного плана дословно; кликабельная ссылка на сайт в подразделе о сайте; на каждый объект
приложений есть ссылка в отчете, и каждая ссылка отчета ведет на существующий объект.

Запуск:
  python3 -I check_docs.py --report out/Отчет….docx out/Приложения….docx
  python3 -I check_docs.py "out/Бизнес-план.docx::17" "out/Отчет_о_развитии_проекта.docx"
"""
import os
import re
import sys
import zipfile
import collections
import pdfplumber
from lxml import etree
from docx import Document
from docx.shared import Mm
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH as WA

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import facts as F  # noqa: E402
import apx  # noqa: E402

# недопустимые формулировки (замечания Фонда к отчету, действуют для всех документов проекта)
FORBID = [r"исследован", r"\bНИР\b", r"\bНИОКР\b", r"\bОКР\b", r"\bTRL\b", r"\bУГТ\b", r"\bMVP\b",
          r"уровн\w* готовности", r"эксперимент", r"\bнаучн", r"лабораторн", r"\bопытн", r"\bтестов",
          r"\bмакет", r"\bГОСТ\b", r"СанПиН", r"\bТУ\b", r"перв\w+ этап", r"\bмы\b", r"\bнами\b",
          r"\bнаш(а|е|и|у|ей|его|им|их|ему)?\b", r"\bя\b"]
# официальное наименование Фонда и заголовок шаблона бизнес-плана Фонда
ALLOW = ["в научно-технической сфере", "научно-технического задела"]

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
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


def norm(t):
    return " ".join(t.replace(" ", " ").replace("­", "").split())


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


def forbidden(t):
    out = []
    clean = t
    for a in ALLOW:
        clean = clean.replace(a, " " * len(a))
    for pat in FORBID:
        for m in re.finditer(pat, clean, flags=re.IGNORECASE):
            out.append(f"недопустимое «{m.group(0)}»: …{t[max(0, m.start() - 40):m.end() + 40]}…")
    return out


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


def check(path, n_sources=None, sources_heading="ПЕРЕЧЕНЬ ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", allow_url=False,
          lettered=False, quiet=False, ordered_sources=False):
    """Общие проверки. lettered=True — приложения: номера объектов с литерой, ссылки внутри не обязательны."""
    doc = Document(path)
    items = list(blocks(doc))
    probs = []
    in_src = False
    tabs, figs = [], []
    rt, rf = collections.defaultdict(list), collections.defaultdict(list)
    cites = set()
    utoch = []
    for i, (k, t) in enumerate(items):
        if t.strip().upper() == sources_heading:
            in_src = True
        if not (allow_url and "доведено до уровня TRL 3" in t):
            probs += forbidden(t)
        m = re.match(r"^Таблица (\d+) – ", t.strip())
        if m and k == "p":
            tabs.append((int(m.group(1)), i))
        m = re.match(r"^Рисунок (\d+) – ", t.strip())
        if m and k == "p":
            figs.append((int(m.group(1)), i))
        if not (m and k == "p"):
            for mm in re.finditer(r"таблиц(?:а|е|у|ы|и|ах)\s+(\d+)(?![\dА-Я])(?:\s+и\s+(\d+)(?![\dА-Я]))?", t,
                                  flags=re.IGNORECASE):
                for g in mm.groups():
                    if g:
                        rt[int(g)].append(i)
            for mm in re.finditer(r"рисун(?:ок|ке|ка|ки|ках)\s+(\d+)(?![\dА-Я])", t, flags=re.IGNORECASE):
                rf[int(mm.group(1))].append(i)
        if not in_src:
            for mm in re.finditer(r"\[(\d+(?:\s*[,–-]\s*\d+)*)\]", t):
                for part in re.split(r"\s*,\s*", mm.group(1)):
                    if re.search(r"[–-]", part):
                        a, b = re.split(r"\s*[–-]\s*", part)
                        cites.update(range(int(a), int(b) + 1))
                    else:
                        cites.add(int(part))
        for mm in re.finditer(r"\[УТОЧНИТЬ[^\]]*\]", t):
            utoch.append(mm.group(0))
    if not lettered:
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
        order = []
        for i, (k, t) in enumerate(items):
            if t.strip().upper() == sources_heading:
                break
            for mm in re.finditer(r"\[(\d+)\]", t):
                if int(mm.group(1)) not in order:
                    order.append(int(mm.group(1)))
        if ordered_sources and order != sorted(order):
            probs.append(f"источники пронумерованы не в порядке первого упоминания: {order}")
    bad, sizes = xml_order(path)
    for k, v in bad.items():
        probs.append(f"порядок XML: {v} x {k[:120]}")
    if utoch:
        probs.append(f"меток [УТОЧНИТЬ]: {len(utoch)}")
    if not quiet:
        print(f"== {os.path.basename(path)}: блоков {len(items)}, таблиц {len(tabs)}, рисунков {len(figs)}")
        print("   размеры шрифта (пт: кол-во прогонов):", dict(sorted(((k or 0), v) for k, v in sizes.items())))
    return probs


# ------------------------------------------------------------------ оформление по правилам Фонда
def _style_chain(style):
    while style is not None:
        yield style
        style = style.base_style


def eff(par, attr):
    """Действующее значение свойства абзаца: прямое форматирование, затем цепочка стилей."""
    v = getattr(par.paragraph_format, attr)
    if v is not None:
        return v
    for st in _style_chain(par.style):
        v = getattr(st.paragraph_format, attr)
        if v is not None:
            return v
    return None


def eff_bold(run, par, attr="bold"):
    v = getattr(run.font, attr)
    if v is not None:
        return v
    for st in _style_chain(par.style):
        v = getattr(st.font, attr)
        if v is not None:
            return v
    return False


def all_paragraphs(doc):
    """(абзац, в_таблице) в порядке документа."""
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            yield Paragraph(child, doc), False
        elif tag == "tbl":
            for row in Table(child, doc).rows:
                seen = set()
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    for p in c.paragraphs:
                        yield p, True


def format_checks(path, pg_start, kind):
    """kind: "report" | "app". Возвращает список проблем оформления."""
    doc = Document(path)
    probs = []
    sec = doc.sections[0]
    got = tuple(round(x.mm, 1) for x in (sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin))
    if got != (30.0, 15.0, 20.0, 20.0):
        probs.append(f"поля {got}, нужно (30, 15, 20, 20)")
    if len(doc.sections) != 1:
        probs.append(f"разделов {len(doc.sections)}")
    pg = sec._sectPr.find(W + "pgNumType")
    start = pg.get(W + "start") if pg is not None else None
    if (start or None) != (str(pg_start) if pg_start else None):
        probs.append(f"нумерация страниц начинается с {start}, нужно {pg_start}")
    foot = sec.footer.paragraphs[0]
    if foot.alignment != WA.CENTER or "PAGE" not in foot._p.xml:
        probs.append("номер страницы не внизу по центру")
    if sec.different_first_page_header_footer:
        probs.append("первая страница без номера")
    cnt = collections.Counter()
    for par, in_tbl in all_paragraphs(doc):
        text = par.text.strip()
        style = par.style.name
        has_pic = "<w:drawing" in par._p.xml
        line = eff(par, "line_spacing")
        if line is None or abs(float(line) - 1.15) > 0.001:
            probs.append(f"интервал {line}: {text[:50]!r}")
        al = eff(par, "alignment")
        ind = eff(par, "first_line_indent")
        ind_cm = round(ind.cm, 2) if ind is not None else 0
        left = eff(par, "left_indent")
        if left is not None and abs(left.cm) > 0.001:
            probs.append(f"отступ слева {left.cm:.2f} см: {text[:50]!r}")
        is_toc = "\t" in par.text and re.search(r"\t\d+$", par.text) is not None
        is_head = style.startswith("Heading")
        is_cap_fig = re.match(r"^Рисунок \d+[А-Я]? – ", text) is not None
        is_cap_tbl = re.match(r"^Таблица \d+[А-Я]? – ", text) is not None
        if in_tbl:
            cnt["ячейки таблиц"] += 1
        elif is_head or has_pic or is_cap_fig:
            cnt["по центру: заголовки, рисунки, подписи рисунков"] += 1
            if al != WA.CENTER:
                probs.append(f"не по центру: {text[:50]!r}")
            if ind_cm != 0:
                probs.append(f"отступ {ind_cm} см у центрированного абзаца: {text[:50]!r}")
        elif is_toc:
            cnt["строки содержания"] += 1
        elif al == WA.CENTER:
            cnt["по центру: заголовки структурных элементов и приложений"] += 1
            if ind_cm != 0:
                probs.append(f"отступ {ind_cm} см у центрированного абзаца: {text[:50]!r}")
        else:
            if al != WA.JUSTIFY:
                probs.append(f"не по ширине ({al}): {text[:50]!r}")
            if is_cap_tbl or not text:
                cnt["по ширине без отступа: подписи таблиц, пустые строки"] += 1
                if ind_cm != 0:
                    probs.append(f"отступ {ind_cm} см: {text[:50]!r}")
            else:
                cnt["по ширине, отступ 1,25 см: текст"] += 1
                if ind_cm != 1.25:
                    probs.append(f"абзацный отступ {ind_cm} см: {text[:50]!r}")
        for run in par.runs:
            if eff_bold(run, par, "bold"):
                probs.append(f"полужирный: {run.text[:40]!r}")
            if eff_bold(run, par, "italic"):
                probs.append(f"курсив: {run.text[:40]!r}")
    # шрифт и кегль каждого прогона, включая ссылки и колонтитул
    z = zipfile.ZipFile(path)
    fonts = collections.Counter()
    for part in [n for n in z.namelist() if re.match(r"word/(document|footer\d*|header\d*)\.xml$", n)]:
        root = etree.fromstring(z.read(part))
        for r in root.iter(W + "r"):
            rpr = r.find(W + "rPr")
            rf = rpr.find(W + "rFonts") if rpr is not None else None
            sz = rpr.find(W + "sz") if rpr is not None else None
            name = rf.get(W + "ascii") if rf is not None else None
            size = int(sz.get(W + "val")) / 2 if sz is not None else None
            fonts[(name, size)] += 1
            if rpr is not None:
                for tag in ("b", "i"):
                    e = rpr.find(W + tag)
                    if e is not None and e.get(W + "val") not in ("0", "false"):
                        probs.append(f"<w:{tag}> в {part}")
    if set(fonts) != {("Times New Roman", 12.0)}:
        probs.append(f"шрифты и кегли прогонов: {dict(fonts)}")
    # стили и значения по умолчанию
    st = etree.fromstring(z.read("word/styles.xml"))
    dflt = st.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}spacing")
    if dflt is None or dflt.get(W + "line") != "276":
        probs.append("интервал по умолчанию не 1,15")
    print(f"   оформление: поля {got}; абзацы: {dict(cnt)}; прогоны: {dict(fonts)}")
    return probs


def headings(doc):
    return [(p.style.name, norm(p.text)) for p in doc.paragraphs if p.style.name.startswith("Heading")]


def report_checks(report, apps):
    probs = []
    doc = Document(report)
    pars = [norm(p.text) for p in doc.paragraphs]
    theme = F.CARD["theme"]
    # тема и работы календарного плана дословно
    for where, upto in (("реферат", "СОДЕРЖАНИЕ"), ("введение", None)):
        if where == "реферат":
            chunk = pars[:pars.index(upto)]
        else:
            a = max(i for i, t in enumerate(pars) if t == "ВВЕДЕНИЕ")
            chunk = pars[a:a + 12]
        if not any(f"«{theme}»" in t for t in chunk):
            probs.append(f"тема не приведена дословно: {where}")
    hs = headings(doc)
    h2 = [t for s, t in hs if s == "Heading 2"]
    want = [f"1.{i + 1} {t}" for i, t in enumerate(F.PLAN)]
    if h2 != want:
        probs.append(f"подразделы не совпадают с работами календарного плана: {h2}")
    h1 = [t for s, t in hs if s == "Heading 1"]
    want1 = ["ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ", "ВВЕДЕНИЕ", "1 ОСНОВНАЯ ЧАСТЬ ОТЧЕТА", "ЗАКЛЮЧЕНИЕ",
             "ПЕРЕЧЕНЬ ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"]
    if h1 != want1:
        probs.append(f"структурные элементы: {h1}")
    if any(s not in ("Heading 1", "Heading 2") for s, _ in hs):
        probs.append("есть заголовки третьего уровня")
    # содержание = заголовки в тексте
    toc = []
    for p in doc.paragraphs:
        m = re.match(r"^(.*)\t(\d+)$", p.text)
        if m:
            toc.append((norm(m.group(1)), int(m.group(2))))
    if [t for t, _ in toc] != [t for _, t in hs]:
        probs.append(f"содержание не совпадает с заголовками: {[t for t, _ in toc]}")
    # короткие абзацы-подзаголовки вне содержания (например, «Цель работы»)
    for p in doc.paragraphs:
        t = norm(p.text)
        if (t and len(t) < 60 and not p.style.name.startswith("Heading") and not t.endswith((".", ";", ":", ")"))
                and "\t" not in p.text and t not in ("РЕФЕРАТ", "СОДЕРЖАНИЕ") and not t.startswith("Отчет ")
                and not re.match(r"^(Таблица|Рисунок) \d", t)):
            probs.append(f"возможный подзаголовок вне содержания: {t!r}")
    # страницы по PDF
    pdf = report[:-5] + ".pdf"
    with pdfplumber.open(pdf) as d:
        pages = [(p.extract_text() or "") for p in d.pages]
    n = len(pages)
    if n + 1 > 20:
        probs.append(f"объем {n + 1} с. (с титульным листом) больше 20")
    for i, t in enumerate(pages):
        last = t.strip().splitlines()[-1].strip() if t.strip() else ""
        if last != str(i + 2):
            probs.append(f"страница файла {i + 1}: номер внизу {last!r}, ожидается {i + 2}")
    for text, page in toc:
        key = text[:38]
        found = None
        for i, t in enumerate(pages):
            if t.lstrip().startswith("СОДЕРЖАНИЕ"):
                continue
            if any(norm(ln) == text or (len(key) >= 38 and norm(ln).startswith(key)) for ln in t.splitlines()):
                found = i + 2
                break
        if found != page:
            probs.append(f"содержание: «{text[:40]}» — страница {page}, в PDF {found}")
    m = re.match(r"^Отчет (\d+) с\.", next(t for t in pars if t.startswith("Отчет ")))
    if not m or int(m.group(1)) != n + 1:
        probs.append(f"реферат: число страниц {m.group(1) if m else None}, фактически {n + 1}")
    # кликабельная ссылка на сайт в подразделе о создании сайта
    rels = {r.rId: r.target_ref for r in doc.part.rels.values() if "hyperlink" in r.reltype}
    site = F.CARD["site_idna"]
    ok = False
    sec = None
    for p in doc.paragraphs:
        if p.style.name == "Heading 2":
            sec = norm(p.text)
        for h in p._p.iter(W + "hyperlink"):
            target = rels.get(h.get(R_NS + "id"), "")
            shown = "".join(t.text or "" for t in h.iter(W + "t"))
            if site in target and sec and sec.startswith("1.4") and F.CARD["site"] in shown:
                ok = True
    if not ok:
        probs.append("нет кликабельной ссылки на сайт в подразделе 1.4")
    # ссылки на объекты приложений
    adoc = Document(apps)
    caps = []
    for p in adoc.paragraphs:
        m = re.match(r"^(Таблица|Рисунок) (\d+)([А-Я]) – ", norm(p.text))
        if m:
            caps.append(f"{m.group(1)} {m.group(2)}{m.group(3)}")
    if sorted(caps) != sorted(apx.all_labels()):
        probs.append(f"объекты приложений не совпадают с реестром: {sorted(set(caps) ^ set(apx.all_labels()))}")
    body = " ".join(pars)
    for kind, tbl in doc_tables_text(doc):
        body += " " + tbl
    refs = set()
    for word, stem in (("Таблица", r"таблиц(?:а|е|у|ы|и|ах)"), ("Рисунок", r"рисун(?:ок|ке|ка|ки|ках)")):
        for m in re.finditer(stem + r"\s+(\d+[А-Я](?:(?:\s*[,–]\s*|\s+и\s+)\d+[А-Я])*)", body, flags=re.IGNORECASE):
            parts = re.findall(r"\d+[А-Я]|–", m.group(1))
            prev = None
            rng = False
            for part in parts:
                if part == "–":
                    rng = True
                    continue
                num, let = int(part[:-1]), part[-1]
                if rng and prev and prev[1] == let:
                    for k in range(prev[0] + 1, num + 1):
                        refs.add(f"{word} {k}{let}")
                refs.add(f"{word} {num}{let}")
                prev, rng = (num, let), False
    for c in caps:
        if c not in refs:
            probs.append(f"{c}: нет ссылки в отчете")
    for r in sorted(refs):
        if r not in caps:
            probs.append(f"{r}: ссылка в отчете на несуществующий объект приложений")
    letters = []
    for m in re.finditer(r"[Пп]риложени(?:и|е|я|ях|ю)\s+([А-Я])(?![А-Яа-яё])", body):
        if m.group(1) not in letters:
            letters.append(m.group(1))
    if letters != [a[0] for a in apx.APPS]:
        probs.append(f"приложения упоминаются в порядке {letters}")
    # приложения: названия = работы календарного плана, тема дословно, нумерация объектов
    apars = [norm(p.text) for p in adoc.paragraphs]
    if not any(f"«{theme}»" in t for t in apars[:6]):
        probs.append("приложения: тема не приведена дословно")
    for letter, title, objs in apx.APPS:
        try:
            i = apars.index(f"ПРИЛОЖЕНИЕ {letter}")
        except ValueError:
            probs.append(f"нет заголовка ПРИЛОЖЕНИЕ {letter}")
            continue
        if apars[i + 1] != title:
            probs.append(f"приложение {letter}: название {apars[i + 1]!r} не совпадает с работой календарного плана")
        for word in ("Таблица", "Рисунок"):
            nums = [int(c.split()[1][:-1]) for c in caps if c.startswith(word) and c.endswith(letter)]
            if nums != list(range(1, len(nums) + 1)):
                probs.append(f"приложение {letter}: нумерация «{word}» {nums}")
    return probs


def doc_tables_text(doc):
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                yield "c", norm(c.text)


def layout_report(report):
    """Положение заголовков на странице и заполнение последних страниц разделов (по PDF)."""
    pdf = report[:-5] + ".pdf"
    out = []
    with pdfplumber.open(pdf) as d:
        for i, page in enumerate(d.pages):
            lines = page.extract_text_lines()
            body = [ln for ln in lines if ln["top"] < page.height - 70]
            if not body:
                continue
            fill = (max(ln["bottom"] for ln in body) - 56.7) / (page.height - 2 * 56.7)
            heads = []
            for ln in body:
                t = norm(ln["text"])
                if re.match(r"^1\.\d ", t) and i > 1:
                    pos = (ln["top"] - 56.7) / (page.height - 2 * 56.7)
                    heads.append(f"{t[:22]}… на {pos:.0%} высоты")
            out.append((i + 2, len(body), fill, heads))
    return out


def run_report(report, apps):
    total = []
    for path, kind, pg in ((report, "report", 2), (apps, "app", None)):
        probs = check(path, n_sources=7 if kind == "report" else None, lettered=(kind == "app"),
                      ordered_sources=True)
        probs += format_checks(path, pg, kind)
        if kind == "report":
            probs += report_checks(report, apps)
        for p in probs:
            print("   ПРОБЛЕМА:", p)
        if not probs:
            print("   проблем нет")
        total += probs
    print("== раскладка отчета по страницам (страница: строк, заполнение, заголовки подразделов)")
    for pg, n, fill, heads in layout_report(report):
        print(f"   {pg:>2}: {n:>2} строк, {fill:>4.0%}" + ("  " + "; ".join(heads) if heads else ""))
    return total


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "--report":
        sys.exit(1 if run_report(args[1], args[2]) else 0)
    bad = False
    for a in args:
        path, _, ns = a.partition("::")
        probs = check(path, int(ns) if ns else None, allow_url="развитии" in path)
        for p in probs:
            print("   ПРОБЛЕМА:", p)
        if not probs:
            print("   проблем нет")
        bad = bad or bool(probs)
    sys.exit(1 if bad else 0)
