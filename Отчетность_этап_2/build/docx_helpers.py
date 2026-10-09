# -*- coding: utf-8 -*-
"""Общие помощники для сборки docx-документов грантовой отчётности (ФСИ).

Форматирование по регламенту: Times New Roman 14 пт, интервал 1,5, абзацный отступ 1,25 см,
выравнивание по ширине; поля 30/15/20/20 мм; подписи таблиц слева сверху «Таблица N – …»,
рисунков по центру снизу «Рисунок N – …»; сквозная нумерация.
"""
import os
import re
import copy
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

NBSP = " "
DASH = "–"  # короткое тире
FONT = "Times New Roman"
TEXT_WIDTH_MM = 165

FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def nb(s):
    """Неразрывные пробелы в числах вида '13 362' и перед единицами."""
    return s.replace(" ", NBSP)


def fmt_num(x, dec=0):
    """Число по-русски: разряды через неразрывный пробел, десятичная запятая."""
    if dec == 0:
        s = f"{int(round(x)):,}".replace(",", NBSP)
    else:
        s = f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", NBSP)
    return s


def _set_run_font(run, size=14, bold=None, italic=None, name=FONT, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=Cm(1.25), spacing=1.5,
              before=0, after=0, left=None, keep_next=False, page_break_before=False):
    pf = p.paragraph_format
    pf.alignment = align
    pf.first_line_indent = first_line
    if left is not None:
        pf.left_indent = left
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = spacing
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_next
    pf.page_break_before = page_break_before
    pf.widow_control = True


def new_document():
    doc = Document()
    # --- стили по умолчанию
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(14)
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # docDefaults (на случай стилей без явного шрифта)
    styles_el = doc.styles.element
    rpr_default = styles_el.find(qn("w:docDefaults"))
    if rpr_default is not None:
        rpd = rpr_default.find(qn("w:rPrDefault"))
        if rpd is not None:
            r = rpd.find(qn("w:rPr"))
            if r is not None:
                rf = r.find(qn("w:rFonts"))
                if rf is None:
                    rf = OxmlElement("w:rFonts")
                    r.insert(0, rf)
                for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                    rf.set(qn(attr), FONT)
                for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
                    if rf.get(qn(a)) is not None:
                        del rf.attrib[qn(a)]
                sz = r.find(qn("w:sz"))
                if sz is None:
                    sz = OxmlElement("w:sz")
                    r.append(sz)
                sz.set(qn("w:val"), "28")
                szcs = r.find(qn("w:szCs"))
                if szcs is None:
                    szcs = OxmlElement("w:szCs")
                    r.append(szcs)
                szcs.set(qn("w:val"), "28")
    # --- страница
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.left_margin = Mm(30)
    sec.right_margin = Mm(15)
    sec.top_margin = Mm(20)
    sec.bottom_margin = Mm(20)
    sec.header_distance = Mm(10)
    sec.footer_distance = Mm(10)
    # --- нумерация страниц в нижнем колонтитуле по центру; титул без номера
    sec.different_first_page_header_footer = True
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.first_line_indent = Cm(0)
    run = fp.add_run()
    _set_run_font(run, size=12)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    rpr2 = OxmlElement("w:rPr")
    rf2 = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf2.set(qn(attr), FONT)
    rpr2.append(rf2)
    sz2 = OxmlElement("w:sz"); sz2.set(qn("w:val"), "24"); rpr2.append(sz2)
    r.append(rpr2)
    t = OxmlElement("w:t"); t.text = "2"; r.append(t)
    fld.append(r)
    run._element.addnext(fld)
    # pgNumType start=1
    sectPr = sec._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectPr.append(pg)
    pg.set(qn("w:start"), "1")
    # удалить пустой первый абзац, если есть
    return doc


def remove_leading_empty_paragraph(doc):
    body = doc.element.body
    ps = body.findall(qn("w:p"))
    if ps and not ps[0].findall(".//" + qn("w:t")) and not ps[0].findall(".//" + qn("w:drawing")):
        body.remove(ps[0])


# ----------------------------------------------------------------- блоки текста

def para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, size=14,
         first_line=Cm(1.25), spacing=1.5, before=0, after=0, left=None, keep_next=False,
         page_break_before=False, color=None):
    """Абзац с поддержкой **жирного** и *курсива* внутри текста (простая разметка)."""
    p = doc.add_paragraph()
    _para_fmt(p, align=align, first_line=first_line, spacing=spacing, before=before, after=after,
              left=left, keep_next=keep_next, page_break_before=page_break_before)
    add_marked_runs(p, text, size=size, bold=bold, italic=italic, color=color)
    return p


_MARK_RE = re.compile(r"(\*\*.+?\*\*|\*.+?\*)")
TEXT_FILTER = None  # функция-резолвер ссылок {T:key}/{F:key}; задаётся сборочным скриптом


def add_marked_runs(p, text, size=14, bold=False, italic=False, color=None):
    if TEXT_FILTER is not None:
        text = TEXT_FILTER(text)
    parts = _MARK_RE.split(text)
    for part in parts:
        if not part:
            continue
        b, i = bold, italic
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part = part[2:-2]
            b = True
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            part = part[1:-1]
            i = True
        run = p.add_run(part)
        _set_run_font(run, size=size, bold=b, italic=i, color=color)


def heading_struct(doc, text, page_break=True):
    """Заголовок структурного элемента: прописными, по центру, жирный, с новой страницы."""
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=Cm(0), keep_next=True,
              page_break_before=page_break, after=0)
    run = p.add_run(text.upper())
    _set_run_font(run, size=14, bold=True)
    p._heading_level = 1
    return p


def heading_sub(doc, text, page_break=False, before=0, after=0):
    """Заголовок подраздела (1.1 …): жирный, с абзацного отступа, без точки в конце."""
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=Cm(1.25), keep_next=True,
              page_break_before=page_break, before=before, after=after)
    run = p.add_run(text)
    _set_run_font(run, size=14, bold=True)
    p._heading_level = 2
    return p


def heading_minor(doc, text):
    """Подзаголовок внутри подраздела (без номера): жирный, с абзацного отступа."""
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line=Cm(1.25), keep_next=True)
    run = p.add_run(text)
    _set_run_font(run, size=14, bold=True)
    return p


def bullets(doc, items, marker=DASH, end=";", last_end="."):
    """Перечисление: «– текст;» с абзацного отступа."""
    n = len(items)
    for i, it in enumerate(items):
        suffix = last_end if i == n - 1 else end
        txt = it.rstrip(".;")
        para(doc, f"{marker} {txt}{suffix}")


def numbered(doc, items, end=";", last_end="."):
    n = len(items)
    for i, it in enumerate(items):
        suffix = last_end if i == n - 1 else end
        txt = it.rstrip(".;")
        para(doc, f"{i + 1}) {txt}{suffix}")


def page_break(doc):
    p = doc.add_paragraph()
    _para_fmt(p, first_line=Cm(0))
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


# ----------------------------------------------------------------- таблицы

def _set_cell_margins(table, top=40, start=60, bottom=40, end=60):
    tbl = table._tbl
    tblPr = tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _set_col_widths(table, widths_mm):
    tbl = table._tbl
    tblPr = tbl.tblPr
    # фиксированная раскладка
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths_mm):
            gc.set(qn("w:w"), str(int(Mm(w).twips)))
    for row in table.rows:
        for idx, w in enumerate(widths_mm):
            if idx < len(row.cells):
                row.cells[idx].width = Mm(w)


def _cell_text(cell, text, size=12, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    cell.text = ""
    p = cell.paragraphs[0]
    _para_fmt(p, align=align, first_line=Cm(0), spacing=1.0)
    lines = str(text).split("\n")
    for li, line in enumerate(lines):
        if li > 0:
            p = cell.add_paragraph()
            _para_fmt(p, align=align, first_line=Cm(0), spacing=1.0)
        add_marked_runs(p, line, size=size, bold=bold, italic=italic)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table(doc, number, title, header, rows, widths_mm, font_size=12, header_size=12,
          align_cols=None, header_fill="E7E6E6", caption=True, merge_rows=None):
    """Таблица с подписью «Таблица N – Название» слева над таблицей.

    merge_rows: список индексов строк (в rows) — строка-разделитель на всю ширину (объединённая ячейка).
    """
    if caption:
        cp = doc.add_paragraph()
        _para_fmt(cp, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=Cm(0), keep_next=True, before=6)
        run = cp.add_run(f"Таблица {number} {DASH} {title}")
        _set_run_font(run, size=14)
    ncols = len(header)
    t = doc.add_table(rows=1, cols=ncols)
    t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_cell_margins(t)
    align_cols = align_cols or [WD_ALIGN_PARAGRAPH.LEFT] * ncols
    # шапка
    hdr = t.rows[0]
    trPr = hdr._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    cs = OxmlElement("w:cantSplit"); trPr.append(cs)
    for i, h in enumerate(header):
        c = hdr.cells[i]
        _cell_text(c, h, size=header_size, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        if header_fill:
            tcPr = c._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), header_fill)
            tcPr.append(shd)
    merge_rows = set(merge_rows or [])
    for ri, row in enumerate(rows):
        r = t.add_row()
        trPr = r._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        if ri in merge_rows:
            merged = r.cells[0].merge(r.cells[ncols - 1])
            _cell_text(merged, row[0], size=font_size, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
            tcPr = merged._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F2F2F2")
            tcPr.append(shd)
            continue
        for ci in range(ncols):
            val = row[ci] if ci < len(row) else ""
            _cell_text(r.cells[ci], val, size=font_size, align=align_cols[ci])
    _set_col_widths(t, widths_mm)
    # отступ после таблицы
    sp = doc.add_paragraph()
    _para_fmt(sp, first_line=Cm(0), spacing=1.0, after=6)
    return t


# ----------------------------------------------------------------- рисунки

def figure(doc, number, title, image_path, width_mm=TEXT_WIDTH_MM):
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=Cm(0), keep_next=True, before=6)
    p.add_run().add_picture(image_path, width=Mm(width_mm))
    cp = doc.add_paragraph()
    _para_fmt(cp, align=WD_ALIGN_PARAGRAPH.CENTER, first_line=Cm(0), after=6)
    run = cp.add_run(f"Рисунок {number} {DASH} {title}")
    _set_run_font(run, size=14)
    return p


def make_placeholder(path, lines, width_px=1650, height_px=950, label="МЕСТО ДЛЯ ФОТО / ВИДЕОКАДРА"):
    """PNG-заглушка под фотографию: серый прямоугольник с подписью, что сюда вставить."""
    img = Image.new("RGB", (width_px, height_px), (240, 240, 240))
    d = ImageDraw.Draw(img)
    d.rectangle([4, 4, width_px - 5, height_px - 5], outline=(120, 120, 120), width=4)
    try:
        f1 = ImageFont.truetype(FONT_BOLD_PATH, 54)
        f2 = ImageFont.truetype(FONT_PATH, 40)
    except Exception:
        f1 = f2 = ImageFont.load_default()
    y = height_px // 2 - 60 - 28 * len(lines)
    w = d.textlength(label, font=f1)
    d.text(((width_px - w) / 2, y), label, fill=(90, 90, 90), font=f1)
    y += 110
    for ln in lines:
        w = d.textlength(ln, font=f2)
        d.text(((width_px - w) / 2, y), ln, fill=(110, 110, 110), font=f2)
        y += 56
    img.save(path)
    return path


def figure_placeholder(doc, number, title, lines, out_dir, width_mm=140, height_px=700, key=None):
    os.makedirs(out_dir, exist_ok=True)
    key = key or f"ph_{number}"
    path = os.path.join(out_dir, f"{key}.png")
    make_placeholder(path, lines, height_px=height_px)
    return figure(doc, number, title, path, width_mm=width_mm)


# ----------------------------------------------------------------- содержание

def toc_entry(doc, text, page, level=1):
    p = doc.add_paragraph()
    _para_fmt(p, align=WD_ALIGN_PARAGRAPH.LEFT, first_line=Cm(0), spacing=1.5,
              left=Cm(0.75 if level == 2 else 0))
    tabs = p.paragraph_format.tab_stops
    tabs.add_tab_stop(Mm(TEXT_WIDTH_MM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    run = p.add_run(f"{text}\t{page}")
    _set_run_font(run, size=14)
    return p


# ----------------------------------------------------------------- титульный лист

def title_block(doc, lines):
    """lines: список (text, align, bold, size, before_pt)."""
    for text, align, bold, size, before in lines:
        p = doc.add_paragraph()
        _para_fmt(p, align=align, first_line=Cm(0), spacing=1.15, before=before)
        run = p.add_run(text)
        _set_run_font(run, size=size, bold=bold)


# ----------------------------------------------------------------- проверка

FORBIDDEN = [
    r"\bисследован", r"\bисследовани", r"\bНИР\b", r"\bНИОКР\b", r"\bTRL\b", r"\bУГТ\b",
    r"https?://", r"www\.", r"\bмы\b", r"\bнами\b", r"\bнаш[аеиу]?\b", r"\bя\b", r"уровн\w* готовности",
]


def check_text(paragraph_texts):
    problems = []
    for i, t in enumerate(paragraph_texts):
        for pat in FORBIDDEN:
            for m in re.finditer(pat, t, flags=re.IGNORECASE):
                problems.append((i, pat, t[max(0, m.start() - 40): m.end() + 40]))
    return problems
