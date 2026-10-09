#!/usr/bin/env python3
"""Схемы fig1–fig6 для официального отчёта (подписи на русском языке).

Стиль: белый фон, чёрно-серая гамма, тёмно-синий #1F3A5F для заголовочных полос,
светло-серый #F2F2F2 для блоков, чёрные рамки 1 px и чёрные стрелки, без теней.
Ширина 165 мм (6.5 дюйма) при 300 dpi. Координаты заданы в дюймах, поэтому геометрия
и кегль не искажаются. Запуск из папки со скриптом: python3 -I make_diagrams.py
"""

import os
import struct
import textwrap

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.font_manager import FontProperties  # noqa: E402
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle  # noqa: E402
from matplotlib.path import Path  # noqa: E402

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_REG = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

DPI = 300
FIG_W = 6.5                   # 165 мм
M = 0.05                      # поле внутри рисунка, дюймы
X0, X1 = M, FIG_W - M         # рабочая ширина схем, 6.4 дюйма
EPS = 1e-3

NAVY = "#1F3A5F"
LIGHT = "#F2F2F2"
BLACK = "#000000"
WHITE = "#FFFFFF"
NOTE = "#4D4D4D"              # серый курсив выносок
LEADER = "#7F7F7F"            # пунктирные выноски

FRAME_LW = 72.0 / DPI         # рамка 1 px, в пунктах
HILITE_LW = 2 * 72.0 / DPI    # выделение ККТ 2 px
ARROW_LW = 0.8

RAD = 0.05                    # радиус скругления углов, дюймы
LS = 1.3                      # межстрочный интервал, доли кегля
PAD_X = 0.05                  # отступ текста от рамки по горизонтали
PAD_Y = 0.06                  # отступ тела блока по вертикали
HEAD_PAD = 0.05               # отступ заголовочной полосы
TITLE_PT = 9                  # заголовки блоков (жирный)
BODY_PT = 8                   # текст блоков (не менее 8 pt)

Z_BLOCK, Z_LEAD, Z_ARROW, Z_TEXT = 3, 5, 6, 7
WRAP_LOG = []                 # (фрагмент, число символов в строке переноса)

# ----------------------------- тексты (дословно) -----------------------------

FIG1_STEPS = [
    ("1. Разработка состава",
     "10 вариантов состава; изомальт трёх производителей; расчёт рецептур на тираж"),
    ("2. Создание образцов",
     "3 партии образцов; силиконовые формы Ø12 и Ø30 см; контроль температуры"),
    ("3. Тестирование образцов",
     "дегустация (фокус-группа 5 чел.); центральное отверстие; хранение"),
    ("4. Сайт проекта",
     "вкусноты.рф; 3D-модели пластинок; форма заказа"),
    ("5. Рекламная кампания",
     "промо-продукт «Синестезия»; Яндекс Директ; наружная реклама; релиз с артисткой"),
]
FIG1_RESULT = ("Результат этапа: состав выбран, образцы изготовлены и проверены, "
               "сайт и промо-инструмент запущены, рекламная кампания проведена")

FIG2_BLOCKS = [
    "10 вариантов состава (сахарная база, изомальт + эритрит, изомальт + патока, "
    "+ мальтодекстрин, + гуммиарабик, + пектин, + какао-масло, «продуктовый», "
    "комбинированные варианты)",
    "Дегустация: оценка вкуса, сладости, посторонних привкусов",
    "Формование и выполнение центрального отверстия: растрескивание, сколы, липкость",
    "Хранение в упаковке: липкость поверхности, выделение влаги",
]
FIG2_LAST = "Итоговый состав: изомальт + эритрит"
FIG2_NOTES = {
    1: "отклонены варианты с невыраженным или посторонним вкусом",
    2: "отклонены сахарно-глюкозная база и варианты с чёрным красителем "
       "(растрескивание, липкость)",
    3: "отклонены варианты, выделяющие влагу при хранении",
}

FIG3_OPS = [
    "Дозирование компонентов (весы 0,01 г)",
    "Сухое смешивание изомальта и эритрита",
    "Нагрев до 165–168 °C, контроль термометром и пирометром",
    "Охлаждение массы до 140 °C",
    "Ввод лимонной кислоты, ароматизатора, красителя",
    "Разлив в силиконовые формы Ø12 / Ø30 см (3–4 мин)",
    "Застывание 10–15 мин, извлечение из форм",
    "Выполнение центрального отверстия",
    "Контроль: геометрия, поверхность, целостность",
    "Упаковка в герметичный пакет (40 мкм)",
    "Маркировка образца (состав, цвет, вкус, дата)",
]
FIG3_KKT = (2, 4)             # блоки 3 и 5 (нумерация с нуля)
CELL3 = {0: (0, 0), 1: (0, 1), 2: (0, 2), 3: (0, 3),
         4: (1, 3), 5: (1, 2), 6: (1, 1), 7: (1, 0),
         8: (2, 0), 9: (2, 1), 10: (2, 2)}          # змейка: (ряд, колонка)

FIG5_STEPS = [
    ("Ввод:", "артист и название трека"),
    ("Проверка запроса:",
     "опечатки, перестановка полей, несуществующий трек («did_you_mean» / «not_found»)"),
    ("Сбор сигналов:",
     "текст песни, метаданные и обложка из iTunes, визуальный анализ обложки"),
    ("Языковая модель (через OpenRouter):", "формирование сенсорной карты"),
    ("Результат:", "палитра 4–5 цветов с названиями, вкус с нотами, пирамида аромата, разбор"),
    ("Карточка:", "короткая ссылка, PNG для сторис, OG-превью; кнопка перехода на вкусноты.рф"),
]
FIG5_QUOTA = ("Квоты: 10 анализов в месяц бесплатно, +3 за подписку на Telegram-канал, "
              "+2 за приглашённого друга; кэш повторных анализов")
CELL5 = {0: (0, 0), 1: (0, 1), 2: (0, 2), 3: (1, 2), 4: (1, 1), 5: (1, 0)}

# ------------------------------- шрифты и измерения -------------------------------


def font(kind, size):
    if kind == "bold":
        fp = FontProperties(fname=FONT_BOLD)
    elif kind == "ital":
        fp = FontProperties(family="DejaVu Sans", style="italic")  # Oblique из комплекта matplotlib
    else:
        fp = FontProperties(fname=FONT_REG)
    fp.set_size(size)
    return fp


_MEASURE_FIG = plt.figure(dpi=DPI)
_RENDERER = _MEASURE_FIG.canvas.get_renderer()


def text_w(s, kind, size):
    """Ширина строки в дюймах."""
    return _RENDERER.get_text_width_height_descent(s, font(kind, size), ismath=False)[0] / DPI


def pitch(size):
    """Межстрочный шаг в дюймах."""
    return size * LS / 72.0


GLUE = [(" °C", "\u00a0°C"), (" мин", "\u00a0мин"), (" см", "\u00a0см"), (" мкм", "\u00a0мкм"),
        (" чел.", "\u00a0чел."), (" руб.", "\u00a0руб."), (" г)", "\u00a0г)")]


def wrap_text(text, n):
    """textwrap; неразрывный пробел перед единицей только для переноса, в результате — обычный пробел."""
    s = text
    for a, b in GLUE:
        s = s.replace(a, b)
    return [ln.replace("\u00a0", " ")
            for ln in textwrap.wrap(s, width=n, break_long_words=False, break_on_hyphens=True)]


def fit(text, avail, kind, size, hi=26, tag=""):
    """textwrap с наибольшей шириной n (от hi вниз), при которой все строки помещаются в avail."""
    for n in range(hi, 4, -1):
        lines = wrap_text(text, n)
        if lines and max(text_w(s, kind, size) for s in lines) <= avail:
            WRAP_LOG.append((tag or text[:24], n))
            return lines
    raise ValueError(f"Текст не помещается в блок: {text!r}, доступно {avail:.3f} дюйма")


# --------------------------------- блоки ---------------------------------


class Card:
    """Блок: заголовочная полоса (тёмно-синяя, белый жирный) и/или тело (светло-серое)."""

    def __init__(self, title, body, avail, *, tsize=TITLE_PT, bsize=BODY_PT, hi=26, tag=""):
        self.tsize = tsize
        self.bsize = bsize
        self.tlines = fit(title, avail, "bold", tsize, hi, tag) if title else []
        self.blines = fit(body, avail, "reg", bsize, hi, tag) if body else []
        self.t_fixed = None
        self.b_fixed = None

    def t_nat(self):
        return len(self.tlines) * pitch(self.tsize) + 2 * HEAD_PAD if self.tlines else 0.0

    def b_nat(self):
        return len(self.blines) * pitch(self.bsize) + 2 * PAD_Y if self.blines else 0.0

    def th(self):
        if not self.tlines:
            return 0.0
        return self.t_fixed if self.t_fixed is not None else self.t_nat()

    def bh(self):
        if not self.blines:
            return 0.0
        return self.b_fixed if self.b_fixed is not None else self.b_nat()

    def height(self):
        if self.tlines and self.blines:
            return self.th() + self.bh()
        return self.th() if self.tlines else self.bh()


def equalize(cards):
    """Одинаковые полосы заголовков и области текста у блоков одного ряда."""
    both = [c for c in cards if c.tlines and c.blines]
    if both:
        t_max = max(c.t_nat() for c in both)
        b_max = max(c.b_nat() for c in both)
        for c in both:
            c.t_fixed, c.b_fixed = t_max, b_max


def shrink(r, d):
    return (r[0] + d, r[1] + d, r[2] - d, r[3] - d)


def grow(r, d):
    return (r[0] - d, r[1] - d, r[2] + d, r[3] + d)


def overlap(a, b):
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def seg_hits(a, b, r):
    """Пересекает ли отрезок ab прямоугольник r (алгоритм Лианга–Барски)."""
    x0, y0, x1, y1 = r
    dx, dy = b[0] - a[0], b[1] - a[1]
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, a[0] - x0), (dx, x1 - a[0]), (-dy, a[1] - y0), (dy, y1 - a[1])):
        if p == 0:
            if q < 0:
                return False
            continue
        t = q / p
        if p < 0:
            if t > t1:
                return False
            t0 = max(t0, t)
        else:
            if t < t0:
                return False
            t1 = min(t1, t)
    return True


class Diagram:
    """Холст в дюймах: y отсчитывается сверху вниз (top), внутри matplotlib — снизу вверх."""

    def __init__(self, H):
        self.H = H
        self.fig = plt.figure(figsize=(FIG_W, H), dpi=DPI)
        self.fig.patch.set_facecolor(WHITE)
        # оси занимают ровно рабочую область: bbox_inches='tight' тогда даёт 6.5 дюйма по ширине
        self.ax = self.fig.add_axes([X0 / FIG_W, M / H, (X1 - X0) / FIG_W, (H - 2 * M) / H])
        self.ax.set_xlim(X0, X1)
        self.ax.set_ylim(M, H - M)
        self.ax.axis("off")
        self.blocks = []      # контуры блоков
        self.texts = []       # все подписи
        self.inside = []      # (подпись, блок): подпись должна быть внутри блока
        self.paths = []       # стрелки и выноски (точки в дюймах, сверху вниз)

    def put_lines(self, lines, cx, first_top, p, size, kind, color, ha="center", container=None):
        for i, ln in enumerate(lines):
            t = self.ax.text(cx, self.H - (first_top + i * p), ln, fontproperties=font(kind, size),
                             color=color, ha=ha, va="center", zorder=Z_TEXT)
            self.texts.append(t)
            if container is not None:
                self.inside.append((t, container))

    def center_text(self, lines, x, cy, size, kind="reg", color=BLACK, ha="center"):
        p = pitch(size)
        self.put_lines(lines, x, cy - (len(lines) - 1) * p / 2, p, size, kind, color, ha=ha)

    def block(self, card, x, top, w, h=None, *, edge=BLACK, edge_lw=FRAME_LW):
        h = card.height() if h is None else h
        H = self.H
        bottom = H - top - h
        outer = FancyBboxPatch((x, bottom), w, h, boxstyle=f"round,pad=0,rounding_size={RAD}",
                               fc=LIGHT, ec="none", zorder=Z_BLOCK)
        self.ax.add_patch(outer)
        self.blocks.append(outer)
        cx = x + w / 2
        if card.tlines:
            sh = card.th() if card.blines else h
            strip = Rectangle((x, H - top - sh), w, sh, fc=NAVY, ec="none", zorder=Z_BLOCK + 0.1)
            self.ax.add_patch(strip)
            strip.set_clip_path(outer)
            p = pitch(card.tsize)
            first = top + (sh - len(card.tlines) * p) / 2 + p / 2
            self.put_lines(card.tlines, cx, first, p, card.tsize, "bold", WHITE, container=outer)
        if card.blines:
            btop = top + (card.th() if card.tlines else 0.0)
            area = h - (btop - top)
            p = pitch(card.bsize)
            first = btop + (area - len(card.blines) * p) / 2 + p / 2
            self.put_lines(card.blines, cx, first, p, card.bsize, "reg", BLACK, container=outer)
        self.ax.add_patch(FancyBboxPatch((x, bottom), w, h, boxstyle=f"round,pad=0,rounding_size={RAD}",
                                         fc="none", ec=edge, lw=edge_lw, zorder=Z_BLOCK + 0.3))
        return h

    def arrow(self, pts):
        P = [(x, self.H - t) for x, t in pts]
        self.paths.append(P)
        path = Path(P, [Path.MOVETO] + [Path.LINETO] * (len(P) - 1))
        self.ax.add_patch(FancyArrowPatch(path=path, arrowstyle="-|>", mutation_scale=11,
                                          color=BLACK, lw=ARROW_LW, shrinkA=0, shrinkB=0,
                                          joinstyle="miter", zorder=Z_ARROW))

    def leader(self, pts):
        P = [(x, self.H - t) for x, t in pts]
        self.paths.append(P)
        self.ax.plot([p[0] for p in P], [p[1] for p in P], color=LEADER, lw=0.6,
                     linestyle=(0, (2, 2)), solid_capstyle="butt", zorder=Z_LEAD)

    def frame(self):
        """Невидимый прямоугольник на всю рабочую ширину: задаёт ровно 165 мм по ширине."""
        self.ax.add_patch(Rectangle((X0, M), X1 - X0, self.H - 2 * M, fc="none", ec="none", zorder=0))

    def check(self, name):
        R = self.fig.canvas.get_renderer()
        inv = self.ax.transData.inverted()

        def box(a):
            bb = a.get_window_extent(R)
            (x0, y0), (x1, y1) = inv.transform([(bb.x0, bb.y0), (bb.x1, bb.y1)])
            return (x0, y0, x1, y1)

        probs = []
        in_block = {id(t) for t, _ in self.inside}
        for t, blk in self.inside:
            tb, bb = box(t), box(blk)
            if tb[0] < bb[0] - EPS or tb[2] > bb[2] + EPS or tb[1] < bb[1] - EPS or tb[3] > bb[3] + EPS:
                probs.append(f"текст выходит за рамку блока: {t.get_text()!r}")
        blocks = [box(b) for b in self.blocks]
        texts = [box(t) for t in self.texts]
        for i, (t, tb) in enumerate(zip(self.texts, texts)):
            if tb[0] < X0 - EPS or tb[2] > X1 + EPS or tb[1] < M - EPS or tb[3] > self.H - M + EPS:
                probs.append(f"подпись за пределами поля: {t.get_text()!r}")
            if id(t) not in in_block:
                for bb in blocks:
                    if overlap(shrink(tb, 0.005), shrink(bb, 0.005)):
                        probs.append(f"подпись наезжает на блок: {t.get_text()!r}")
            for j in range(i + 1, len(texts)):
                if overlap(shrink(tb, 0.003), shrink(texts[j], 0.003)):
                    probs.append(f"подписи пересекаются: {t.get_text()!r} / {self.texts[j].get_text()!r}")
        for P in self.paths:
            for a, b in zip(P[:-1], P[1:]):
                for bb in blocks:
                    if seg_hits(a, b, shrink(bb, 0.03)):
                        probs.append(f"стрелка заходит в блок: {a} -> {b}")
                for t, tb in zip(self.texts, texts):
                    if seg_hits(a, b, grow(tb, 0.02)):
                        probs.append(f"стрелка пересекает текст: {t.get_text()!r}")
        return probs

    def save(self, name):
        self.frame()
        probs = self.check(name)
        path = os.path.join(OUT_DIR, name)
        self.fig.savefig(path, dpi=DPI, bbox_inches="tight", pad_inches=0.05, facecolor=WHITE)
        plt.close(self.fig)
        with open(path, "rb") as fh:
            w, h = struct.unpack(">II", fh.read(24)[16:24])
        print(f"{name}: {w} x {h} px = {w / DPI * 25.4:.1f} x {h / DPI * 25.4:.1f} мм; "
              f"замечаний: {len(probs)}")
        for p in probs:
            print("   -", p)


def snake(D, cards, cell, tops, rh, cw, gx, hl=()):
    """Блоки в ряды; стрелки по порядку, в конце ряда — переход на следующий ряд."""
    def xl(c):
        return X0 + c * (cw + gx)

    def xr(c):
        return xl(c) + cw

    for i, card in enumerate(cards):
        r, c = cell[i]
        hi = i in hl
        D.block(card, xl(c), tops[r], cw, rh[r], edge=NAVY if hi else BLACK,
                edge_lw=HILITE_LW if hi else FRAME_LW)
    for i in range(len(cards) - 1):
        r1, c1 = cell[i]
        r2, c2 = cell[i + 1]
        if r1 == r2:
            mid = tops[r1] + rh[r1] / 2
            if c2 > c1:
                D.arrow([(xr(c1), mid), (xl(c2), mid)])
            else:
                D.arrow([(xl(c1), mid), (xr(c2), mid)])
        else:
            x = xl(c1) + cw / 2
            D.arrow([(x, tops[r1] + rh[r1]), (x, tops[r2])])
    return xl, xr


# ------------------------------------ схемы ------------------------------------


def fig1():
    n, gap = 5, 0.15
    bw = (X1 - X0 - (n - 1) * gap) / n
    cards = [Card(t, b, bw - 2 * PAD_X, tag=f"fig1 блок {i + 1}") for i, (t, b) in enumerate(FIG1_STEPS)]
    equalize(cards)
    rh = max(c.height() for c in cards)
    strip = Card(None, FIG1_RESULT, X1 - X0 - 2 * PAD_X, bsize=TITLE_PT, hi=120, tag="fig1 результат")
    y_strip = M + rh + 0.3
    H = y_strip + strip.height() + M
    D = Diagram(H)
    for i, c in enumerate(cards):
        D.block(c, X0 + i * (bw + gap), M, bw, rh)
    mid = M + rh / 2
    for i in range(n - 1):
        xr_ = X0 + i * (bw + gap) + bw
        D.arrow([(xr_, mid), (xr_ + gap, mid)])
    D.block(strip, X0, y_strip, X1 - X0)
    D.save("fig1.png")


def fig2():
    xc = 1.55
    widths = [2.7, 2.3, 2.0, 1.8, 1.6]
    cards = [Card(None, FIG2_BLOCKS[i], widths[i] - 2 * PAD_X, tag=f"fig2 блок {i + 1}") for i in range(4)]
    cards.append(Card(FIG2_LAST, None, widths[4] - 2 * PAD_X, tag="fig2 блок 5"))
    heights = [c.height() for c in cards]
    gap = 0.17
    tops, y = [], M
    for h in heights:
        tops.append(y)
        y += h + gap
    H = y - gap + M
    D = Diagram(H)
    call_x = 4.35
    call_w = X1 - call_x
    for i, c in enumerate(cards):
        D.block(c, xc - widths[i] / 2, tops[i], widths[i], heights[i])
    for i in range(4):
        D.arrow([(xc, tops[i] + heights[i]), (xc, tops[i + 1])])
    for i, text in FIG2_NOTES.items():
        mid = tops[i] + heights[i] / 2
        D.leader([(xc + widths[i] / 2 + 0.06, mid), (call_x - 0.06, mid)])
        lines = fit(text, call_w - 0.02, "ital", BODY_PT, hi=34, tag=f"fig2 выноска {i + 1}")
        D.center_text(lines, call_x, mid, BODY_PT, "ital", NOTE, ha="left")
    D.save("fig2.png")


def fig3():
    ncol, gx = 4, 0.28
    cw = (X1 - X0 - (ncol - 1) * gx) / ncol
    cards = [Card(str(i + 1), FIG3_OPS[i], cw - 2 * 0.09, tag=f"fig3 блок {i + 1}") for i in range(11)]
    rows = [[i for i in range(11) if CELL3[i][0] == r] for r in range(3)]
    for r in rows:
        equalize([cards[i] for i in r])
    rh = [max(cards[i].height() for i in r) for r in rows]
    tops = [M + 0.24]
    tops.append(tops[0] + rh[0] + 0.42)
    tops.append(tops[1] + rh[1] + 0.42)
    H = tops[2] + rh[2] + M
    D = Diagram(H)
    _, xr = snake(D, cards, CELL3, tops, rh, cw, gx, hl=FIG3_KKT)
    for i in FIG3_KKT:
        r, c = CELL3[i]
        D.center_text(["ККТ"], xr(c) - 0.03, tops[r] - 0.1, BODY_PT, "bold", NAVY, ha="right")
    D.save("fig3.png")


def fig4():
    A_X, A_W = 0.22, 1.14
    BC_X, BC_W = 1.62, 1.50
    D_X, D_W = 3.50, 1.20
    EG_X, EG_W = 4.98, 1.42
    cA = Card("Пользователь (браузер)", None, A_W - 2 * PAD_X, tag="fig4 A")
    cB = Card("вкусноты.рф", "сайт проекта (хостинг Beget, статика; подпапка /синестезия)",
              BC_W - 2 * PAD_X, tag="fig4 B")
    cC = Card("вкусныйтрек.рф", "промо-продукт «Синестезия» (хостинг Beget, статика)",
              BC_W - 2 * PAD_X, tag="fig4 C")
    cD = Card("synesthesia-app.vercel.app (Vercel):", "страница + серверные функции api/*",
              D_W - 2 * PAD_X, tag="fig4 D")
    cE = Card("Supabase «sinestezia»", "(Postgres): квоты, кэш анализов, короткие ссылки, рефералы",
              EG_W - 2 * PAD_X, tag="fig4 E")
    cF = Card("Supabase «vkus-tech»", "(Postgres): база данных сайта и заказов",
              A_W - 2 * PAD_X, tag="fig4 F")
    cG = Card("Внешние сервисы:", "языковая модель через OpenRouter; каталог iTunes (метаданные, обложки)",
              EG_W - 2 * PAD_X, tag="fig4 G")
    equalize([cB, cE])
    equalize([cC, cF, cG])
    h1 = max(cB.height(), cE.height())
    h2 = max(cD.height(), cA.height())
    h3 = max(cC.height(), cF.height(), cG.height())
    t1 = M
    t2 = t1 + h1 + 0.42
    t3 = t2 + h2 + 0.42
    H = t3 + h3 + M
    D = Diagram(H)

    hA, hD = cA.height(), cD.height()
    mid2 = t2 + h2 / 2
    tA, tD = mid2 - hA / 2, mid2 - hD / 2
    D.block(cB, BC_X, t1, BC_W, h1)
    D.block(cE, EG_X, t1, EG_W, h1)
    D.block(cA, A_X, tA, A_W, hA)
    D.block(cD, D_X, tD, D_W, hD)
    D.block(cC, BC_X, t3, BC_W, h3)
    D.block(cF, A_X, t3, A_W, h3)
    D.block(cG, EG_X, t3, EG_W, h3)

    b_mid = t1 + h1 / 2
    c_mid = t3 + h3 / 2
    f_mid = t3 + h3 / 2
    a_r = A_X + A_W
    a_up, a_dn = mid2 - 0.12, mid2 + 0.12
    bc_r = BC_X + BC_W
    bc_cx = BC_X + BC_W / 2
    d_cx, d_r = D_X + D_W / 2, D_X + D_W
    eg_cx = EG_X + EG_W / 2
    d_up, d_dn = mid2 - 0.2, mid2 + 0.2
    route_x = 0.09

    D.arrow([(a_r, a_up), (bc_cx, a_up), (bc_cx, t1 + h1)])          # A -> B
    D.arrow([(a_r, a_dn), (bc_cx, a_dn), (bc_cx, t3)])               # A -> C
    D.arrow([(bc_r, b_mid), (d_cx, b_mid), (d_cx, tD)])              # B -> D
    D.arrow([(bc_r, c_mid), (d_cx, c_mid), (d_cx, tD + hD)])         # C -> D
    D.arrow([(d_r, d_up), (eg_cx, d_up), (eg_cx, t1 + h1)])          # D -> E
    D.arrow([(d_r, d_dn), (eg_cx, d_dn), (eg_cx, t3)])               # D -> G
    D.arrow([(BC_X, b_mid), (route_x, b_mid), (route_x, f_mid), (A_X, f_mid)])  # B -> F

    D.center_text(["заявки/заказы"], (route_x + BC_X) / 2, b_mid - 0.11, BODY_PT)
    lx = (bc_r + d_cx) / 2
    D.center_text(["запросы к API"], lx, b_mid - 0.11, BODY_PT)
    D.center_text(["запросы к API"], lx, c_mid + 0.11, BODY_PT)
    D.save("fig4.png")


def fig5():
    ncol, gx = 3, 0.45
    cw = (X1 - X0 - (ncol - 1) * gx) / ncol
    cards = [Card(t, b, cw - 2 * PAD_X, tag=f"fig5 блок {i + 1}") for i, (t, b) in enumerate(FIG5_STEPS)]
    rows = [[0, 1, 2], [3, 4, 5]]
    for r in rows:
        equalize([cards[i] for i in r])
    rh = [max(cards[i].height() for i in r) for r in rows]
    tops = [M, M + rh[0] + 0.4]
    quota = Card(None, FIG5_QUOTA, X1 - X0 - 2 * PAD_X, bsize=TITLE_PT, hi=120, tag="fig5 квоты")
    y_q = tops[1] + rh[1] + 0.35
    H = y_q + quota.height() + M
    D = Diagram(H)
    snake(D, cards, CELL5, tops, rh, cw, gx)
    D.block(quota, X0, y_q, X1 - X0)
    D.save("fig5.png")


def fig6():
    kw, px_, pw = 2.1, 2.65, 2.15
    gx0 = 5.15
    gw = X1 - gx0
    kav, pav, gav = kw - 2 * PAD_X, pw - 2 * PAD_X, gw - 2 * PAD_X
    k1 = Card("Яндекс Директ", "(контекстная реклама, бюджет 15 000 руб.)", kav, tag="fig6 K1")
    k2 = Card("Наружная реклама:", "билборды на двух улицах Москвы, 6 дней", kav, tag="fig6 K2")
    k3 = Card("Совместный релиз с артисткой SavannaOki:",
              "трек на французском языке; пластинка оранжевого цвета, вкус «ваниль-бурбон»",
              kav, tag="fig6 K3")
    p1 = Card("вкусныйтрек.рф",
              "промо-продукт «Синестезия» (вовлечение, карточка трека, виральные сторис)",
              pav, tag="fig6 P1")
    p2 = Card("вкусноты.рф", "сайт проекта (описание продукта, форма заказа)", pav, tag="fig6 P2")
    goal = Card("Заявки от артистов, лейблов и организаторов мероприятий", None, gav, tag="fig6 цель")
    hk1, hk2, hk3 = k1.height(), k2.height(), k3.height()
    hp1, hp2, hg = p1.height(), p2.height(), goal.height()
    c1 = M + max(hk1, hp1) / 2
    c2 = c1 + hp1 / 2 + 0.6 + hp2 / 2
    c3 = c2 + hk2 / 2 + 0.45 + hk3 / 2
    H = c3 + hk3 / 2 + M
    D = Diagram(H)
    D.block(k1, X0, c1 - hk1 / 2, kw, hk1)
    D.block(k2, X0, c2 - hk2 / 2, kw, hk2)
    D.block(k3, X0, c3 - hk3 / 2, kw, hk3)
    D.block(p1, px_, c1 - hp1 / 2, pw, hp1)
    D.block(p2, px_, c2 - hp2 / 2, pw, hp2)
    D.block(goal, gx0, c2 - hg / 2, gw, hg)
    kr, pr = X0 + kw, px_ + pw
    pcx = px_ + pw / 2
    D.arrow([(kr, c1), (px_, c1)])                          # K1 -> P1
    D.arrow([(kr, c1 + 0.22), (px_, c2 - 0.2)])             # K1 -> P2
    D.arrow([(kr, c2), (px_, c2)])                          # K2 -> P2
    D.arrow([(kr, c3 - 0.2), (px_, c2 + 0.2)])              # K3 -> P2
    D.arrow([(pcx, c1 + hp1 / 2), (pcx, c2 - hp2 / 2)])     # P1 -> P2
    D.arrow([(pr, c2), (gx0, c2)])                          # P2 -> цель
    label = fit("переход «заказать пластинку»", 1.2, "reg", BODY_PT, hi=18, tag="fig6 подпись")
    D.center_text(label, pcx + 0.1, (c1 + hp1 / 2 + c2 - hp2 / 2) / 2, BODY_PT, ha="left")
    D.save("fig6.png")


def main():
    fig1()
    fig2()
    fig3()
    fig4()
    fig5()
    fig6()
    print("Переносы textwrap (фрагмент: символов в строке):")
    print("; ".join(f"{t}: {n}" for t, n in WRAP_LOG))


if __name__ == "__main__":
    main()
