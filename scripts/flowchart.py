"""GOST 19.701-90 flowchart: text description -> OUT.png (300 dpi, true size, for Word) + OUT.drawio (editable).
  python flowchart.py CHART.txt OUT   (syntax: flowcharts.md)
"""
import os, re, sys
from xml.sax.saxutils import escape
from PIL import Image, ImageDraw, ImageFont

FONT_FILE, DPI, SS = r"C:\Windows\Fonts\times.ttf", 300, 2
DRAWIO_K = 96 / 25.4
SIMPLE = {"start": "start", "end": "end", "do": "process", "io": "io", "call": "call", "prep": "prep", "ref": "ref"}


class Cfg:
    a, font, gap, hgap = 20.0, 12.0, 8.0, 8.0
    uniform, yes_right, minw, minh = 0.0, False, 0.0, 0.0

    def __init__(self):
        self.measure = ImageFont.truetype(FONT_FILE, 100)

    @property
    def em(self): return self.font * 25.4 / 72

    @property
    def lh(self): return self.em * 1.15

    def tw(self, text): return self.measure.getlength(text) / 100 * self.em


class Shape:
    n = 0

    def __init__(self, kind, lines, w, h):
        Shape.n += 1
        self.id, self.kind, self.lines, self.w, self.h, self.x, self.y = f"s{Shape.n}", kind, lines, w, h, 0.0, 0.0

    def pt(self, fx, fy): return (self.x - self.w / 2 + fx * self.w, self.y + fy * self.h)


class Edge:
    def __init__(self, src, exit_, tgt, entry, wps, label=None, loff=(0, 0), dashed=False, end=None):
        self.src, self.exit, self.tgt, self.entry, self.wps = src, exit_, tgt, entry, list(wps)
        self.label, self.loff, self.dashed, self.end = label, loff, dashed, end

    def points(self):
        last = self.end if self.end is not None else self.tgt.pt(*self.entry)
        return [self.src.pt(*self.exit)] + self.wps + [last]

    def arrow(self):
        if self.end is not None:
            return True
        p = self.points()
        return not self.dashed and any(b[1] < a[1] - 1e-6 or b[0] < a[0] - 1e-6 for a, b in zip(p, p[1:]))


class Pending:
    """Unconnected exit: line from src runs through wps, then straight down at x=col to the box bottom."""
    def __init__(self, src, exit_, wps, col, label=None, loff=(0, 0)):
        self.src, self.exit, self.wps, self.col, self.label, self.loff = src, exit_, list(wps), col, label, loff


class Box:
    def __init__(self):
        self.shapes, self.edges, self.exits, self.entry = [], [], [], None
        self.left = self.right = self.h = 0.0

    def shift(self, dx, dy):
        for s in self.shapes: s.x += dx; s.y += dy
        for e in self.edges:
            e.wps = [(x + dx, y + dy) for x, y in e.wps]
            if e.end is not None: e.end = (e.end[0] + dx, e.end[1] + dy)
        for p in self.exits: p.wps = [(x + dx, y + dy) for x, y in p.wps]; p.col += dx
        return self

    def take(self, other):
        self.shapes += other.shapes; self.edges += other.edges


YES, NO = (3.5, 3.0), (4.5, -3.0)


def connect(p, tgt, join_y):
    wps = p.wps + ([(p.col, join_y), (0.0, join_y)] if abs(p.col) > 1e-6 else [])
    return Edge(p.src, p.exit, tgt, (0.5, 0), wps, p.label, p.loff)


def merge(exits, join_y, full):
    """One exit draws the whole path (full); the others stop with an arrow where they enter its line at join_y."""
    if len(exits) <= 1:
        return [full(p) for p in exits]
    main = next((p for p in exits if abs(p.col) < 1e-6), None) or max(exits, key=lambda p: p.col)
    rest = [p for p in exits if p is not main]
    outer = max(rest, key=lambda p: p.col) if abs(main.col) < 1e-6 else None
    edges = [full(main)]
    for p in rest:
        if p is outer:
            edges.append(Edge(p.src, p.exit, None, None, p.wps + [(p.col, join_y)], p.label, p.loff,
                              end=(0.0, join_y)))
        else:
            edges.append(Edge(p.src, p.exit, None, None, p.wps, p.label, p.loff, end=(p.col, join_y)))
    return edges


class Builder:
    def __init__(self, cfg):
        self.c, self.warn = cfg, []

    @staticmethod
    def chunks(par):
        """Words joined so that a line breaks only after , ; or between two plain words (formulas stay whole)."""
        word = re.compile(r"[A-Za-zА-Яа-яЁё]{2,}[.,;:!?]?")
        out = []
        for t in par.split():
            if out and not (out[-1][-1] in ",;" or (word.fullmatch(out[-1].split()[-1]) and word.fullmatch(t))):
                out[-1] += " " + t
            else:
                out.append(t)
        return out

    def wrap(self, text, width):
        out = []
        for par in text.split("\\n"):
            cur = ""
            for w in self.chunks(par):
                if self.c.tw(w) > width: return None
                t = (cur + " " + w).strip()
                if self.c.tw(t) <= width: cur = t
                else: out.append(cur); cur = w
            out.append(cur)
        return out

    def fit(self, kind, text):
        c, a = self.c, self.c.a
        b = 1.5 * a
        w, h = {"start": (b, a / 2), "end": (b, a / 2), "ref": (a / 2, a / 2)}.get(kind, (b, a))
        if kind != "ref":
            w = max(w, c.minw)
            if kind not in ("start", "end", "decision"):
                h = max(h, c.minh)
        for _ in range(80):
            nmax = max(1, int((h - 1.5) // c.lh))
            for n in range(1, nmax + 1):
                avail = {"decision": w * (1 - n * c.lh / h) - 3, "io": w - a / 2 - 3, "call": w - a / 4 - 4,
                         "prep": w - a / 2 - 2, "hex": w - a / 2 - 2, "start": w - h, "end": w - h}.get(kind, w - 4)
                lines = self.wrap(text, avail) if avail > 0 else None
                if lines is not None and len(lines) <= n:
                    return lines, w, h
            if kind == "ref":
                w += 5; h += 5
            elif kind == "decision" and (w >= 2.5 * b or c.minw):
                h += 5
            else:
                w += 5
        self.warn.append(f"текст не помещается в блок: {text[:40]}")
        return [text], w, h

    def shape(self, kind, text):
        return Shape(kind, *self.fit(kind, text))

    def simple(self, kind, text, comment=None):
        s = self.shape(kind, text)
        bx = Box(); bx.shapes = [s]; bx.left = bx.right = s.w / 2; bx.h = s.h
        bx.entry = None if kind == "start" else s
        bx.exits = [] if kind == "end" else [Pending(s, (0.5, 1), [], 0.0)]
        if comment:
            lines = self.wrap(comment, 45) or [comment]
            cw = max(self.c.tw(l) for l in lines) + 7
            cm = Shape("comment", lines, cw, len(lines) * self.c.lh + 2)
            cm.x, cm.y = s.w / 2 + 8 + cw / 2, s.h / 2 - cm.h / 2
            bx.shapes.append(cm)
            bx.edges.append(Edge(s, (1, 0.5), cm, (0, 0.5), [], dashed=True))
            bx.right = s.w / 2 + 8 + cw; bx.h = max(s.h, cm.y + cm.h)
        return bx

    def seq(self, boxes):
        G, out, y, prev = self.c.gap, Box(), 0.0, None
        for bx in boxes:
            if prev is not None:
                y += G
            bx.shift(0, y)
            if prev is not None:
                if bx.entry is None:
                    self.warn.append("блок «start» не в начале схемы")
                else:
                    out.edges += merge(prev.exits, y - G, lambda p: connect(p, bx.entry, y - G))
            elif out.entry is None:
                out.entry = bx.entry
            out.take(bx)
            out.left, out.right = max(out.left, bx.left), max(out.right, bx.right)
            y += bx.h; prev = bx
        out.h, out.exits = y, prev.exits if prev else []
        return out

    def if_(self, cond, T, E):
        G, H = self.c.gap, self.c.hgap
        d = self.shape("decision", cond)
        mid = d.h / 2
        T.shift(0, d.h + G)
        bx = Box(); bx.shapes = [d]; bx.entry = d
        down, right = ("нет", "да") if self.c.yes_right else ("да", "нет")
        bx.edges.append(Edge(d, (0.5, 1), T.entry, (0.5, 0), [], down, YES))
        if E:
            xr = max(d.w / 2 + 8, T.right + H + E.left)
            E.shift(xr, d.h + G)
            bx.edges.append(Edge(d, (1, 0.5), E.entry, (0.5, 0), [(xr, mid)], right, NO))
            bx.take(E); bx.exits = T.exits + E.exits
            bx.right = xr + E.right
        else:
            xr = max(d.w / 2 + 8, T.right + H)
            bx.exits = T.exits + [Pending(d, (1, 0.5), [(xr, mid)], xr, right, NO)]
            bx.right = xr
        bx.take(T)
        bx.left = max(d.w / 2, T.left)
        bx.h = d.h + G + max(T.h, E.h if E else 0) + G / 2
        return bx

    def while_(self, cond, B, kind="decision"):
        G = self.c.gap
        d = self.shape(kind, cond)
        mid = d.h / 2
        B.shift(0, d.h + G)
        yb = d.h + G + B.h + G / 2
        xl, xr = max(d.w / 2, B.left) + 8, max(d.w / 2 + 8, B.right + 8)
        lab = kind == "decision"
        bx = Box(); bx.shapes = [d]; bx.entry = d
        bx.edges.append(Edge(d, (0.5, 1), B.entry, (0.5, 0), [], "да" if lab else None, YES))
        bx.edges += merge(B.exits, yb, lambda p: Edge(p.src, p.exit, d, (0, 0.5),
                                                      p.wps + [(p.col, yb), (-xl, yb), (-xl, mid)], p.label, p.loff))
        bx.exits = [Pending(d, (1, 0.5), [(xr, mid)], xr, "нет" if lab else None, NO)]
        bx.take(B)
        bx.left, bx.right, bx.h = xl, xr, yb + G / 2
        return bx

    def repeat(self, B, cond, until):
        G = self.c.gap
        d = self.shape("decision", cond)
        head = G / 2
        B.shift(0, head)
        yd = head + B.h + G
        d.y = yd
        xl = max(d.w / 2, B.left) + 8
        bx = Box(); bx.shapes = [d]; bx.entry = B.entry
        bx.edges += merge(B.exits, head + B.h, lambda p: connect(p, d, head + B.h))
        bx.edges.append(Edge(d, (0, 0.5), B.entry, (0.5, 0), [(-xl, yd + d.h / 2), (-xl, 0.0), (0.0, 0.0)],
                             "нет" if until else "да", (-4.5, -3.0)))
        bx.exits = [Pending(d, (0.5, 1), [], 0.0, "да" if until else "нет", YES)]
        bx.take(B)
        bx.left, bx.right, bx.h = xl, max(B.right, d.w / 2), yd + d.h
        return bx

    def build(self, nodes):
        boxes = []
        for nd in nodes:
            k = nd[0]
            if k == "simple":
                boxes.append(self.simple(*nd[1:]))
            elif k == "if":
                boxes.append(self.if_(nd[1], self.build(nd[2]), self.build(nd[3]) if nd[3] else None))
            elif k == "while":
                boxes.append(self.while_(nd[1], self.build(nd[2])))
            elif k == "loop":
                boxes.append(self.while_(nd[1], self.build(nd[2]), "hex"))
            elif k == "repeat":
                boxes.append(self.repeat(self.build(nd[1]), nd[2], nd[3]))
            elif k == "for":
                boxes.append(self.seq([self.simple("loopstart", nd[1]), self.build(nd[3]), self.simple("loopend", nd[2])]))
        return self.seq(boxes)


def parse(text, cfg):
    items = []
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith(";"):
            continue
        m = re.match(r"^(a|font|gap|hgap|uniform)\s*:\s*([\d.]+)\s*$", raw.strip())
        if m and not items:
            setattr(cfg, m.group(1), float(m.group(2))); continue
        m = re.match(r"^yes\s*:\s*(right|down)\s*$", raw.strip())
        if m and not items:
            cfg.yes_right = m.group(1) == "right"; continue
        items.append((len(raw) - len(raw.lstrip()), raw.strip()))

    def kw(i):
        return items[i][1].split(None, 1)[0] if i < len(items) else ""

    def rest(i):
        p = items[i][1].split(None, 1)
        return p[1] if len(p) > 1 else ""

    def body(i, ind, what):
        if i + 1 >= len(items) or items[i + 1][0] <= ind:
            sys.exit(f"error: пустое тело «{what}»: {items[i][1]}")
        return block(i + 1, items[i + 1][0])

    def block(i, ind):
        nodes = []
        while i < len(items) and items[i][0] >= ind:
            if items[i][0] > ind:
                sys.exit(f"error: лишний отступ: {items[i][1]}")
            k, r, own = kw(i), rest(i), i
            if k == "if":
                t, i = body(i, ind, "if")
                e = None
                if i < len(items) and items[i][0] == ind and items[i][1] == "else":
                    e, i = body(i, ind, "else")
                nodes.append(("if", r, t, e))
            elif k in ("while", "loop"):
                b, i = body(i, ind, k)
                nodes.append((k, r, b))
            elif k == "for":
                b, i = body(i, ind, "for")
                head, _, end = r.partition("|")
                nodes.append(("for", head.strip(), (end or head).strip(), b))
            elif k == "repeat":
                b, i = body(i, ind, "repeat")
                if i >= len(items) or items[i][0] != ind or kw(i) not in ("until", "while"):
                    sys.exit(f"error: после тела repeat нужна строка «until условие» или «while условие»")
                nodes.append(("repeat", b, rest(i), kw(i) == "until")); i += 1
            else:
                kind = SIMPLE.get(k)
                parts = re.split(r"\s//\s", items[own][1] + " ", maxsplit=1)
                cm = parts[1].strip() if len(parts) > 1 else None
                t = parts[0].strip()
                if kind:
                    t = (t.split(None, 1) + [""])[1] or {"start": "Начало", "end": "Конец"}.get(kind, "")
                nodes.append(("simple", kind or "process", t, cm)); i += 1
        return nodes, i

    return block(0, items[0][0] if items else 0)[0]


def bounds(shapes, edges, cfg):
    xs, ys = [], []
    for s in shapes:
        xs += [s.x - s.w / 2, s.x + s.w / 2]; ys += [s.y, s.y + s.h]
        if s.kind == "comment":
            xs.append(s.x + s.w / 2)
    for e in edges:
        for x, y in e.points(): xs.append(x); ys.append(y)
        if e.label:
            x, y = e.points()[0]; xs.append(x + e.loff[0] + 4); ys += [y + e.loff[1] - 3, y + e.loff[1] + 3]
            xs.append(x + e.loff[0] - 4)
    m = 2
    return min(xs) - m, min(ys) - m, max(xs) + m, max(ys) + m


def outline(s, a):
    x0, y0, x1, y1, cx = s.x - s.w / 2, s.y, s.x + s.w / 2, s.y + s.h, s.x
    q = a / 4
    return {"decision": [(cx, y0), (x1, y0 + s.h / 2), (cx, y1), (x0, y0 + s.h / 2)],
            "io": [(x0 + q, y0), (x1, y0), (x1 - q, y1), (x0, y1)],
            "prep": [(x0 + q, y0), (x1 - q, y0), (x1, y0 + s.h / 2), (x1 - q, y1), (x0 + q, y1), (x0, y0 + s.h / 2)],
            "loopstart": [(x0 + q, y0), (x1 - q, y0), (x1, y0 + q), (x1, y1), (x0, y1), (x0, y0 + q)],
            "loopend": [(x0, y0), (x1, y0), (x1, y1 - q), (x1 - q, y1), (x0 + q, y1), (x0, y1 - q)],
            }.get("prep" if s.kind == "hex" else s.kind, [(x0, y0), (x1, y0), (x1, y1), (x0, y1)])


def render_png(shapes, edges, cfg, path):
    minx, miny, maxx, maxy = bounds(shapes, edges, cfg)
    k = DPI * SS / 25.4
    W, H = int((maxx - minx) * k) + 1, int((maxy - miny) * k) + 1
    img = Image.new("RGB", (W, H), "white")
    dr = ImageDraw.Draw(img)
    lw = max(1, round(0.3 * k))
    font = ImageFont.truetype(FONT_FILE, round(cfg.font / 72 * DPI * SS))
    P = lambda x, y: ((x - minx) * k, (y - miny) * k)

    def text(lines, cx, cy, left=None):
        y = cy - len(lines) * cfg.lh / 2 + cfg.lh / 2
        for ln in lines:
            if left is None:
                dr.text(P(cx, y), ln, font=font, fill="black", anchor="mm")
            else:
                dr.text(P(left, y), ln, font=font, fill="black", anchor="lm")
            y += cfg.lh

    for e in edges:
        pts = e.points()
        if e.dashed:
            for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
                L = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5; t = 0.0
                while t < L:
                    t2 = min(t + 2, L)
                    dr.line([P(x0 + (x1 - x0) * t / L, y0 + (y1 - y0) * t / L),
                             P(x0 + (x1 - x0) * t2 / L, y0 + (y1 - y0) * t2 / L)], fill="black", width=lw)
                    t += 3
        else:
            dr.line([P(*p) for p in pts], fill="black", width=lw, joint="curve")
        if e.arrow():
            (x0, y0), (x1, y1) = pts[-2], pts[-1]
            L = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5 or 1
            ux, uy = (x1 - x0) / L, (y1 - y0) / L
            bx, by = x1 - ux * 3, y1 - uy * 3
            dr.line([P(bx - uy * 1.2, by + ux * 1.2), P(x1, y1), P(bx + uy * 1.2, by - ux * 1.2)],
                    fill="black", width=lw, joint="curve")
        if e.label:
            x, y = pts[0]
            dr.text(P(x + e.loff[0], y + e.loff[1]), e.label, font=font, fill="black", anchor="mm")
    for s in shapes:
        x0, y0, x1, y1 = s.x - s.w / 2, s.y, s.x + s.w / 2, s.y + s.h
        if s.kind in ("start", "end"):
            dr.rounded_rectangle([P(x0, y0), P(x1, y1)], radius=s.h / 2 * k, fill="white", outline="black", width=lw)
        elif s.kind == "ref":
            dr.ellipse([P(x0, y0), P(x1, y1)], fill="white", outline="black", width=lw)
        elif s.kind == "comment":
            dr.line([P(x0 + 4, y0), P(x0, y0), P(x0, y1), P(x0 + 4, y1)], fill="black", width=lw)
            text(s.lines, 0, s.y + s.h / 2, left=x0 + 2.5)
            continue
        else:
            dr.polygon([P(*p) for p in outline(s, cfg.a)], fill="white", outline="black", width=lw)
            if s.kind == "call":
                i = cfg.a / 8
                for xx in (x0 + i, x1 - i):
                    dr.line([P(xx, y0), P(xx, y1)], fill="black", width=lw)
        text(s.lines, s.x, s.y + s.h / 2)
    img = img.resize((W // SS, H // SS), Image.LANCZOS)
    img.save(path, dpi=(DPI, DPI))
    return maxx - minx, maxy - miny


def render_drawio(shapes, edges, cfg, path):
    minx, miny, _, _ = bounds(shapes, edges, cfg)
    K = DRAWIO_K
    X = lambda x: round((x - minx) * K + 10, 1)
    Y = lambda y: round((y - miny) * K + 10, 1)
    q = round(cfg.a / 4 * K, 1)
    fs = round(cfg.font * 96 / 72)
    base = (f"html=1;whiteSpace=wrap;fontFamily=Times New Roman;fontSize={fs};fontColor=#000000;"
            "strokeColor=#000000;fillColor=#FFFFFF;")
    cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
    val = lambda lines: escape("<br>".join(escape(l) for l in lines), {'"': "&quot;"})
    for s in shapes:
        st = {"process": "rounded=0;", "start": "rounded=1;arcSize=50;", "end": "rounded=1;arcSize=50;",
              "decision": "rhombus;", "ref": "ellipse;aspect=fixed;",
              "io": f"shape=parallelogram;perimeter=parallelogramPerimeter;fixedSize=1;size={q};",
              "call": f"shape=process;backgroundOutline=1;size={round(cfg.a / 8 / s.w, 4)};",
              "prep": f"shape=hexagon;perimeter=hexagonPerimeter2;fixedSize=1;size={q};",
              "hex": f"shape=hexagon;perimeter=hexagonPerimeter2;fixedSize=1;size={q};",
              "loopstart": f"shape=loopLimit;size={q};",
              "loopend": f"shape=loopLimit;size={q};flipV=1;",
              }.get(s.kind)
        x, w = s.x - s.w / 2, s.w
        if s.kind == "comment":
            st = (f"html=1;whiteSpace=nowrap;fontFamily=Times New Roman;fontSize={fs};fontColor=#000000;"
                  "strokeColor=#000000;fillColor=none;shape=partialRectangle;right=0;"
                  "labelPosition=right;verticalLabelPosition=middle;align=left;verticalAlign=middle;spacingLeft=2;")
            w = 4
        else:
            st = base + st
        cells.append(f'<mxCell id="{s.id}" value="{val(s.lines)}" style="{st}" vertex="1" parent="1">'
                     f'<mxGeometry x="{X(x)}" y="{Y(s.y)}" width="{round(w * K, 1)}" height="{round(s.h * K, 1)}" as="geometry"/></mxCell>')
    for i, e in enumerate(edges, 1):
        entry = "" if e.end is not None else \
            f"entryX={e.entry[0]};entryY={e.entry[1]};entryDx=0;entryDy=0;"
        st = (f"edgeStyle=none;rounded=0;html=1;endArrow={'open' if e.arrow() else 'none'};endFill=0;endSize=8;"
              f"exitX={e.exit[0]};exitY={e.exit[1]};exitDx=0;exitDy=0;{entry}"
              f"strokeColor=#000000;fontFamily=Times New Roman;fontSize={fs};"
              f"fontColor=#000000;labelBackgroundColor=none;" + ("dashed=1;" if e.dashed else ""))
        pts = "".join(f'<mxPoint x="{X(x)}" y="{Y(y)}"/>' for x, y in e.wps)
        off = f'<mxPoint x="{round(e.loff[0] * K, 1)}" y="{round(e.loff[1] * K, 1)}" as="offset"/>' if e.label else ""
        if e.end is not None:
            tgt, tp = "", f'<mxPoint x="{X(e.end[0])}" y="{Y(e.end[1])}" as="targetPoint"/>'
        else:
            tgt, tp = f' target="{e.tgt.id}"', ""
        cells.append(f'<mxCell id="e{i}" value="{escape(e.label or "")}" style="{st}" edge="1" parent="1" '
                     f'source="{e.src.id}"{tgt}><mxGeometry x="-1" relative="1" as="geometry">'
                     f'{off}{tp}<Array as="points">{pts}</Array></mxGeometry></mxCell>')
    xml = ('<mxfile host="bsuir-word-docs"><diagram id="flowchart" name="Блок-схема">'
           '<mxGraphModel grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" '
           'pageScale="1" pageWidth="827" pageHeight="1169" math="0" shadow="0"><root>'
           + "".join(cells) + "</root></mxGraphModel></diagram></mxfile>")
    open(path, "w", encoding="utf-8").write(xml)


def main(src, out):
    cfg = Cfg()
    nodes = parse(open(src, encoding="utf-8-sig").read(), cfg)
    if not nodes:
        sys.exit("error: пустое описание схемы")
    b = Builder(cfg)
    box = b.build(nodes)
    if cfg.uniform:
        blocks = [s for s in box.shapes if s.kind not in ("start", "end", "ref", "comment", "decision")]
        if blocks:
            cfg.minw = max(s.w for s in box.shapes if s.kind not in ("ref", "comment", "decision"))
            cfg.minh = max(s.h for s in blocks)
            b = Builder(cfg)
            box = b.build(nodes)
    base = os.path.splitext(out)[0]
    w, h = render_png(box.shapes, box.edges, cfg, base + ".png")
    render_drawio(box.shapes, box.edges, cfg, base + ".drawio")
    print(f"png: {base}.png ({w / 10:.1f} x {h / 10:.1f} cm)\ndrawio: {base}.drawio")
    if h > 230:
        b.warn.append(f"высота {h / 10:.1f} см > 23 см: разбейте схему соединителями ref на листы")
    if w > 165:
        b.warn.append(f"ширина {w / 10:.1f} см > 16,5 см: при вставке будет уменьшена")
    for wmsg in b.warn:
        print("warn:", wmsg)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
