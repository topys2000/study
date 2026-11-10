import re, sys
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "맑은 고딕"
NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x55, 0x97)
GRAY = RGBColor(0x59, 0x59, 0x59)
BOLD = re.compile(r"\*\*(.+?)\*\*")

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Mm(210), Mm(297)
for s in ("left_margin", "right_margin"):
    setattr(sec, s, Mm(18))
sec.top_margin, sec.bottom_margin = Mm(16), Mm(16)

st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(11.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(3)
st.paragraph_format.line_spacing = 1.2

# header / footer
hp = sec.header.paragraphs[0]
hp.text = "작전 3그룹 교관 강평자료 · 교관용 · 학생 배포 금지"
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for r in hp.runs:
    r.font.size = Pt(8); r.font.color.rgb = GRAY
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
def _field(p, code):
    r = p.add_run()
    a = OxmlElement("w:fldChar"); a.set(qn("w:fldCharType"), "begin")
    b = OxmlElement("w:instrText"); b.set(qn("xml:space"), "preserve"); b.text = code
    c = OxmlElement("w:fldChar"); c.set(qn("w:fldCharType"), "end")
    r._r.append(a); r._r.append(b); r._r.append(c)
    r.font.size = Pt(8)
_field(fp, "PAGE")


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), fill)
    tcPr.append(sh)


def cell_margins(table, top=60, bottom=60, left=100, right=100):
    tblPr = table._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
    tblPr.append(m)


def borders(table, color="A6A6A6", size=4, only_left=None):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for k in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{k}")
        if only_left is not None:
            if k == "left":
                e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "24"); e.set(qn("w:color"), only_left)
            else:
                e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(size)); e.set(qn("w:color"), color)
        b.append(e)
    tblPr.append(b)


def runs(p, text, size=None, color=None, bold_all=False):
    pos = 0
    parts = []
    for m in BOLD.finditer(text):
        parts.append((text[pos:m.start()], False)); parts.append((m.group(1), True)); pos = m.end()
    parts.append((text[pos:], False))
    for t, b in parts:
        if not t:
            continue
        r = p.add_run(t)
        r.bold = b or bold_all
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = color


def fill_cell(cell, text, size=10.5, bold=False, color=None, align=None):
    cell.paragraphs[0].text = ""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        if align: p.alignment = align
        runs(p, ln, size=size, color=color, bold_all=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def H1(t):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    runs(p, t, size=16, color=NAVY, bold_all=True)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr"); bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single"); bt.set(qn("w:sz"), "12"); bt.set(qn("w:color"), "1F3864"); bt.set(qn("w:space"), "2")
    bdr.append(bt); pPr.append(bdr)


def H2(t):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    runs(p, "▎" + t, size=13, color=BLUE, bold_all=True)


def P(t, size=11.5, color=None):
    p = doc.add_paragraph(); runs(p, t, size=size, color=color); return p


def B(t, lvl=0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(4 + 5 * lvl)
    p.paragraph_format.first_line_indent = Mm(-3.5)
    p.paragraph_format.space_after = Pt(2)
    runs(p, ("• " if lvl == 0 else "– ") + t, size=11.5)


def BOX(title, body, fill="FFF4D6", edge="C9A227"):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    borders(t, only_left=edge); cell_margins(t, 100, 100, 160, 140)
    c = t.cell(0, 0); shade(c, fill)
    c.paragraphs[0].text = ""
    if title:
        runs(c.paragraphs[0], title, size=12.5, color=NAVY, bold_all=True)
        c.paragraphs[0].paragraph_format.space_after = Pt(3)
    for i, ln in enumerate(body.split("\n")):
        p = c.paragraphs[0] if (i == 0 and not title) else c.add_paragraph()
        p.paragraph_format.space_after = Pt(2); p.paragraph_format.line_spacing = 1.2
        runs(p, ln, size=11.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def TABLE(rows, widths, head_fill="1F3864", first_col_fill="EEF3FA", size=10.5):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t); cell_margins(t)
    total = 174
    ws = [Mm(total * w / sum(widths)) for w in widths]
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), ws):
        gc.set(qn("w:w"), str(int(w.twips if hasattr(w, 'twips') else w / 635)))
    for ri, r in enumerate(rows):
        tr = t.rows[ri]._tr
        trPr = tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        if ri == 0:
            th = OxmlElement("w:tblHeader"); trPr.append(th)
        for ci, txt in enumerate(r):
            c = t.cell(ri, ci); c.width = ws[ci]
            if ri == 0:
                shade(c, head_fill)
                fill_cell(c, txt, size=size, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                if ci == 0 and first_col_fill:
                    shade(c, first_col_fill)
                    fill_cell(c, txt, size=size, bold=True)
                else:
                    fill_cell(c, txt, size=size)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


exec(open(sys.argv[1], encoding="utf-8").read())
doc.save(sys.argv[2])
print("saved", sys.argv[2])
