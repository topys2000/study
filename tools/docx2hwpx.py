"""Convert a python-docx document (paragraphs, tables, shaded boxes) into HWPX."""
import sys
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from hwpx import HwpxDocument

MM = 283.465
EMU_PER_MM = 36000

src, dst = sys.argv[1], sys.argv[2]
d = docx.Document(src)
default_pt = (d.styles['Normal'].font.size.pt if d.styles['Normal'].font.size else 10.5)

h = HwpxDocument.new()
h.set_page_margins(left=int(16 * MM), right=int(16 * MM), top=int(15 * MM), bottom=int(15 * MM))
TEXT_W = int(178 * MM)
FONT = '맑은 고딕'
_style_cache = {}


def char_style(bold=False, italic=False, color=None, size=None):
    key = (bold, italic, color, size)
    if key not in _style_cache:
        _style_cache[key] = h.ensure_run_style(bold=bold, italic=italic, color=color, size=size, font=FONT)
    return _style_cache[key]


def run_props(r, para_bold=False):
    rpr = r._r.rPr
    size = None
    color = None
    if rpr is not None:
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            size = int(sz.get(qn('w:val'))) / 2
        c = rpr.find(qn('w:color'))
        if c is not None and c.get(qn('w:val')) not in (None, 'auto'):
            color = '#' + c.get(qn('w:val'))
    bold = bool(r.bold) or para_bold
    return dict(bold=bold, italic=bool(r.italic), color=color, size=size or default_pt)


def align_of(p):
    a = p.alignment
    if a is None:
        return None
    return {0: 'LEFT', 1: 'CENTER', 2: 'RIGHT', 3: 'JUSTIFY'}.get(int(a))


def is_bullet(p):
    return p.style is not None and 'List' in (p.style.name or '')


def emit_runs(make_para, p):
    """Write a docx paragraph; '\n' in runs starts a new paragraph."""
    segs = [[]]
    if is_bullet(p):
        segs[-1].append(('• ', dict(bold=False, italic=False, color=None, size=default_pt)))
    for r in p.runs:
        props = run_props(r)
        parts = r.text.split('\n')
        for i, part in enumerate(parts):
            if i > 0:
                segs.append([])
            if part:
                segs[-1].append((part, props))
    paras = []
    for seg in segs:
        hp = make_para()
        for text, props in seg:
            hp.add_run(text, char_pr_id_ref=char_style(**props))
        if not seg:
            hp.add_run('', char_pr_id_ref=char_style(size=default_pt))
        paras.append(hp)
    return paras


def shading_of(el_pr):
    if el_pr is None:
        return None
    shd = el_pr.find(qn('w:shd'))
    if shd is None:
        return None
    fill = shd.get(qn('w:fill'))
    return None if fill in (None, 'auto', 'FFFFFF') else '#' + fill


def has_page_break(p):
    return any(br.get(qn('w:type')) == 'page' for br in p._p.iter(qn('w:br')))


page_break_pending = False
para_index_formats = []  # (paragraph object, align, page_break)


def body_para():
    hp = h.add_paragraph('', include_run=False, inherit_style=False)
    return hp


def fill_cell(cell_h, cell_d):
    first = [True]

    def make():
        if first[0]:
            first[0] = False
            p0 = cell_h.paragraphs[0]
            p0.clear_text()
            return p0
        return cell_h.add_paragraph('')
    for p in cell_d.paragraphs:
        emit_runs(make, p)


def add_box(p, fill):
    t = h.add_table(1, 1, width=TEXT_W, border_fill_id_ref=h.ensure_border_fill(border_color='#13294B', fill_color=fill))
    c = t.cell(0, 0)
    fill_cell(c, type('C', (), {'paragraphs': [p]})())
    h.add_paragraph('', include_run=True, inherit_style=False)


for child in d.element.body.iterchildren():
    if child.tag == qn('w:p'):
        p = Paragraph(child, d)
        if has_page_break(p) and not p.text.strip():
            page_break_pending = True
            continue
        fill = shading_of(child.pPr)
        if fill:
            add_box(p, fill)
            continue
        start = len(h.paragraphs)
        emit_runs(body_para, p)
        pf = p.paragraph_format
        fmt = {}
        if align_of(p):
            fmt['alignment'] = align_of(p)
        if page_break_pending:
            fmt['page_break_before'] = True
        if pf.space_before is not None:
            fmt['spacing_before_pt'] = pf.space_before.pt
        if pf.space_after is not None:
            fmt['spacing_after_pt'] = pf.space_after.pt
        if is_bullet(p):
            fmt['indent_left_mm'] = 4
        n_new = len(h.paragraphs) - start
        for k in range(n_new):
            f2 = dict(fmt)
            if k > 0:
                f2.pop('page_break_before', None); f2.pop('spacing_before_pt', None)
            if k < n_new - 1:
                f2.pop('spacing_after_pt', None)
            if f2:
                para_index_formats.append((start + k, f2))
        page_break_pending = False
    elif child.tag == qn('w:tbl'):
        t = Table(child, d)
        rows, cols = len(t.rows), len(t.columns)
        widths = [c.width or 1 for c in t.columns]
        total = sum(widths)
        ht = h.add_table(rows, cols, width=TEXT_W)
        ht.set_column_widths([int(TEXT_W * w / total) for w in widths])
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row.cells):
                hc = ht.cell(ri, ci)
                fill_cell(hc, cell)
                f = shading_of(cell._tc.tcPr)
                if f:
                    ht.set_cell_shading(ri, ci, f) if True else None
        h.add_paragraph('', include_run=True, inherit_style=False)

# paragraph formatting pass
all_paras = list(h.paragraphs)
for idx, kw in para_index_formats:
    h.set_paragraph_format(paragraph_index=idx, **kw)

h.save_to_path(dst)
print('saved', dst)
