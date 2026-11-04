"""Minimal markdown-ish -> HWPX builder.
Syntax per line:
  # H1   ## H2   ### H3
  > text           boxed/highlighted single-cell paragraph (shaded)
  | a | b |        table rows (first row = header, shaded). '|---|' rows ignored.
  - text           bullet
  blank            paragraph break
  other            normal paragraph. **bold** inline supported.
  @widths 1,3,2    column weight hint for next table
"""
import re
import sys
from hwpx import HwpxDocument

BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def build(src: str, out: str, header: str | None = None):
    doc = HwpxDocument.new()
    try:
        doc.set_page_margins(left=4252, right=4252, top=4252, bottom=4252)  # ~15mm
    except Exception:
        pass
    if header:
        try:
            doc.set_header_text(header)
        except Exception:
            pass
    st = {
        "h1": doc.ensure_run_style(bold=True, size=17, color="#1F3864"),
        "h2": doc.ensure_run_style(bold=True, size=13, color="#1F3864"),
        "h3": doc.ensure_run_style(bold=True, size=11, color="#2E5597"),
        "b": doc.ensure_run_style(bold=True, size=10),
        "n": doc.ensure_run_style(size=10),
        "s": doc.ensure_run_style(size=9, color="#555555"),
    }

    def para(text, base="n"):
        p = doc.add_paragraph("", char_pr_id_ref=st[base], include_run=False)
        pos = 0
        for m in BOLD_RE.finditer(text):
            if m.start() > pos:
                p.add_run(text[pos:m.start()], char_pr_id_ref=st[base])
            p.add_run(m.group(1), char_pr_id_ref=st["b"] if base == "n" else st[base])
            pos = m.end()
        if pos < len(text) or not text:
            p.add_run(text[pos:], char_pr_id_ref=st[base])
        return p

    lines = src.splitlines()
    i = 0
    widths = None
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith("@widths"):
            widths = [float(x) for x in ln.split()[1].split(",")]
            i += 1
            continue
        if ln.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                r = lines[i].strip()
                if not re.fullmatch(r"\|[\s\-:|]+\|", r):
                    rows.append([c.strip() for c in r.strip("|").split("|")])
                i += 1
            ncol = max(len(r) for r in rows)
            t = doc.add_table(len(rows), ncol)
            for ri, r in enumerate(rows):
                for ci in range(ncol):
                    txt = r[ci] if ci < len(r) else ""
                    txt = txt.replace("<br>", "\n")
                    t.set_cell_text(ri, ci, BOLD_RE.sub(r"\1", txt), split_paragraphs=True)
                    if ri == 0:
                        t.set_cell_shading(ri, ci, "#DCE6F2")
            if widths and len(widths) == ncol:
                t.set_column_widths(widths)
            widths = None
            para("")
            continue
        if ln.startswith("> "):
            t = doc.add_table(1, 1)
            t.set_cell_text(0, 0, BOLD_RE.sub(r"\1", ln[2:]).replace("<br>", "\n"), split_paragraphs=True)
            t.set_cell_shading(0, 0, "#FFF2CC")
            para("")
        elif ln.startswith("### "):
            para(ln[4:], "h3")
        elif ln.startswith("## "):
            para(ln[3:], "h2")
        elif ln.startswith("# "):
            para(ln[2:], "h1")
        elif ln.startswith("- "):
            para("  • " + ln[2:])
        elif ln.startswith("※"):
            para(ln, "s")
        elif ln == "":
            pass
        else:
            para(ln)
        i += 1
    doc.save_to_path(out)


if __name__ == "__main__":
    src = open(sys.argv[1], encoding="utf-8").read()
    build(src, sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    d = HwpxDocument.open(sys.argv[2])
    txt = d.export_text()
    print("OK", sys.argv[2], len(txt), "chars")
