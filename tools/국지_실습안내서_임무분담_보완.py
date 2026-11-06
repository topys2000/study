"""실습안내서에 소조 임무 분담(발표 4·질의 4)과 질의 준비 양식을 보완한다."""
import copy
import re
import sys
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Emu

SRC, DST = sys.argv[1], sys.argv[2]
NAVY, BLUE, GREY = '13294B', '2B5AA8', '5A6475'
d = docx.Document(SRC)
body = d.element.body


def runs(p, text, size=10.5, color=None, bold_all=False):
    for i, part in enumerate(re.split(r'\*\*', text)):
        if not part:
            continue
        r = p.add_run(part)
        r.font.size = Pt(size)
        r.font.name = '맑은 고딕'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        r.bold = bold_all or (i % 2 == 1)
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
    return p


def shade(pr, fill):
    s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), fill)
    pr.append(s)


new = []  # elements created, in order, to be moved


def para(text='', size=10.5, color=None, bold=False, before=None, style=None):
    p = d.add_paragraph(style=style)
    if before:
        p.paragraph_format.space_before = Pt(before)
    runs(p, text, size, color, bold)
    new.append(p._p)
    return p


def box(lines, fill='EEF3FA'):
    p = para()
    for i, ln in enumerate(lines):
        if i:
            p.add_run().add_break()
        runs(p, ln, 10.5)
    shade(p._p.get_or_add_pPr(), fill)
    return p


def table(header, rows, widths, first_col_fill='EEF3FA', size=10):
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    tot = sum(widths)
    for ri, row in enumerate([header] + rows):
        for ci, txt in enumerate(row):
            c = t.cell(ri, ci)
            c.width = Emu(int(6400000 * widths[ci] / tot))
            p = c.paragraphs[0]
            for k, ln in enumerate(txt.split('\n')):
                if k:
                    p.add_run().add_break()
                runs(p, ln, size, 'FFFFFF' if ri == 0 else None, ri == 0)
            if ri == 0:
                shade(c._tc.get_or_add_tcPr(), NAVY)
            elif ci == 0 and first_col_fill:
                shade(c._tc.get_or_add_tcPr(), first_col_fill)
    tblPr = t._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side); e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), 'BFBFBF')
        b.append(e)
    tblPr.append(b)
    new.append(t._tbl)
    new.append(d.add_paragraph()._p)
    return t


def take():
    out = list(new); new.clear()
    for el in out:
        body.remove(el)
    return out


def text_of(el):
    return ''.join(x.text or '' for x in el.iter(qn('w:t')))


def set_cell(cell, text):
    p = cell.paragraphs[0]
    rs = p.runs
    rs[0].text = text
    for r in rs[1:]:
        r._r.getparent().remove(r._r)


# 1) 기존 장 번호 1~7 → 2~8
for p in d.paragraphs:
    m = re.match(r'^([1-7])\. ', p.text)
    if m and p.runs and p.runs[0].bold:
        r0 = p.runs[0]
        r0.text = re.sub(r'^[1-7]', str(int(m.group(1)) + 1), r0.text)

# 2) 새 1장: 소조 편성과 임무 분담
para('1. 소조 편성과 임무 분담 — 발표팀 4명 · 질의팀 4명', 14.5, NAVY, True, before=10)
box(['**위기평가판단서는 소조 8명이 함께 쓴다.** 판단서가 완성되면 소조를 둘로 나눈다.',
     '**발표팀 4명**은 우리 조의 분석과 대응방책을 합참의장 보고 형식으로 발표하고 질의에 답한다.',
     '**질의팀 4명**은 다른 조의 발표자료를 미리 받아 분석과 방책의 빈틈을 찾고 질의를 준비한다.',
     '두 조는 **교차**한다: 1조 발표 ↔ 2조 질의팀, 2조 발표 ↔ 1조 질의팀.'], fill='FFF8E8')
para('발표팀 — 우리 조의 분석과 대응방책', 12.5, BLUE, True, before=6)
table(['역할', '맡는 범위', '발표의 핵심 질문'], [
    ['발표 1', '① 상황개요 · ② 전략 상황', '무슨 일이 있었고, 어떤 정세 속에서 일어났는가?'],
    ['발표 2', '② 군사 상황 · ③ 종합평가 · ④ 대응방향', '적은 왜 도발했고 다음에 무엇을 할 것인가? 그래서 합참이 나서야 하는가(필요성)?'],
    ['발표 3', '작전목적 · 최종상태 · 작전목표 · 2항 임무 · 3항 작전개념', '어떤 상태로 위기를 끝낼 것인가? 대응의 수준은 목적에 맞는가(비례성)?'],
    ['발표 4', '부대별 과업 · 위험·완화방안(북한 역대응 포함) · SC · 결론', '누가 무엇을 하는가? 이 방책이 만드는 위험은 무엇이고 어떻게 줄이는가?'],
], [0.9, 2.6, 3.5])
para('질의팀 — 다른 조의 분석과 대응방책 검증', 12.5, BLUE, True, before=6)
table(['역할', '검증 대상 (상대 조)', '질의 착안점'], [
    ['질의 1', '상황평가', '적 의도·예상 양상 판단에 근거(징후·정보)가 있는가? 다른 해석은 없는가?'],
    ['질의 2', '종합평가 · 대응방향', '합참 대응의 필요성을 부대 자위권과 구분해 판단했는가? 위기완화(또는 타격)를 택하지 않은 이유는?'],
    ['질의 3', '작전목적 · 최종상태 · 작전개념', '대응 수준이 비례성에 맞는가? 최종상태가 위기 이전으로의 회복을 지향하는가? 작전개념이 합동작전인가?'],
    ['질의 4', '위험·완화방안 · 연계 확인란', '평가에서 찾은 위협이 방책에서 모두 다뤄졌는가? 북한이 역대응하면 다음 날은 어떤 상태인가? + 질의 종합'],
], [0.9, 2.2, 3.9])
para('운영 규칙', 12.5, BLUE, True, before=6)
for t in ['판단서 완성 직후 두 조는 **발표자료를 교환**한다. 질의팀은 발표 전까지 상대 자료를 검토해 질의를 준비한다(양식 ③).',
          '질의는 **"상대 자료의 어느 판단에서 → 무엇이 빠졌거나 맞지 않는가 → 무엇을 묻는가"** 순서로 한다. 근거 없는 질의, 결론만 바꾸라는 질의는 하지 않는다.',
          '발표팀은 답변할 때 **판단서의 어느 항목(연계 확인란)에 근거했는지** 함께 말한다. 즉답이 어려우면 보완 방향을 제시한다.',
          '질의팀도 평가 대상이다. **좋은 질의는 상대 방책을 더 낫게 만드는 질의**다.']:
    para(t, style='List Bullet')
section1 = take()

# 임무 표(3번 요소) 뒤 빈 문단 다음에 삽입
children = list(body.iterchildren())
anchor = children[4]  # 임무 표 다음 빈 문단
for el in reversed(section1):
    anchor.addnext(el)

# 3) 진행 계획 표 수정
plan = next(t for t in d.tables if t.rows[0].cells[0].text.strip() == '시간')
for row in plan.rows:
    k = row.cells[0].text.strip()
    if k.startswith('0'):
        set_cell(row.cells[2], '임무·산출물·연계 원칙 확인, 발표팀·질의팀 편성')
    elif k.startswith('120'):
        set_cell(row.cells[2], '목적·최종상태·목표 → 임무·실시 → 위험·완화, 판단서 완성 → 발표자료 교환')
    elif k.startswith('180'):
        set_cell(row.cells[1], '발표·질의')
        set_cell(row.cells[2], '1조 발표(15분) → 2조 질의팀 질의·응답(15분) → 2조 발표(15분) → 1조 질의팀 질의·응답(15분) → 종합토의(30분)')

# 4) 제출 전 점검표에 항목 추가
chk = next(t for t in d.tables if t.rows[0].cells[0].text.strip() == '확인')
last = chk.rows[-1]._tr
for txt in ['질의팀은 상대 조의 **연계 확인란**을 기준으로 질의 4개를 준비했는가?',
            '발표팀 4명이 각자 맡은 범위의 **근거 항목**을 답변할 수 있는가?']:
    tr = copy.deepcopy(last)
    last.addnext(tr); last = tr
from docx.table import _Row
for i, txt in enumerate(['질의팀은 상대 조의 연계 확인란을 기준으로 질의 4개를 준비했는가?',
                         '발표팀 4명이 각자 맡은 범위의 근거 항목을 답변할 수 있는가?']):
    row = chk.rows[-2 + i]
    set_cell(row.cells[1], txt)
    for h in row._tr.iter(qn('w:trHeight')):
        h.getparent().remove(h)

# 5) 작성 양식 ③ — 질의 준비 (문서 끝, sectPr 앞)
from docx.enum.text import WD_BREAK
from docx.shared import Cm
pb = d.add_paragraph(); pb.add_run().add_break(WD_BREAK.PAGE); new.append(pb._p)
para('9. 작성 양식 ③ — 질의 준비 (질의팀)', 14.5, NAVY, True, before=10)
p = para('질의 대상: ___조 「________________」     질의팀: ______________________', 12, 'FFFFFF', True)
shade(p._p.get_or_add_pPr(), NAVY)
table(['역할', '대상 항목\n(상대 판단서 쪽·항목)', '상대 자료의 판단 (인용)', '질의', '기대하는 답 · 평가'], [
    ['질의 1\n상황평가', '', '', '', ''],
    ['질의 2\n필요성', '', '', '', ''],
    ['질의 3\n비례성·방책', '', '', '', ''],
    ['질의 4\n위험·연계', '', '', '', ''],
], [1.1, 1.4, 1.8, 1.8, 1.4])
qt = d.tables[-1]
for r in qt.rows[1:]:
    r.height = Cm(3.2)
para('**작성법** — 상대 판단서의 문장을 그대로 인용하고, 그 판단이 방책의 어디로 이어지는지(또는 끊겼는지)를 짚어 질의한다. 예: "③ 종합평가에서 \'추가 도발 가능성 높음\'이라 했는데, 위험·완화방안에 추가 도발 시 대응이 없습니다. 그 경우 2단계는 무엇입니까?"', 10.5)
tail = take()
sect = body.find(qn('w:sectPr'))
for el in tail:
    sect.addprevious(el)

# 제목 줄에 개정 표시
sub = d.paragraphs[1]
sub.runs[-1].text = sub.runs[-1].text + ' | 개정: 발표·질의 임무 분담 반영'
d.save(DST)
print('ok')
