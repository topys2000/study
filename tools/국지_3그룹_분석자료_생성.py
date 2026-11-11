"""국지 2그룹 교관 최종 강평(초안 대비 비교) docx 생성."""
import re, sys
import docx
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

DST = sys.argv[1]
NAVY, BLUE, GREY, RED, GREEN = '13294B', '2B5AA8', '5A6475', 'B5482B', '2E7D4F'
d = docx.Document()
for s in d.sections:
    s.left_margin = s.right_margin = Cm(2.0); s.top_margin = s.bottom_margin = Cm(1.8)
st = d.styles['Normal']; st.font.name = '맑은 고딕'; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
st.paragraph_format.space_after = Pt(3); st.paragraph_format.line_spacing = 1.15
W = 17.0  # usable width cm


def runs(p, text, size=10.5, color=None, bold_all=False):
    for i, part in enumerate(re.split(r'\*\*', text)):
        if not part: continue
        r = p.add_run(part); r.font.size = Pt(size); r.font.name = '맑은 고딕'
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '맑은 고딕')
        r.bold = bold_all or (i % 2 == 1)
        if color: r.font.color.rgb = RGBColor.from_string(color)
    return p


def para(text='', size=10.5, color=None, bold=False, before=None, style=None, align=None):
    p = d.add_paragraph(style=style)
    if before: p.paragraph_format.space_before = Pt(before)
    if align: p.alignment = align
    return runs(p, text, size, color, bold)


def h1(t):
    p = para(t, 14, NAVY, True, before=12)
    if PB[0]:
        p.paragraph_format.page_break_before = True; PB[0] = False
    return p
def h2(t): return para(t, 12, BLUE, True, before=6)
def bullet(t): return para(t, style='List Bullet')


def shade(pr, fill):
    s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), fill); pr.append(s)


def box(lines, fill='EEF3FA'):
    p = d.add_paragraph()
    for i, ln in enumerate(lines):
        if i: p.add_run().add_break()
        runs(p, ln)
    shade(p._p.get_or_add_pPr(), fill)
    p.paragraph_format.space_after = Pt(6)
    return p


MARK = {'○': GREEN, '△': 'B07D00', '✕': RED}


def table(header, rows, widths, first_fill='EEF3FA', size=9.5):
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    tot = sum(widths)
    for ri, row in enumerate([header] + rows):
        for ci, txt in enumerate(row):
            c = t.cell(ri, ci); c.width = Cm(W * widths[ci] / tot)
            p = c.paragraphs[0]
            for k, ln in enumerate(txt.split('\n')):
                if k: p.add_run().add_break()
                col = 'FFFFFF' if ri == 0 else MARK.get(ln.strip()[:1]) if ln.strip()[:1] in MARK and len(ln.strip()) <= 3 else None
                runs(p, ln, size, col, ri == 0 or (col is not None and ri > 0))
            if ri == 0: shade(c._tc.get_or_add_tcPr(), NAVY)
            elif ci == 0 and first_fill: shade(c._tc.get_or_add_tcPr(), first_fill)
    t.autofit = False
    for ci, wd in enumerate(widths):
        t.columns[ci].width = Cm(W * wd / tot)
    tw = OxmlElement('w:tblW'); tw.set(qn('w:w'), str(int(W * 567))); tw.set(qn('w:type'), 'dxa')
    for old in t._tbl.tblPr.findall(qn('w:tblW')): t._tbl.tblPr.remove(old)
    t._tbl.tblPr.append(tw)
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); t._tbl.tblPr.append(lay)
    b = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side); e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '4'); e.set(qn('w:color'), 'BFBFBF'); b.append(e)
    t._tbl.tblPr.append(b)
    d.add_paragraph()
    return t


PB = [False]
def pagebreak():
    PB[0] = True

para('국지 3그룹 분석자료 — 초안 대비 최종안 대조', 16, NAVY, True)
para('교관 참고용 · 강평서(교관강평_최종_1조2조.docx)의 상세 근거 · ○ 반영 / △ 일부 / ✕ 미반영', 9.5, GREY)

h1('1조 「유나이티드 리졸브」 → 「역린」')
h2('최종안 개요')
table(['항목', '내용'], [
    ['임무', '단호한 대응능력·의지 현시로 추가 도발 의지 분쇄, 전면전 확대 방지, 추가도발 시 자위권 행사'],
    ['최종상태', '추가도발 의지 분쇄·위기고조 이전 단계 환원 / 접적지역 경계태세 상향 / 공보·SC로 정당성·지지 확보'],
    ['작전개념', '연합위기관리(7함대 항모·핵잠·B-2), 한·미 전략자산 배치·무력시위(강화도 서방), 정보작전, SC, 추가도발 시 대응수준 A·B·C 지정 표적 타격, 정부건의'],
    ['새로 생긴 것', '상황개요(3.23~4.19), 전작권 전환 후 연합지휘 실효성 시험이라는 의도, 비례성 결과 "유사시 추가 대응 필요"'],
], [1.8, 10])
h2('초안 지적 → 최종 대조')
table(['초안에서 지적한 점', '최종안', '판정'], [
    ['비례성 결론 "추가 대응 불필요 / 필요" 혼재', '"추가 대응 불필요(타격) / 유사시 추가 대응 필요(억제수단)"로 정리', '○'],
    ['결론 "충분하고 단호한 대응으로 상황 종결" ↔ 수단은 무력시위·담화', '임무·대응개념에 "추가도발 시 자위권 행사" 명시. 그러나 종합평가 "강력한 군사적 대응 필요", 비례성 결론 문장은 그대로', '△'],
    ['연평도 남방 실사격훈련 = 2010년 포격 구실', '장소를 강화도 서방으로, 표현을 "전략자산 배치"로 바꿈. 위험표에는 여전히 이 위험 없음', '△'],
    ['언론통제·"언론의 자유 일시 제한" 건의', '4/9 라 ①, 5/9 바 ⑤⑥ 그대로', '✕'],
    ['과업·협조·위험이 다른 조 문구, 실사격 참가부대 과업 없음', '"3. 과업"(6/9)과 "4. 예하부대 과업"(7~8/9) 중복, 7~8/9는 지원시설 예비타격 등 다른 조 문구. 미 7함대·공작사 11·20전비 과업 없음. 번호 4 중복', '✕'],
    ['작전개념 기능별 나열, 결정적·여건조성 구분 없음', '같음', '✕'],
    ['최종상태 모호', '"위기고조 이전 단계 환원" 등 3개로 구체화', '○'],
], [4.3, 6.8, 0.8])

pagebreak()
h1('2조 「파수」 → 「결연한 의지」')
h2('최종안 개요')
table(['항목', '내용'], [
    ['임무', '자위권 차원 계획된 화력계획으로 원점·지휘세력·핵심 지원시설 정밀타격, 대응능력·의지 현시 → 추가 도발 의지 분쇄, 전면전 확대 방지'],
    ['최종상태', '원점·지원·지휘세력 파괴로 추가 도발 불가, 아군 피해 없음·지휘통제 유지, 국민 일상, 국제 지지·무역금융 피해 없음'],
    ['작전개념', '결정적: 각 군 원점·지휘·지원세력 타격 / 여건조성: 대비태세·ISR·NLL 해상우세·SC·전략자산 가시화 / 지속지원'],
    ['표적', '해상 VSV·C전대본부·지원세력 타격, 공중표적 전부·D함대사·E사단·F군단 미타격, 44포병여단·지상 지원세력 대기(보류)'],
    ['새로 생긴 것', '국내상황 PESII, 의도별 연도 근거, 예상 양상↔의도 판단근거 연결, 대북 소통창구 유지'],
], [1.8, 10])
h2('초안 지적 → 최종 대조')
table(['초안에서 지적한 점', '최종안', '판정'], [
    ['종합평가 "현 단계는 억제·확전 통제" ↔ 방책 시간표 타격', '평가 결론을 "강력한 대응"으로, 작전명·작전개념까지 함께 수정', '○'],
    ['단계 전환이 시계 기준, 결심조건표와 단절', '3단계 시간표 삭제. 과업은 "4.19 14:00부 타격"으로 시각 지정, 결심조건표(B수준)는 본문과 여전히 단절', '△'],
    ['비례성을 "대군 신뢰도·이익"으로 판단', '"공격 저지 및 추가 공격 방지를 위한 무력대응"으로 바꿈', '○'],
    ['작전명 「파수」(감시) ↔ 공세적 타격', '「결연한 의지」로 변경', '○'],
    ['표적표 F사단 제외 ↔ 공작사 "F군단 타격", A전투비행사단 타격', '공중표적 전부 미타격, F군단 미타격, 공작사는 비상대기로 정리. 그러나 타격수단표 "결심" 칸에 미타격 표적 잔존', '△'],
    ['"확전통제는 문제" ↔ "확전통제 필요" 긴장', '"우리군 현실태 및 문제점" 쪽 삭제로 해소', '○'],
    ['(최종에서 새로 생긴 문제) 최종상태', '"추가 도발이 불가능해야", "무역·금융 피해 없어야" — 과도·확인 불가', '✕'],
], [4.3, 6.8, 0.8])

h1('두 조 비교 — 같은 판단, 다른 결론')
table(['판단 요소', '1조 「역린」', '2조 「결연한 의지」'], [
    ['1차 대응 평가', '충분(비례성 충분)', '충분, 그러나 추가 도발 의도'],
    ['추가 도발 의도', '식별(지상·해상·공중 징후)', '식별(같은 징후)'],
    ['결론', '직접 타격은 확전 빌미 → 억제수단, 유사시 A·B·C', '추가 공격 방지 위해 지금 제한 타격, 지상은 보류'],
    ['가른 것', '임박성을 "불확실"로 봄', '임박성을 "충분"으로 봄 → 근거(결심조건표 PIR②)를 방책에 연결해야'],
    ['과거 사례', '판문점 1976(압도적 전개 + 제한 행동)', '연평도 2010 이후 원점·지원·지휘세력 대응 원칙'],
], [2.2, 4.9, 4.9])
para('※ 원본: 국지/3그룹/학생초안_사진, 학생최종안_사진 · 초안 메모: 1조·2조_초안_분석메모.md', 9, GREY, before=6)
d.save(DST)
print('saved', DST)
