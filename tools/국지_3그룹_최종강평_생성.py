"""국지 3그룹 교관 최종 강평 docx 생성."""
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

def good_bad(rows, kind):
    """rows: (요지, 근거 쪽·문장, 의미/바로잡기)"""
    head = ['잘된 점', '어디서 (쪽 · 문장)', '왜 잘됐나'] if kind == 'g' else ['잘못된 점', '어디서 (쪽 · 문장)', '이렇게 바로잡는다']
    t = table(head, rows, [3.2, 4.2, 4.6], first_fill='E8F3EC' if kind == 'g' else 'FBECE8', size=9.5)
    return t

# ---------- 머리 ----------
para('합동국지도발대비작전 위기조치 실습 — 국지 3그룹 교관 강평(최종)', 16, NAVY, True)
para('1조 「역린(逆鱗)」(초안 「유나이티드 리졸브」) · 2조 「결연한 의지」(초안 「파수」)', 11, GREY)
para('교관용 · 학생 배포 금지 · 강평 15분 · 초안(11.10) 대비 최종안(11.11)', 9.5, GREY)

h1('한눈에 보기')
table(['구분', '1조 「역린」', '2조 「결연한 의지」'], [
    ['지금 하는 것', '한·미 전략자산 배치·무력시위(강화도 서방), 정보작전·SC', '해상 도발원점(VSV)·C전대본부·해상 지원세력 제한 타격'],
    ['추가도발 시', '대응수준 A·B·C에 따른 지정 표적 타격', '결심조건표(PIR·FFIR) → 수준 결정, 지상 표적은 보류 중'],
    ['초안 대비', '조건부 대응을 임무·대응개념에 명시, 상황개요 신설', '평가 결론·작전명·작전개념을 함께 고쳐 "지금 제한 타격"으로 통일'],
    ['한 줄 평', '**억제의 논리는 섰다. 남의 문장과 언론통제가 발목을 잡는다.**', '**비례성의 상한선을 그었다. 최종상태와 타격 시점이 과하다.**'],
], [1.6, 5.2, 5.2])
box(['**강평의 축 — 같은 비례성 판단, 다른 결론.** 두 조 모두 "1차 대응은 충분, 추가 도발 의도는 식별"에서 출발했다. ',
     '1조는 "지금 치면 확전 빌미 → 억제", 2조는 "추가 공격 방지를 위해 지금 제한 타격"으로 갈렸다. 둘 다 답이 될 수 있다. 판단 근거가 방책 끝까지 이어졌는지를 본다.'], 'FFF8E8')
h2('강평 순서 (15분)')
para('총평·한눈에 보기 2분 → 1조 5분 → 2조 5분 → 종합 질문·마무리 3분', 10)

# ---------- 1조 ----------
pagebreak()
h1('1조 「역린」 — 억제를 택한 조')
good_bad([
    ['**물리적 대응을 하지 않는 이유를 썼다**', '방책 1/9 비례성: "1차 도발 이후 추가 도발행위가 없었으므로 직접 타격 시 추가도발·확전 빌미 → 타격이 아닌 억제수단", 결과 "유사시 추가 대응 필요"', '필요성 충족 ≠ 즉시 타격. 대응하지 않는 선택도 근거로 적어야 한다는 원칙을 정확히 보여줌. 초안의 "필요/불필요" 혼재도 "유사시"로 정리.'],
    ['**조건부 대응을 끝까지 연결했다**', '임무 "추가도발 시 자위권 행사", 대응개념 ② "필요성·비례성 판단 하 단호한 대응, 도발 이전 원상회복", 작전개념 마 "도발수준별 지정 표적, 대응수준 A·B·C"', '"억제가 실패한 다음 날"을 대응수준(A·B·C)으로 미리 정함. 최종상태 "위기고조 이전 단계 환원"과도 일치.'],
    ['**사실에서 출발했다**', '평가 1/6 주요상황 요약(3.23~4.19 시간대별), 3/6 "전작권 전환 후 한·미 연합지휘 실효성 시험", 4/6 작전적 의도 ⑤', '초안에 없던 상황개요와 "전작권 이후 첫 시험대"라는 새로운 의도 해석으로 판단의 바닥을 깔았다.'],
], 'g')
good_bad([
    ['**언론통제·언론 자유 제한 건의가 그대로 남았다**', '방책 4/9 라 ① "언론통제(뉴스·SNS·유튜브)", 5/9 바 ⑤ "언론통제(보이콧)", ⑥ "헌법·국가보안법 등 언론의 자유 일시 제한"', '합참 권한 밖이고, 이 조가 SC로 얻으려는 국내외 정당성·국민 지지와 정면 충돌. → 삭제하고 "사실 기반 신속 공개, 허위정보 정정"으로.'],
    ['**과업이 이 조의 작전개념과 따로 논다**', '6/9 "3. 과업"과 7~8/9 "4. 예하부대 과업"이 중복(번호 4도 협조지시와 중복). 7~8/9는 "F군단·D함대사 지원시설 예비타격", 무력시위 참가 미 7함대·공작사 11·20전비 과업 없음', '다른 조 문서의 문구가 섞인 흔적. 방책의 중심인 "전략자산 배치·무력시위"를 수행할 부대 과업부터 다시 쓴다. 과업은 하나로.'],
    ['**결론 문장과 위험표가 억제 방책을 받쳐 주지 않는다**', '종합평가 6/6 "강력한 군사적 대응 필요", 방책 1/9 결론 "적 의도를 분쇄하는 충분하고 단호한 대응" / 위험표 9/9는 2개 항목뿐', '말은 "강력한 대응", 수단은 억제 → "강력한 억제와 조건부 대응"으로 맞춘다. 무력시위가 2010년처럼 도발 구실이 될 위험, 언론통제 역풍을 위험표에 넣는다.'],
], 'b')
h2('강평 멘트')
box(['"1조는 \'지금 치지 않는 이유\'를 쓴 조입니다. 필요성이 충족돼도 지금은 억제가 맞다는 판단, 그리고 실패하면 A·B·C 수준으로 간다는 설계 — 이것이 비례성입니다. ',
     '그런데 같은 방책 안에 \'언론의 자유 일시 제한\'이 있습니다. 국민의 지지를 얻겠다면서 국민의 입을 막을 수 있습니까? ',
     '그리고 과업 두 쪽은 여러분의 무력시위를 수행하지 않습니다. 남의 문장은 여러분의 판단을 책임지지 않습니다."'])

# ---------- 2조 ----------
pagebreak()
h1('2조 「결연한 의지」 — 제한 타격을 택한 조')
good_bad([
    ['**평가와 방책을 한 방향으로 고쳤다**', '초안: 종합평가 "현 단계는 강력한 억제와 확전 통제" ↔ 방책 4.20 06:00 시간표 타격 / 최종: 종합평가 6쪽 "자위권 차원의 강력한 대응", 작전명 「파수」→「결연한 의지」, 결정적작전 "원점·지휘·지원세력 타격"', '초안의 가장 큰 모순을 결론 하나 바꾸는 데 그치지 않고 평가·작전명·작전개념까지 함께 고쳤다. 고친다는 것은 이렇게 하는 것.'],
    ['**비례성의 상한선을 표로 그었다**', '9쪽 비례성 "공격 저지 및 추가 공격 방지를 위한 무력대응", 15쪽 표적표: 해상 VSV·C전대본부·지원세력 타격, 공중표적 전부·D함대사·E사단·F군단 미타격, 44포병여단 "대기(보류)"', '비례성을 "우리 신뢰도·이익"(초안)에서 "적의 위협"으로 바로잡았다. 어디까지 치고 어디서 멈추는지를 보여준 유일한 조.'],
    ['**출구와 결심 기준을 함께 뒀다**', '12쪽 합참 과업 ⑤ "군사정전위원회·남북 통신채널로 대북 소통창구 유지", 17쪽 자위권 행사 결심조건표(PIR 3·FFIR 6 → A·B·C·No-Go)', '타격하면서도 위기를 끝낼 통로를 열어 둠. 결심조건표는 3그룹에서 가장 완성도 높은 산출물.'],
], 'g')
good_bad([
    ['**최종상태가 과하다**', '10쪽 최종상태 1) "원점·지원·지휘세력은 파괴되어 추가 도발이 불가능해야", 4) "국제 무역·금융체계 내 한국의 피해는 발생하지 않아야"', '국지도발 대응으로 북한의 도발 능력을 없앨 수 없고, 달성 여부도 확인할 수 없다. 작전목표 4) "도발 이전상태 회복"과도 충돌. → "추가 도발 의지 상실, 위기 이전 상태 회복"으로.'],
    ['**타격 시점의 자위권 근거가 약하다**', '과업 12쪽 "4월 19일 14:00부 1차 지휘세력·해상 원점 화력타격" ↔ 근거는 9쪽 "추가 도발의도 식별"뿐. 결심조건표는 방책 본문에 연결 안 됨', '도발이 끝난 뒤의 타격은 보복으로 읽힐 수 있다. 결심조건표의 PIR ②(추가 도발 임박 징후) 충족을 근거로 명시해야 자위권이 선다. 추가 도발 시 C수준으로 가는 기준도 필요.'],
    ['**문서 안의 앞뒤가 맞지 않는다**', '16쪽 타격수단표 "결심" 칸에 D함대사·공군사령부·F군단·A비행사(모두 미타격), 작전목표 1)(지금 타격)과 3)(추가 도발 시 타격) 구분 모호, 위험표 "현장지휘관 주관 자위권 대응"', '표적표에서 지운 표적은 타격수단표에서도 지운다. "지금 칠 표적"과 "추가 도발 시 칠 표적(보류)"을 나눠 적는다.'],
], 'b')
h2('강평 멘트')
box(['"2조는 \'지금 친다, 그러나 여기까지만\'을 표로 보여준 조입니다. 공중표적을 모두 내려놓고 지상 표적을 보류한 그 판단이 비례성입니다. ',
     '그런데 최종상태는 \'추가 도발이 불가능해야\'입니다. VSV와 전대본부를 쳐서 북한의 도발 능력이 없어집니까? ',
     '그리고 4월 19일 오후 2시의 타격은 자위권입니까, 보복입니까? 여러분이 만든 결심조건표가 그 답을 갖고 있습니다. 방책에 연결하십시오."'])

# ---------- 종합 ----------
pagebreak()
h1('종합 질문 3')
table(['#', '질문', '유도 방향'], [
    ['Q1', '두 조는 같은 사실(1차 대응 충분, 추가 도발 의도 식별)에서 다른 결론(억제 / 제한 타격)에 이르렀다. 무엇이 갈랐는가?', '"추가 공격의 임박성"을 어떻게 봤는가. 임박하면 자위권 타격, 불확실하면 억제와 조건부 대응. 결심조건표가 그 판단 도구.'],
    ['Q2', '1조에게: 무력시위가 2010년 연평도처럼 도발의 구실이 되면? / 2조에게: 우리의 제한 타격에 북한이 역대응하면 다음 날은?', '두 방책 모두 "다음 날"의 위험을 위험표에 써야 한다. 훈련 사전 통보·장소 선정 / 대응수준 C 기준·작전적 정지.'],
    ['Q3', '정당성은 무엇으로 얻는가? (1조의 언론통제, 2조의 타격 시점)', '정당성은 통제가 아니라 사실 공개와 자위권 요건 충족에서 나온다.'],
], [0.6, 5.6, 5.8])
h1('과거 사례 연결 (강평보조 PPT)')
table(['사례', '연결'], [
    ['판문점 1976', '1조처럼 압도적 전력을 전개하되, **실제로 나무를 자르는 제한된 행동**으로 메시지를 완성. → 1조: 여러분의 "나무"는 무엇인가?'],
    ['연평도 2010', '해상사격훈련이 포격의 구실이 됐고, 대응은 규칙상 비례였지만 원점 제압·재발 방지는 못 했다. → 1조: 훈련 장소·통보 설계 / 2조: 제한 타격 이후의 다음 수.'],
], [2.0, 10])
h1('마무리 멘트')
box(['"3그룹은 처음으로 두 조의 결론이 갈린 그룹입니다. 한 조는 참았고, 한 조는 제한적으로 쳤습니다. 어느 쪽도 틀리지 않았습니다. ',
     '틀리는 것은 결론이 아니라 연결입니다. 1조는 남의 문장과 언론통제를 걷어내고, 2조는 최종상태를 낮추고 타격 근거를 결심조건표에 묶으십시오. ',
     '분석에서 쓴 말은 방책이 책임집니다."'], 'FFF8E8')
para('※ 학생 실명 미기재. 상세 근거: 국지/3그룹/3그룹_분석자료.docx', 9, GREY, before=6)
d.save(DST)
print('saved', DST)
