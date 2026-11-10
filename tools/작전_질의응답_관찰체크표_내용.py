sec.header.paragraphs[0].runs[0].text = "작전 과목 질의·응답 관찰 체크표 · 교관용"
sec.top_margin, sec.bottom_margin = Mm(12), Mm(10)
sec.left_margin = sec.right_margin = Mm(14)

def H(t, h):
    for r in t.rows[1:]:
        trPr = r._tr.get_or_add_trPr(); e = OxmlElement("w:trHeight")
        e.set(qn("w:val"), str(int(h * 56.7))); e.set(qn("w:hRule"), "atLeast"); trPr.append(e)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
runs(p, "질의·응답 관찰 체크표", size=17, color=NAVY, bold_all=True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
runs(p, "____그룹   일자 ________   표기: ○ 잘함  △ 보통  ✕ 미흡   (쪽 = 논문 쪽수 인용)", size=10, color=GRAY)

for title in ["질의·응답 Ⅰ   (____소조 질의 → ____소조 응답)", "질의·응답 Ⅱ   (____소조 질의 → ____소조 응답)"]:
    q = doc.add_paragraph(); q.paragraph_format.space_before = Pt(4); q.paragraph_format.space_after = Pt(2)
    runs(q, "▎" + title, size=12, color=BLUE, bold_all=True)
    t = TABLE([
        ["질의", "질문 요지 (키워드)", "근거\n쪽", "질문형\n(✕=주장)", "응답\n논리", "논문·조\n판단 일치", "회피\n없음", "후속\n질의", "★"],
        ["1", "", "", "", "", "", "", "", ""],
        ["2", "", "", "", "", "", "", "", ""],
        ["3", "", "", "", "", "", "", "", ""],
        ["4", "", "", "", "", "", "", "", ""],
        ["5", "", "", "", "", "", "", "", ""],
    ], [0.55, 4.2, 0.65, 0.8, 0.7, 0.85, 0.7, 0.7, 0.45], size=9.5)
    H(t, 8.5)

q = doc.add_paragraph(); q.paragraph_format.space_before = Pt(4); q.paragraph_format.space_after = Pt(2)
runs(q, "▎강평 메모", size=12, color=BLUE, bold_all=True)
t = TABLE([
    ["항목", "메모"],
    ["우수 질의 (누가·무엇)", ""],
    ["우수 응답 (누가·무엇)", ""],
    ["자기 논문·판단과 어긋난 응답", ""],
    ["놓친 후속 질의 (\"그렇다면 ~?\")", ""],
    ["강평 마무리 질문 1개", ""],
], [2.4, 4.6], size=10)
H(t, 8)

BOX(None,
    "**일치**: 응답이 자기 논문·조의 판단에 맞나  |  **회피 없음**: 물은 것에 답했나  |  **후속**: 한 번 더 파고들었나\n"
    "**녹음 시** 질의 구간 종료(45분·90분) 때 녹취 텍스트를 올리면 강평 전에 분석을 받을 수 있음",
    fill="EAF1FB", edge="2E5597")
