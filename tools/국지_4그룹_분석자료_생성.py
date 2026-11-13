"""국지 4그룹 분석자료 docx: md 메모들을 그대로 옮김. 사용: python3 이 파일 출력.docx md1 md2 ..."""
import sys, re
exec(open(__file__.replace('국지_4그룹_분석자료_생성.py','국지_4그룹_최종강평_생성.py')).read().split('# ---------- 머리')[0])
para('합동국지도발대비작전 위기조치 실습 — 국지 4그룹 분석자료', 16, NAVY, True)
para('교관용 · 강평서의 상세 근거 · 쪽 표기: 초안 p1~p10 / 1조 최종 1~6장 / 2조 최종 F1~F10', 9.5, GREY)
first = True
for md in sys.argv[2:]:
    lines = open(md).read().split('\n'); i = 0
    while i < len(lines):
        ln = lines[i]
        if ln.startswith('# '):
            if not first: pagebreak()
            first = False; h1(ln[2:])
        elif ln.startswith('## '): h2(ln[3:])
        elif ln.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not set(''.join(cells)) <= set('-: '): rows.append(cells)
                i += 1
            n = len(rows[0]); wd = {5: [0.5, 2.8, 0.9, 5.2, 1.6], 7: [0.4, 2.2, 2.0, 2.4, 0.7, 2.4, 2.4], 6: [0.5, 1.8, 2.6, 2.6, 2.4, 2.4], 4: [2.6, 1.2, 4, 4]}.get(n, [1] * n)
            table(rows[0], rows[1:], wd, size=8.5); continue
        elif ln.startswith('- '): bullet(ln[2:])
        elif ln.strip(): para(ln, 10)
        i += 1
d.save(DST); print('saved', DST)
