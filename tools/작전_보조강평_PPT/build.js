// 작전 강평 보조자료: 혁명인가 진화인가, 그보다 먼저 물어야 할 것 (작전 3·4그룹)
const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/753a0074-9fd2-46a7-8496-f4c30f0cf975_cbf5194e-dd57-4b79-bf4a-bcfc6ca40106/pptx/scripts/apply_theme.js');

const FONT = 'Malgun Gothic';
const THEME = {
  name: 'Situation Room', headFontFace: FONT, bodyFontFace: FONT,
  colors: {
    dk1: '1B1F27', lt1: 'FFFFFF', dk2: '13294B', lt2: 'EEF2F7',
    accent1: 'E0A526', accent2: 'B5482B', accent3: '2B5AA8', accent4: '6B7686',
    accent5: '3F7F5F', accent6: '8A99B0', hlink: '2B5AA8', folHlink: '6B7686',
  },
};
const HEX = THEME.colors;
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = '혁명인가 진화인가, 그보다 먼저 물어야 할 것';
pres.author = '교관 강평 보조자료';
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: 'DARK_TITLE', background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: 'bottom', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.0, w: 11.7, h: 1.4, fontSize: 18, color: C.accent6, valign: 'top', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'DARK_SCENE', background: { color: HEX.dk1 },
  objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.45, w: 11.9, h: 0.9, fontSize: 32, bold: true, color: C.background1, valign: 'middle', align: 'left', margin: 0 }, text: '' } }],
  slideNumber: { x: 12.4, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6 },
});
pres.defineSlideMaster({
  title: 'LIGHT_CONTENT', background: { color: HEX.lt1 },
  objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.45, w: 11.9, h: 0.9, fontSize: 32, bold: true, color: C.text2, valign: 'middle', align: 'left', margin: 0 }, text: '' } }],
  slideNumber: { x: 12.4, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent4 },
});

let STEP = 0, CNT = 0, FX = 'fade';
const step = (n, fx) => { STEP = n; FX = fx || 'fade'; };
const nm = (base) => (STEP > 0 ? `a${STEP}${FX === 'wipe' ? 'w' : 'f'}_${base || 'o'}_${CNT++}` : `${base || 'o'}_${CNT++}`);
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0, objectName: nm('t') }, o));
const IC = (n, c) => `${__dirname}/icons/${n}_${c}.png`;
const icon = (s, n, c, x, y, w, alt) => s.addImage({ path: IC(n, c), x, y, w, h: w, objectName: nm('ic'), altText: alt || n });
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: nm(name) });
const circ = (s, x, y, d, fill) => s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: nm('circ') });
const chevron = (s, x, y, w, h, fill, txt, color) => {
  s.addShape(pres.shapes.CHEVRON, { x, y, w, h, fill: { color: fill }, line: { color: fill }, objectName: nm('chev') });
  T(s, txt, { x: x + 0.3, y, w: w - 0.6, h, fontSize: 15, bold: true, color, align: 'center', valign: 'middle' });
};
const pill = (s, x, y, txt, fill, color, w) => {
  w = w || 2.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.42, rectRadius: 0.2, fill: { color: fill }, line: { color: fill }, objectName: nm('pill') });
  T(s, txt, { x, y, w, h: 0.42, fontSize: 13, bold: true, color, align: 'center', valign: 'middle' });
};
const note = (s, txt, dark) => T(s, txt, { x: 0.7, y: 6.75, w: 11.4, h: 0.35, fontSize: 11, color: dark ? HEX.accent6 : HEX.accent4 });

// ===== 1. 표지
pres.addSection({ title: '도입' });
let s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '도입' });
s.addText('혁명인가 진화인가,\n그보다 먼저 물어야 할 것', { placeholder: 'title' });
s.addText('러우전쟁의 교훈으로 다시 보는 합동 6대 기능과 전영역 통합작전 — 작전 3·4그룹', { placeholder: 'body' });
pill(s, 0.8, 1.2, '교관 강평 보조자료', HEX.accent1, HEX.dk1);
circ(s, 9.3, 0.9, 3.2, '1F3A66');
icon(s, 'GiDeliveryDrone', 'E0A526', 9.85, 1.45, 2.1, '드론');
s.addNotes('오늘 두 논문을 놓고 혁명이냐 진화냐를 토론했습니다. 이 자료는 그 논쟁 뒤에 남는 질문, 곧 "그래서 우리 군은 무엇부터, 어떤 순서로 바꿀 것인가"를 생각해 보게 하는 보조자료입니다.');

// ===== 2. 두 논문
s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '도입' });
s.addText('두 논문은 같은 전장을 보았다', { placeholder: 'title' });
card(s, 0.7, 1.6, 5.8, 3.6, HEX.lt2, 'rev');
pill(s, 1.0, 1.85, '논문 ① 혁명론', HEX.dk2, HEX.lt1);
T(s, 'Zagorodnyuk (2026, Carnegie)', { x: 3.4, y: 1.85, w: 3.0, h: 0.42, fontSize: 13, color: HEX.accent4, valign: 'middle' });
T(s, '"단절이 왔다. 합의를 기다리지 말고 지금 바꿔라"', { x: 1.0, y: 2.5, w: 5.2, h: 0.9, fontSize: 18, bold: true, color: HEX.dk2 });
T(s, [
  { text: '저렴한 정밀 대량전력', options: { bullet: true, breakLine: true } },
  { text: '공중영역의 분절과 기동의 축소', options: { bullet: true, breakLine: true } },
  { text: '네트워크전(델타), 초고속 적응', options: { bullet: true } },
], { x: 1.0, y: 3.45, w: 5.2, h: 1.5, fontSize: 15, color: HEX.dk1, paraSpaceAfter: 6 });
card(s, 6.85, 1.6, 5.8, 3.6, HEX.lt2, 'evo');
pill(s, 7.15, 1.85, '논문 ② 진화론', HEX.accent5, HEX.lt1);
T(s, '정연봉 (2026, 한국국가전략)', { x: 9.55, y: 1.85, w: 3.0, h: 0.42, fontSize: 13, color: HEX.accent4, valign: 'middle' });
T(s, '"일반화하지 말라. 적응을 통해 기동은 돌아온다"', { x: 7.15, y: 2.5, w: 5.2, h: 0.9, fontSize: 18, bold: true, color: HEX.dk2 });
T(s, [
  { text: '상당 기간은 소모전 우위', options: { bullet: true, breakLine: true } },
  { text: '머레이: 평시 혁신 + 전시 적응', options: { bullet: true, breakLine: true } },
  { text: '미군 제파 중심 접근으로 기동 회복', options: { bullet: true } },
], { x: 7.15, y: 3.45, w: 5.2, h: 1.5, fontSize: 15, color: HEX.dk1, paraSpaceAfter: 6 });
step(1);
card(s, 0.7, 5.5, 11.95, 1.0, HEX.dk2, 'common');
icon(s, 'GiStopwatch', 'E0A526', 1.0, 5.65, 0.7, '시계');
T(s, [{ text: '공통 결론  ', options: { color: HEX.accent1, bold: true } }, { text: '적응을 더 빨리 하는 쪽이 이긴다 — 갈리는 것은 "변화의 성격"과 "처방의 순서"다', options: { color: HEX.lt1, bold: true } }], { x: 1.9, y: 5.5, w: 10.6, h: 1.0, fontSize: 16, valign: 'middle' });
s.addNotes('두 논문은 결론이 반대처럼 보이지만, 현상 진단(투명한 전장, 드론의 대량 운용, 기동의 곤란)은 같습니다. 차이는 그것이 영구적 단절인가, 다음 기동전으로 가는 과도기인가입니다. 클릭하면 공통 결론이 나옵니다.');

// ===== 장면 ①
pres.addSection({ title: '장면 ① 투명한 전장' });
s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '장면 ① 투명한 전장' });
s.addText('장면 ①\n숨을 곳이 없는 전장', { placeholder: 'title' });
s.addText('투명한 전투공간, 그리고 소모전의 숫자', { placeholder: 'body' });
circ(s, 9.3, 1.3, 3.2, '1F3A66');
icon(s, 'GiEyeTarget', 'E0A526', 9.85, 1.85, 2.1, '감시');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ① 투명한 전장' });
s.addText('하루 진격 거리 — 1916년 솜 전투보다 느리다', { placeholder: 'title' });
s.addChart(pres.charts.BAR, [{ name: '일일 진격 거리(m)', labels: ['1916 솜 전투', '포크롭스크', '쿠피안스크', '차시우 야르'], values: [80, 70, 23, 15] }], {
  x: 0.7, y: 1.6, w: 7.2, h: 4.8, barDir: 'bar', chartColors: ['8A99B0', 'E0A526', 'E0A526', 'E0A526'], varyColors: true,
  showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: 'FFFFFF', dataLabelFontSize: 16, dataLabelFontFace: '+mn-lt', dataLabelFormatCode: '0"m"',
  catAxisLabelColor: 'FFFFFF', catAxisLabelFontSize: 15, catAxisLabelFontFace: '+mn-lt', catAxisOrientation: 'maxMin',
  valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false, valAxisMaxVal: 95,
  catAxisLineShow: false, barGapWidthPct: 60,
});
step(1);
card(s, 8.4, 1.7, 4.25, 2.15, '24304A', 'st1');
T(s, '115,000 → 8,435㎢', { x: 8.65, y: 1.9, w: 3.9, h: 0.75, fontSize: 28, bold: true, color: HEX.accent1 });
T(s, '러시아 영토 획득: 개전 2개월 vs 2024~2025년 2년간', { x: 8.65, y: 2.7, w: 3.8, h: 0.9, fontSize: 14, color: HEX.lt1 });
step(2);
card(s, 8.4, 4.1, 4.25, 2.15, '24304A', 'st2');
T(s, '약 70%', { x: 8.65, y: 4.3, w: 3.9, h: 0.75, fontSize: 32, bold: true, color: HEX.accent2 });
T(s, '사상자 가운데 드론에 의한 비율 (논문 ②). 논문 ①은 "80% 이상" — 출처가 다르다', { x: 8.65, y: 5.1, w: 3.8, h: 1.0, fontSize: 14, color: HEX.lt1 });
note(s, '※ 출처: 정연봉(2026) Ⅲ장. 수치는 논문이 인용한 공개 자료 기준이며, 인용할 때는 원출처를 함께 밝힐 것.', true);
s.addNotes('일일 진격 거리 비교는 논문 ②가 소모전의 근거로 든 지표입니다. 질문: 이 숫자는 기동전이 죽었다는 증거입니까, 아니면 아직 적응이 끝나지 않았다는 증거입니까? 드론 사상자 비율이 두 논문에서 70%와 80% 이상으로 다르다는 점도 짚어 줍니다. 수치의 출처를 따지는 습관이 질의의 출발점입니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ① 투명한 전장' });
s.addText('킬체인 — 수십 분에서 수 초로', { placeholder: 'title' });
const kc = [['탐지', '광학·열상·전자·음향 센서의 네트워크 융합'], ['판단', '델타: 드론의 정밀성은 사실 "체계 수준 상황인식"의 결과'], ['타격', '드론이 탐지와 동시에 타격, 10~20km 살상 구역']];
kc.forEach(([h, d], i) => {
  step(i + 1, 'wipe');
  const x = 0.7 + i * 4.05;
  chevron(s, x, 1.75, 3.9, 0.9, i === 2 ? HEX.accent2 : HEX.accent1, h, HEX.dk1);
  T(s, d, { x: x + 0.2, y: 2.85, w: 3.5, h: 1.1, fontSize: 15, color: HEX.lt1 });
});
step(4);
card(s, 0.7, 4.3, 11.95, 2.1, '24304A', 'q');
T(s, '생각해 볼 질문', { x: 1.0, y: 4.5, w: 4, h: 0.45, fontSize: 16, bold: true, color: HEX.accent1 });
T(s, [
  { text: '우리 군에서 표적을 탐지해 타격 승인이 나기까지 몇 단계, 몇 분이 걸리는가?', options: { bullet: true, breakLine: true } },
  { text: '"상대도 나를 본다"는 전제에서, 우리 지휘소·집결지·포병은 몇 분을 버티는가?', options: { bullet: true } },
], { x: 1.0, y: 5.0, w: 11.3, h: 1.3, fontSize: 17, color: HEX.lt1, paraSpaceAfter: 8 });
s.addNotes('논문 ①은 센서-슈터 주기가 수 분에서 수 초로 줄었다고 하고, 논문 ②도 킬체인이 수십 분에서 수 초로 줄었다고 씁니다. 학생들에게 우리 부대의 표적 처리 절차를 떠올려 보게 합니다. 기술보다 승인 절차가 더 긴 경우가 많습니다.');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ① 투명한 전장' });
s.addText('같은 현상, 다른 처방', { placeholder: 'title' });
const hdr = (t, f) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: f }, align: 'center', valign: 'middle' } });
const c1 = (t) => ({ text: t, options: { bold: true, color: HEX.dk2, fill: { color: HEX.lt2 }, valign: 'middle' } });
const cc = (t) => ({ text: t, options: { color: HEX.dk1, valign: 'middle' } });
s.addTable([
  [hdr('구분', HEX.dk2), hdr('논문 ① 혁명론', HEX.dk2), hdr('논문 ② 진화론', HEX.accent5)],
  [c1('변화의 성격'), cc('구조적 단절 — 기존 교리가 작동하지 않는다'), cc('순환적 진화 — 참호전 뒤 전격전이 왔듯 기동은 돌아온다')],
  [c1('지금 할 일'), cc('전력구조·획득제도를 지금 바꾼다'), cc('소모전을 버티면서 다음 기동전을 준비한다')],
  [c1('투자 우선'), cc('저가 정밀 대량전력, 신속획득, 여단 중심 혁신'), cc('탄약·동원, 제파식 기동, 적층식 방호')],
  [c1('경계할 오해'), cc('"드론 숫자를 늘리면 된다"'), cc('"기다리면 기동전이 돌아온다"')],
], { x: 0.7, y: 1.6, w: 11.95, colW: [2.2, 4.875, 4.875], rowH: 0.78, fontFace: FONT, fontSize: 15, border: { type: 'solid', pt: 0.75, color: 'C9D2DE' }, margin: [4, 8, 4, 8], objectName: nm('tbl') });
step(1);
T(s, '두 처방은 반대가 아니라 "순서"가 다르다 — 무엇을 먼저 하느냐가 우리 군의 선택이다', { x: 0.7, y: 6.0, w: 11.95, h: 0.5, fontSize: 18, bold: true, color: HEX.accent2 });
s.addNotes('표를 보면 두 처방은 서로를 배제하지 않습니다. 혁명론도 소모전을 부정하지 않고, 진화론도 드론을 부정하지 않습니다. 결국 차이는 우선순위입니다. 앞선 그룹에서도 혁명론 조와 진화론 조의 대응방향이 거의 같았습니다.');

// ===== 장면 ②
pres.addSection({ title: '장면 ② 6대 기능' });
s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '장면 ② 6대 기능' });
s.addText('장면 ②\n합동 6대 기능으로 번역하기', { placeholder: 'title' });
s.addText('교훈은 기능별 질문이 될 때 대응방향이 된다', { placeholder: 'body' });
circ(s, 9.3, 1.3, 3.2, '1F3A66');
icon(s, 'GiPuzzle', 'E0A526', 9.85, 1.85, 2.1, '퍼즐');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ② 6대 기능' });
s.addText('합동 6대 기능 — 기능마다 하나씩 물어보자', { placeholder: 'title' });
const fn = [
  ['GiRadioTower', '지휘통제', '센서-슈터 주기가 분·초 단위, 여단 중심 분권 혁신', '결심 권한을 어느 제대까지 내릴 수 있는가?'],
  ['GiRadarDish', '정보', '델타식 다출처 융합, "상대도 나를 본다"', '우리 상황도는 대대 이하까지 실시간으로 내려가는가?'],
  ['GiCrosshair', '화력', '30~300km 중거리 드론 타격, 제파식 2·3파 사격', '공중우세 없이 차단작전을 할 수단이 있는가?'],
  ['GiBattleTank', '이동 및 기동', '10~20km 살상 구역, "분산 후 집중"', '산악 기동로에서 "집중의 순간"을 어떻게 만드는가?'],
  ['GiShieldReflect', '방호', '요격 비용 비대칭, 재밍을 피하는 광섬유 드론', '싼 드론을 무엇으로, 얼마에 막는가?'],
  ['GiAutoRepair', '지속지원', '부대 자체 정비·개조·소량생산, "수리할 권리"', '대대가 드론을 고쳐 쓸 권한과 부품이 있는가?'],
];
fn.forEach(([ic, h, l, q], i) => {
  step(Math.floor(i / 3) + 1);
  const x = 0.7 + (i % 3) * 4.05, y = 1.55 + Math.floor(i / 3) * 2.6;
  card(s, x, y, 3.85, 2.4, HEX.lt2, 'fn');
  icon(s, ic, '13294B', x + 0.25, y + 0.22, 0.6, h);
  T(s, h, { x: x + 1.0, y: y + 0.22, w: 2.7, h: 0.6, fontSize: 20, bold: true, color: HEX.dk2, valign: 'middle' });
  T(s, l, { x: x + 0.25, y: y + 0.95, w: 3.4, h: 0.65, fontSize: 14, color: HEX.accent4 });
  T(s, q, { x: x + 0.25, y: y + 1.6, w: 3.4, h: 0.7, fontSize: 15, bold: true, color: HEX.accent2 });
});
note(s, '※ 회색: 러우전쟁 교훈(논문 ①·②)   빨간색: 우리 군에 던질 질문', false);
s.addNotes('기능마다 교훈 한 줄, 질문 한 줄입니다. 학생들의 대응방향이 이 질문에 답하고 있는지 확인합니다. 특히 방호와 화력을 혼동하지 않도록 합니다. 적의 포병·미사일·드론을 막는 천궁, L-SAM, 요격드론은 방호이고, 화력은 우리가 적을 타격하는 능력입니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ② 6대 기능' });
s.addText('전영역 통합작전 — 전장은 땅 위에만 있지 않았다', { placeholder: 'title' });
const dom = [
  ['우주', '스타링크 차단(2026.2)이 전장 네트워크를 흔들었다 (논문 ①)'],
  ['전자기', '재밍이 드론을 막자 광섬유 드론이 등장했다'],
  ['공중', '고고도는 방공, 저고도는 드론 — "대역 제한적 공중우세" (논문 ①)'],
  ['사이버·네트워크', '델타: 드론·전자정보·위성·병사 보고를 하나의 상황도로'],
  ['인지', '공보·심리전 — 전과 영상이 곧 여론전의 무기'],
];
dom.forEach(([h, d], i) => {
  step(i + 1, 'wipe');
  const y = 1.6 + i * 0.82;
  pill(s, 0.8, y + 0.12, h, i === 1 ? HEX.accent2 : HEX.accent1, HEX.dk1, 2.4);
  T(s, d, { x: 3.5, y, w: 9.1, h: 0.66, fontSize: 17, color: HEX.lt1, valign: 'middle' });
});
step(6);
card(s, 0.7, 5.8, 11.95, 0.85, '24304A', 'q');
T(s, [{ text: '생각해 볼 질문  ', options: { color: HEX.accent1, bold: true } }, { text: '영역과 영역 사이의 "이음매"는 누가 책임지는가? 우리 합동 지휘통제는 그 이음매를 실시간으로 묶는가?', options: { color: HEX.lt1 } }], { x: 1.0, y: 5.8, w: 11.4, h: 0.85, fontSize: 16, valign: 'middle' });
s.addNotes('논문 ②의 교훈 ⑤는 우주·사이버·전자기로 전장이 확장되면서 진화가 더 빠르고 다양해지고, 새로운 전격전과 마주할 수 있다고 경고합니다. 6대 기능은 영역을 가로지를 때 비로소 통합작전이 됩니다. 기능별 대응을 나열하는 데 그치지 말고 기능과 영역의 연결을 묻게 합니다.');

// ===== 장면 ③ 한반도화
pres.addSection({ title: '장면 ③ 한반도화' });
s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '장면 ③ 한반도화' });
s.addText('장면 ③\n우크라이나를 한반도로 옮기면', { placeholder: 'title' });
s.addText('같은 교훈도 지형·종심·병력·제도에 따라 다른 답이 된다', { placeholder: 'body' });
circ(s, 9.3, 1.3, 3.2, '1F3A66');
icon(s, 'GiMountains', '3F7F5F', 9.85, 1.85, 2.1, '산악');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ③ 한반도화' });
s.addText('그대로 옮길 수 없는 다섯 가지 차이', { placeholder: 'title' });
const kr = [
  ['전선', '약 1,000km 평원', '군사분계선 약 250km, 산악 위주', '분산·드론 운용 거리와 차폐가 달라진다'],
  ['종심', '수도까지 수백 km', '수도권까지 수십 km', '초기 저지의 시간 여유가 거의 없다'],
  ['양상', '수년째 장기 소모전', '단기 고강도 가능성', '소모전을 버틸 탄약·생산은 있는가?'],
  ['병력', '대규모 동원', '병력 감소 추세', '무인화·분산이 선택이 아닌 필수'],
  ['제도', '여단이 정비·개조·소량생산', '소요제기→획득→전력화 수년', '적응 속도의 병목은 어디인가?'],
];
s.addTable([
  [hdr('구분', HEX.dk2), hdr('우크라이나', HEX.dk2), hdr('한반도', HEX.accent5), hdr('그래서 무엇이 달라지나', HEX.accent2)],
  ...kr.map(([a, b, c, d]) => [c1(a), cc(b), cc(c), { text: d, options: { color: HEX.accent2, bold: true, valign: 'middle' } }]),
], { x: 0.7, y: 1.6, w: 11.95, colW: [1.3, 2.75, 3.15, 4.75], rowH: 0.8, fontFace: FONT, fontSize: 15, border: { type: 'solid', pt: 0.75, color: 'C9D2DE' }, margin: [4, 8, 4, 8], objectName: nm('tbl') });
note(s, '※ 지리 수치는 개략값. 발표에서 인용할 때는 공식 자료로 확인할 것.', false);
s.addNotes('앞선 그룹 발표에서 가장 많이 빠진 것이 한반도 적용성이었습니다. 대응방향 하나를 고를 때마다 오른쪽 열을 대입해 보게 합니다. 예를 들어 FPV 드론의 운용 거리와 통신은 산악에서 크게 달라집니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ③ 한반도화' });
s.addText('적응 속도 — 진짜 경쟁 상대는 시간이다', { placeholder: 'title' });
T(s, '고전적 획득 모델', { x: 0.7, y: 1.55, w: 6, h: 0.45, fontSize: 16, bold: true, color: HEX.accent6 });
['소요제기', '입찰', '개발', '시험', '생산'].forEach((t, i) => chevron(s, 0.7 + i * 2.4, 2.1, 2.35, 0.8, '3A4660', t, HEX.lt1));
T(s, '→ 혁신주기: 년 단위', { x: 0.7, y: 3.0, w: 6, h: 0.45, fontSize: 15, color: HEX.accent6 });
step(1, 'wipe');
T(s, '우크라이나의 전시 적응 (논문 ①)', { x: 0.7, y: 3.65, w: 7, h: 0.45, fontSize: 16, bold: true, color: HEX.accent1 });
circ(s, 0.75, 4.2, 1.3, '24304A');
icon(s, 'GiCycle', 'E0A526', 0.95, 4.4, 0.9, '순환');
T(s, [
  { text: '여단이 직접 정비·개조·소량생산 → 전선에서 바로 시험 → 다시 개량', options: { breakLine: true } },
  { text: '혁신주기: 년 → 월 → 주 → 일', options: { bold: true, color: HEX.accent1 } },
], { x: 2.3, y: 4.2, w: 10.3, h: 1.3, fontSize: 17, color: HEX.lt1, valign: 'middle', paraSpaceAfter: 6 });
step(2);
card(s, 0.7, 5.75, 11.95, 0.9, '24304A', 'q');
T(s, [{ text: '생각해 볼 질문  ', options: { color: HEX.accent1, bold: true } }, { text: '우리 군의 적응을 가장 늦추는 것은 기술인가, 제도인가, 문화인가?', options: { color: HEX.lt1 } }], { x: 1.0, y: 5.75, w: 11.4, h: 0.9, fontSize: 17, valign: 'middle' });
s.addNotes('두 논문이 유일하게 완전히 일치하는 지점이 적응 속도입니다. 논문 ①은 고전적 획득 모델이 너무 느리다고 하고, 논문 ②는 머레이를 빌려 평시 혁신과 전시 적응을 가장 잘하는 군대가 이긴다고 합니다. 유도 방향: 소요제기·획득 절차, 부대 정비 권한, 표적 승인 단계, 보고 문화.');

// ===== 정리
pres.addSection({ title: '정리' });
s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '정리' });
s.addText('예산이 100이라면 — 어떻게 나누겠는가', { placeholder: 'title' });
const bud = [
  ['GiFactory', '탄약·동원·생산', '소모전을 버틸 힘'],
  ['GiShieldReflect', '대드론 다층 방호', '싸게, 겹겹이 막기'],
  ['GiRadarDish', '탐지-타격망', '센서·AI·결심 단축'],
  ['GiBattleTank', '제파식 기동·무인 돌파', '기동을 되살리기'],
  ['GiAutoRepair', '획득·정비 제도 개혁', '적응 속도 높이기'],
];
bud.forEach(([ic, h, d], i) => {
  const x = 0.7 + i * 2.43;
  card(s, x, 1.6, 2.28, 3.1, HEX.lt2, 'b');
  circ(s, x + 0.69, 1.8, 0.9, i === 0 ? HEX.dk2 : HEX.lt1);
  icon(s, ic, i === 0 ? 'FFFFFF' : '13294B', x + 0.84, 1.95, 0.6, h);
  T(s, h, { x: x + 0.15, y: 2.85, w: 1.98, h: 0.8, fontSize: 16, bold: true, color: HEX.dk2, align: 'center', valign: 'middle' });
  T(s, d, { x: x + 0.15, y: 3.65, w: 1.98, h: 0.45, fontSize: 14, color: HEX.accent4, align: 'center' });
  T(s, '( ____ )', { x: x + 0.15, y: 4.1, w: 1.98, h: 0.45, fontSize: 16, bold: true, color: HEX.accent2, align: 'center' });
});
step(1);
T(s, [
  { text: '혁명론자라면 어디에 가장 많이 넣겠는가? 진화론자라면?', options: { bullet: true, breakLine: true } },
  { text: '돈이 거의 들지 않는 칸은 어디인가? 그런데 왜 가장 늦게 바뀌는가?', options: { bullet: true } },
], { x: 0.7, y: 5.0, w: 11.9, h: 1.3, fontSize: 18, color: HEX.dk1, paraSpaceAfter: 8 });
s.addNotes('소조별로 2분간 숫자를 적게 한 뒤 비교합니다. 혁명론은 신속획득·저가 대량전력 쪽으로, 진화론은 탄약·동원과 제파식 기동 쪽으로 기울 것입니다. 마지막 칸, 제도 개혁은 예산이 거의 들지 않지만 가장 느리게 바뀐다는 점을 짚습니다. 우리 군의 적응 속도 문제로 연결합니다.');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '정리' });
s.addText('앞 그룹이 자주 빠진 함정 네 가지', { placeholder: 'title' });
const trap = [
  ['드론 숫자로 단순화', '논문 ①이 가장 경계한 오해다. 바뀌는 것은 드론 수가 아니라 작전개념·편성·획득절차다.'],
  ['사업명 나열', 'KJCCS, L-SAM, 425사업 … 사업마다 "그래서 무엇이 달라지는가"를 붙여야 대응방향이 된다.'],
  ['기능 분류 혼동', '적의 포병·드론을 막는 천궁·요격드론은 "화력"이 아니라 "방호"다.'],
  ['맡은 논문의 고유 처방 누락', '진화론을 맡았다면 결론은 "기동을 어떻게 되살리나"여야 한다. 처방이 상대 논문과 같아지면 토론이 서지 않는다.'],
];
trap.forEach(([h, d], i) => {
  step(i + 1, 'wipe');
  const y = 1.6 + i * 1.22;
  circ(s, 0.75, y + 0.1, 0.7, HEX.dk2);
  T(s, String(i + 1), { x: 0.75, y: y + 0.1, w: 0.7, h: 0.7, fontSize: 22, bold: true, color: HEX.lt1, align: 'center', valign: 'middle' });
  T(s, h, { x: 1.75, y, w: 3.4, h: 0.9, fontSize: 19, bold: true, color: HEX.dk2, valign: 'middle' });
  T(s, d, { x: 5.2, y, w: 7.45, h: 0.9, fontSize: 15, color: HEX.dk1, valign: 'middle' });
});
note(s, '※ 작전 2그룹 발표·강평에서 나온 보완점을 일반화한 것', false);
s.addNotes('2그룹 강평에서 실제로 나온 보완점입니다. 3·4그룹은 발표 준비 단계에서 이 네 가지를 미리 점검하도록 안내할 수 있습니다. 출처 없는 수치, 정치 발언 수치(사상자 80% 이상 vs 약 70%)도 함께 짚어 줍니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '정리' });
s.addText('여러분의 판단에 다시 묻는다', { placeholder: 'title' });
const qs = [
  ['Q1', '혁명이든 진화든 처방이 같다면, 두 논문의 차이는 어디에서 드러나야 하는가?'],
  ['Q2', '"상당 기간 소모전 우위"라면, 내일 전쟁이 나면 우리는 무엇으로 버티는가?'],
  ['Q3', '오늘 찾은 개선점 하나가 실제 대대에 적용되기까지 몇 년이 걸리는가?'],
];
qs.forEach(([k, q], i) => {
  step(i + 1);
  const y = 1.7 + i * 1.4;
  card(s, 0.7, y, 11.95, 1.15, '24304A', 'q');
  T(s, k, { x: 1.0, y, w: 1.0, h: 1.15, fontSize: 24, bold: true, color: HEX.accent1, valign: 'middle' });
  T(s, q, { x: 2.1, y, w: 10.3, h: 1.15, fontSize: 19, color: HEX.lt1, valign: 'middle' });
});
s.addNotes('Q1 유도: 우선순위와 순서. Q2 유도: 탄약·동원·생산, 그리고 초기 저지. Q3 유도: 그 연수가 곧 우리 군의 적응 속도. 종합토의에서 한 문항씩 던집니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '정리' });
s.addText('더 빨리 보고, 더 빨리 결심하고,\n더 빨리 고치는 군대가 이긴다', { placeholder: 'title' });
s.addText('혁명인지 진화인지는 역사가 판정한다. 우리에게 필요한 것은 판정이 아니라 순서다.', { placeholder: 'body' });
circ(s, 9.6, 0.7, 2.6, '1F3A66');
icon(s, 'GiChessKnight', 'E0A526', 10.1, 1.2, 1.6, '기사');
s.addNotes('마무리 멘트: 두 논문은 같은 결론에 도착했습니다. 더 빨리 보고, 더 빨리 결심하고, 더 빨리 고치는 군대가 이깁니다. 오늘 찾은 개선점 하나가 실제 부대에 적용되기까지 걸리는 시간, 그 숫자가 바로 우리 군의 적응 속도입니다.');

(async () => {
  const out = process.argv[2] || 'deck.pptx';
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('wrote', out);
})();
