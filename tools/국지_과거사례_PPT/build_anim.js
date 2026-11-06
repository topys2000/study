// 국지 강평 보조자료: 비례성만으로 위기는 끝나는가 — 판문점 1976 · 연평도 2010
const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/753a0074-9fd2-46a7-8496-f4c30f0cf975_cbf5194e-dd57-4b79-bf4a-bcfc6ca40106/pptx/scripts/apply_theme.js');

const FONT = 'Malgun Gothic';
const THEME = {
  name: 'Situation Room',
  headFontFace: FONT, bodyFontFace: FONT,
  colors: {
    dk1: '1B1F27', lt1: 'FFFFFF', dk2: '13294B', lt2: 'EEF2F7',
    accent1: 'E0A526', accent2: 'B5482B', accent3: '2B5AA8', accent4: '6B7686',
    accent5: '3F7F5F', accent6: '8A99B0', hlink: '2B5AA8', folHlink: '6B7686',
  },
};
const HEX = THEME.colors;
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = '비례성만으로 위기는 끝나는가';
pres.author = '교관 강평 보조자료';
const C = pres.SchemeColor;

// ---------- layouts
pres.defineSlideMaster({
  title: 'DARK_TITLE', background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: 'bottom', align: 'left', margin: 0 }, text: '' } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.8, y: 4.0, w: 11.7, h: 1.4, fontSize: 18, color: C.accent6, valign: 'top', margin: 0 }, text: '' } },
  ],
});
pres.defineSlideMaster({
  title: 'DARK_SCENE', background: { color: HEX.dk1 },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.45, w: 11.9, h: 0.9, fontSize: 32, bold: true, color: C.background1, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
  ],
  slideNumber: { x: 12.4, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6 },
});
pres.defineSlideMaster({
  title: 'LIGHT_CONTENT', background: { color: HEX.lt1 },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.45, w: 11.9, h: 0.9, fontSize: 32, bold: true, color: C.text2, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
  ],
  slideNumber: { x: 12.4, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent4 },
});

let STEP = 0, CNT = 0, FX = 'fade';
const step = (n, fx) => { STEP = n; FX = fx || 'fade'; };
const nm = (base) => (STEP > 0 ? `a${STEP}${FX === 'wipe' ? 'w' : 'f'}_${base || 'o'}_${CNT++}` : `${base || 'o'}_${CNT++}`);
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0, objectName: nm('t') }, o));
const IC = (n, c) => `${__dirname}/icons/${n}_${c}.png`;
const icon = (s, n, c, x, y, w, alt) => s.addImage({ path: IC(n, c), x, y, w, h: w, objectName: nm('ic'), altText: alt || n });
const arrow = (s, x1, y1, x2, y2, color, name) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.01), h: Math.max(Math.abs(y2 - y1), 0.01), flipH: x2 < x1, flipV: y2 < y1, line: { color, width: 3.5, endArrowType: 'triangle' }, objectName: nm(name || 'arr') });
const card = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: nm(name) });
const pill = (s, x, y, txt, fill, color, name) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.9, h: 0.42, rectRadius: 0.2, fill: { color: fill }, line: { color: fill }, objectName: name });
  T(s, txt, { x, y, w: 1.9, h: 0.42, fontSize: 13, bold: true, color, align: 'center', valign: 'middle' });
};

// 1. 표지
pres.addSection({ title: '도입' });
let s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '도입' });
s.addText('비례성만으로 위기는 끝나는가', { placeholder: 'title' });
s.addText('두 개의 실제 위기로 다시 보는 오늘의 위기평가판단서 — 판문점 1976 · 연평도 2010', { placeholder: 'body' });
pill(s, 0.8, 1.4, '교관 강평 보조자료', HEX.accent1, HEX.dk1, 'tag');
s.addNotes('오늘 실습에서 여러분은 필요성과 비례성을 기준으로 위기를 평가하고 방책을 세웠습니다. 이 자료는 그 기준이 실제 역사 속 위기에서 어떻게 작동했는지, 또는 작동하지 않았는지를 두 장면으로 보여줍니다.');

// 2. 질문
s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '도입' });
s.addText('오늘 여러분은 이 저울로 판단했습니다', { placeholder: 'title' });
card(s, 0.7, 1.7, 5.7, 4.6, HEX.lt2, 'scale-card');
icon(s, 'GiScales', '13294B', 5.3, 1.95, 0.85, '저울');
T(s, '필요성', { x: 1.1, y: 2.0, w: 4.9, h: 0.6, fontSize: 24, bold: true, color: HEX.dk2 });
T(s, '합참이 나서야 할 만큼 상황이 악화될 가능성이 있는가', { x: 1.1, y: 2.6, w: 4.9, h: 0.9, fontSize: 16, color: HEX.dk1 });
T(s, '비례성', { x: 1.1, y: 3.8, w: 4.9, h: 0.6, fontSize: 24, bold: true, color: HEX.dk2 });
T(s, '도발을 저지·격퇴·중지시키는 목적의 범위 안에서 대응하는가', { x: 1.1, y: 4.4, w: 4.9, h: 0.9, fontSize: 16, color: HEX.dk1 });
step(1); T(s, '그렇다면 역사 속 지휘관들도\n이 저울로 위기를 끝냈을까?', { x: 7.0, y: 2.0, w: 5.6, h: 1.6, fontSize: 28, bold: true, color: HEX.accent2 });
step(2); T(s, [
  { text: '장면 ①  1976.8.18 판문점 — 나무 한 그루', options: { bullet: true, breakLine: true } },
  { text: '장면 ②  2010.11.23 연평도 — 13분', options: { bullet: true } },
], { x: 7.0, y: 4.0, w: 5.6, h: 1.4, fontSize: 18, color: HEX.dk1, paraSpaceAfter: 10 });
s.addNotes('질문을 먼저 던지고 답을 미룹니다. 학생들에게 각 장면마다 직접 판단하게 한 뒤 실제 결과를 공개합니다.');

// ===== 장면 1
pres.addSection({ title: '장면 ① 판문점 1976' });
s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '장면 ① 판문점 1976' });
s.addText('장면 ①  나무 한 그루', { placeholder: 'title' });
s.addText('1976년 8월 18일, 판문점 공동경비구역(JSA)', { placeholder: 'body' });
s.addShape(pres.shapes.OVAL, { x: 8.9, y: 1.5, w: 3.6, h: 3.6, fill: { color: '1F3A66' }, line: { color: '1F3A66' }, objectName: nm('circ') });
icon(s, 'GiPineTree', '3F7F5F', 9.45, 1.9, 2.5, '미루나무');
icon(s, 'GiWoodAxe', 'FFFFFF', 10.9, 3.2, 1.3, '도끼');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ① 판문점 1976' });
s.addText('상황 보고 — 1976.8.18 오전', { placeholder: 'title' });
const tl1 = [
  ['10:30', '유엔사 경비대, JSA 내 시야를 가리는 미루나무 가지치기 작업 시작'],
  ['10:45', '북한군 장교가 작업 중지 요구 → 거부되자 북한군 30여 명 증원'],
  ['10:50', '도끼·곤봉으로 습격. 보니파스 대위·배럿 중위 사망, 한·미 장병 다수 부상'],
  ['8.19', '주한미군·한국군 데프콘 3 격상. 북한은 "미군이 먼저 도발" 주장'],
];
tl1.forEach(([t, d], i) => {
  step(i + 1, 'wipe');
  const y = 1.7 + i * 1.1;
  s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.12, w: 0.36, h: 0.36, fill: { color: i === 2 ? HEX.accent2 : HEX.accent1 }, line: { color: i === 2 ? HEX.accent2 : HEX.accent1 }, objectName: nm('dot') });
  T(s, t, { x: 1.4, y, w: 1.4, h: 0.6, fontSize: 20, bold: true, color: HEX.accent1, valign: 'middle' });
  T(s, d, { x: 2.9, y, w: 9.6, h: 0.6, fontSize: 18, color: HEX.lt1, valign: 'middle' });
});
T(s, '※ 당시 JSA에는 군사분계선 구분이 없어 양측 경비병이 자유롭게 오갔다.', { x: 0.8, y: 6.3, w: 11.5, h: 0.4, fontSize: 13, color: HEX.accent6 });
s.addNotes('시간은 공개 자료 기준 대략적 시각입니다. 학생들에게 당시 현장 지휘관이 된 것처럼 읽어 줍니다.');


s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ① 판문점 1976' });
s.addText('그날의 JSA — 개략도', { placeholder: 'title' });
card(s, 0.7, 1.6, 7.4, 5.0, '24304A', 'jsa');
T(s, '북측', { x: 0.95, y: 1.75, w: 2, h: 0.4, fontSize: 15, bold: true, color: HEX.accent6 });
T(s, '남측 (유엔사)', { x: 0.95, y: 6.05, w: 3, h: 0.4, fontSize: 15, bold: true, color: HEX.accent6 });
s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.4, w: 7.0, h: 0, line: { color: HEX.accent6, width: 1.5, dashType: 'dash' }, objectName: nm('mdl') });
T(s, '군사분계선 (당시 JSA 안에는 표시 없음)', { x: 0.95, y: 3.0, w: 3.8, h: 0.35, fontSize: 12, color: HEX.accent6, align: 'left' });
icon(s, 'GiPineTree', '3F7F5F', 5.4, 3.9, 1.6, '미루나무');
T(s, '미루나무', { x: 5.2, y: 5.5, w: 2.0, h: 0.4, fontSize: 15, bold: true, color: HEX.lt1, align: 'center' });
step(1);
icon(s, 'GiPerson', 'FFFFFF', 3.2, 4.3, 0.75); icon(s, 'GiPerson', 'FFFFFF', 3.9, 4.5, 0.75); icon(s, 'GiPerson', 'FFFFFF', 4.6, 4.3, 0.75);
T(s, '유엔사 가지치기 작업조', { x: 2.6, y: 5.3, w: 2.8, h: 0.4, fontSize: 14, color: HEX.lt1, align: 'center' });
T(s, [{ text: '① 10:30', options: { bold: true, color: HEX.accent1, breakLine: true } }, { text: '시야를 가리는 나무를 정리하러 들어간다', options: {} }], { x: 8.5, y: 1.8, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
step(2);
icon(s, 'GiPerson', 'B5482B', 4.4, 1.9, 0.7); icon(s, 'GiPerson', 'B5482B', 5.1, 1.75, 0.7); icon(s, 'GiPerson', 'B5482B', 5.8, 1.9, 0.7); icon(s, 'GiPerson', 'B5482B', 6.5, 1.75, 0.7);
arrow(s, 5.3, 2.7, 4.6, 4.2, 'B5482B'); arrow(s, 6.3, 2.7, 6.0, 3.9, 'B5482B');
T(s, [{ text: '② 10:45', options: { bold: true, color: HEX.accent1, breakLine: true } }, { text: '북한군 30여 명이 몰려와 작업 중지를 요구', options: {} }], { x: 8.5, y: 3.1, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
step(3);
icon(s, 'GiExplosionRays', 'B5482B', 3.5, 3.6, 1.2, '습격');
T(s, [{ text: '③ 10:50', options: { bold: true, color: HEX.accent2, breakLine: true } }, { text: '도끼·곤봉 습격 — 미군 장교 2명 사망', options: { bold: true } }], { x: 8.5, y: 4.4, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
T(s, '※ 위치 관계를 보여주기 위한 개략도이며 축척·실제 배치와 다르다.', { x: 0.7, y: 6.7, w: 11.0, h: 0.3, fontSize: 11, color: HEX.accent6 });
s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ① 판문점 1976' });
s.addText('당신이 유엔군사령관이라면? — 비례성의 저울에 올려 보자', { placeholder: 'title' });
const opts1 = [
  ['A', '동종·동량 보복', '같은 수준의 피해를 주도록 북한군 초소 제한 타격'],
  ['B', '외교적 항의', '군사정전위원회 항의, 사과 요구, 대비태세 유지'],
  ['C', '대규모 보복 타격', '도발 원점과 지원 부대까지 화력으로 응징'],
];
opts1.forEach(([k, h, d], i) => {
  step(i + 1);
  const x = 0.7 + i * 4.05;
  card(s, x, 1.8, 3.8, 3.4, '2A3140', 'opt' + k);
  T(s, k, { x: x + 0.3, y: 2.0, w: 1, h: 0.8, fontSize: 40, bold: true, color: HEX.accent1 });
  icon(s, ['GiCannon', 'GiMegaphone', 'GiExplosionRays'][i], 'E0A526', x + 2.6, 2.05, 0.9);
  T(s, h, { x: x + 0.3, y: 2.9, w: 3.2, h: 0.6, fontSize: 22, bold: true, color: HEX.lt1 });
  T(s, d, { x: x + 0.3, y: 3.6, w: 3.2, h: 1.3, fontSize: 16, color: HEX.accent6 });
});
step(4); T(s, '토의 2분 — 비례성으로 보면 가장 "정답"에 가까운 것은? 그런데 실제 선택은 A·B·C 어디에도 없었다.', { x: 0.7, y: 5.6, w: 11.9, h: 0.8, fontSize: 18, bold: true, color: HEX.accent1 });
s.addNotes('학생들이 A를 고르는 경우가 많습니다. 비례성 교과서대로라면 A가 자연스럽기 때문입니다. 실제 결과는 다음 장에서 공개합니다.');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ① 판문점 1976' });
s.addText('실제 선택: 폴 버니언 작전 (8.21) — "나무를 자른다"', { placeholder: 'title' });
const stats1 = [['B-52', '전략폭격기 전개'], ['항모', '미드웨이 항모전단'], ['64명', '한국군 특전사 결사대'], ['0발', '실제 사격']];
stats1.forEach(([n, l], i) => {
  step(i + 1);
  const x = 0.7 + i * 3.05;
  card(s, x, 1.7, 2.8, 2.2, i === 3 ? HEX.dk2 : HEX.lt2, 'stat' + i);
  icon(s, ['GiStealthBomber', 'GiShipBow', 'GiParachute', 'GiDove'][i], i === 3 ? 'E0A526' : '13294B', x + 1.05, 1.82, 0.7);
  T(s, n, { x, y: 2.5, w: 2.8, h: 0.8, fontSize: 38, bold: true, color: i === 3 ? HEX.accent1 : HEX.dk2, align: 'center' });
  T(s, l, { x, y: 3.3, w: 2.8, h: 0.45, fontSize: 15, color: i === 3 ? HEX.lt1 : HEX.dk1, align: 'center' });
});
step(5); T(s, [
  { text: '미군 공병과 한·미 경비병력이 문제의 미루나무를 절단. F-4·F-111 전투기까지 띄운 채 약 45분 만에 종료', options: { bullet: true, breakLine: true } },
  { text: '북한이 반격하면 개성 일대 포병을 칠 계획도 준비했지만, 북한군은 대응하지 않았다', options: { bullet: true, breakLine: true } },
  { text: '김일성이 유엔사에 유감 표명 → 위기 종결. 이후 JSA 경비구역 분리', options: { bullet: true } },
], { x: 0.8, y: 4.3, w: 11.8, h: 2.3, fontSize: 18, color: HEX.dk1, paraSpaceAfter: 8 });
s.addNotes('2명이 사망한 사건에 전략폭격기와 항모가 동원되었습니다. 비례성의 저울로만 보면 과잉입니다. 그런데 한 발도 쏘지 않았고, 위기는 끝났습니다.');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ① 판문점 1976' });
s.addText('왜 통했나 — 저울이 아니라 메시지였다', { placeholder: 'title' });
step(1); card(s, 0.7, 1.7, 5.8, 3.0, HEX.dk2, 'msg1');
icon(s, 'GiStealthBomber', 'E0A526', 5.5, 1.85, 0.8);
T(s, '능력 (Capability)', { x: 1.0, y: 1.9, w: 5.2, h: 0.6, fontSize: 22, bold: true, color: HEX.accent1 });
T(s, '압도적 전력 전개\n= "우리는 전쟁할 준비가 되어 있다"', { x: 1.0, y: 2.6, w: 5.2, h: 1.6, fontSize: 20, color: HEX.lt1 });
step(2); card(s, 6.85, 1.7, 5.8, 3.0, HEX.lt2, 'msg2');
icon(s, 'GiMegaphone', 'B5482B', 11.65, 1.85, 0.8);
T(s, '의사전달 (Communication)', { x: 7.15, y: 1.9, w: 5.2, h: 0.6, fontSize: 22, bold: true, color: HEX.accent2 });
T(s, '문제의 근원(나무)만 제거, 사격 없음\n= "그래도 전쟁은 원하지 않는다"', { x: 7.15, y: 2.6, w: 5.2, h: 1.6, fontSize: 20, color: HEX.dk1 });
step(3); T(s, '비례성으로 보면 "과잉 동원 + 과소 사용". 그러나 두 메시지를 동시에 보내 위기완화에 성공했다.', { x: 0.7, y: 5.1, w: 11.9, h: 0.6, fontSize: 18, bold: true, color: HEX.dk2 });
T(s, '함께 볼 점: 작전 중 일부 한국군이 북한군 초소를 파괴한 일은 통제와 확전 관리의 숙제를 남겼다.', { x: 0.7, y: 5.8, w: 11.9, h: 0.6, fontSize: 15, color: HEX.accent4 });

// ===== 장면 2
pres.addSection({ title: '장면 ② 연평도 2010' });
s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '장면 ② 연평도 2010' });
s.addText('장면 ②  13분', { placeholder: 'title' });
s.addText('2010년 11월 23일 오후, 서해 연평도', { placeholder: 'body' });
s.addShape(pres.shapes.OVAL, { x: 8.9, y: 1.5, w: 3.6, h: 3.6, fill: { color: '1F3A66' }, line: { color: '1F3A66' }, objectName: nm('circ') });
icon(s, 'GiExplosionRays', 'B5482B', 9.4, 1.9, 2.6, '포격');
icon(s, 'GiCannon', 'E0A526', 10.6, 3.3, 1.5, '해안포');

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ② 연평도 2010' });
s.addText('상황 보고 — 2010.11.23 오후', { placeholder: 'title' });
const tl2 = [
  ['14:34', '북한, 연평도 군 기지와 민가 지역에 해안포·방사포 포격 (6·25 이후 첫 민간 거주지 포격)'],
  ['14:47', '해병대 연평부대 K-9 자주포 대응사격 개시 — 첫 포격 13분 후'],
  ['15:10경', '북한 2차 포격, 아군 2차 대응사격으로 이어짐'],
  ['결과', '북한 170여 발 vs 아군 80여 발. 해병 2명·민간인 2명 사망, 다수 부상'],
];
tl2.forEach(([t, d], i) => {
  step(i + 1, 'wipe');
  const y = 1.7 + i * 1.1;
  s.addShape(pres.shapes.OVAL, { x: 0.8, y: y + 0.12, w: 0.36, h: 0.36, fill: { color: i === 3 ? HEX.accent2 : HEX.accent1 }, line: { color: i === 3 ? HEX.accent2 : HEX.accent1 }, objectName: nm('dot') });
  T(s, t, { x: 1.4, y, w: 1.4, h: 0.6, fontSize: 20, bold: true, color: HEX.accent1, valign: 'middle' });
  T(s, d, { x: 2.9, y, w: 9.6, h: 0.6, fontSize: 18, color: HEX.lt1, valign: 'middle' });
});
T(s, '※ 초계 중이던 전투기는 공대공 무장만 갖추고 있었다. 이후 공대지 무장 전투기를 출격시켰다.', { x: 0.8, y: 6.3, w: 11.5, h: 0.4, fontSize: 13, color: HEX.accent6 });
s.addNotes('시각과 발수는 국방부 발표 등 공개 자료 기준의 개략치입니다.');


s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ② 연평도 2010' });
s.addText('그날의 서해 — 개략도', { placeholder: 'title' });
card(s, 0.7, 1.6, 7.4, 5.0, '16324F', 'sea');
s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.6, w: 7.4, h: 1.1, fill: { color: '3A3F4A' }, line: { color: '3A3F4A' }, objectName: nm('coast') });
T(s, '북한 해안 — 해안포·방사포 진지', { x: 0.95, y: 1.75, w: 5, h: 0.4, fontSize: 15, bold: true, color: HEX.lt1 });
icon(s, 'GiCannon', 'B5482B', 1.6, 2.05, 0.6); icon(s, 'GiCannon', 'B5482B', 5.6, 2.05, 0.6);
s.addShape(pres.shapes.LINE, { x: 0.8, y: 3.35, w: 7.2, h: 0, line: { color: HEX.accent1, width: 2, dashType: 'dash' }, objectName: nm('nll') });
T(s, 'NLL', { x: 7.2, y: 2.95, w: 0.8, h: 0.35, fontSize: 14, bold: true, color: HEX.accent1, align: 'right' });
s.addShape(pres.shapes.OVAL, { x: 3.1, y: 4.7, w: 2.0, h: 1.0, fill: { color: '3F7F5F' }, line: { color: '3F7F5F' }, objectName: nm('island') });
T(s, '연평도', { x: 3.1, y: 4.7, w: 2.0, h: 1.0, fontSize: 16, bold: true, color: HEX.lt1, align: 'center', valign: 'middle' });
T(s, '약 12km', { x: 0.9, y: 3.9, w: 1.6, h: 0.35, fontSize: 13, color: HEX.accent6 });
step(1);
arrow(s, 1.9, 2.7, 3.5, 4.7, 'B5482B'); arrow(s, 5.9, 2.7, 4.7, 4.7, 'B5482B');
icon(s, 'GiExplosionRays', 'B5482B', 3.55, 4.25, 0.9, '포탄 낙탄');
T(s, [{ text: '① 14:34  북한 포격', options: { bold: true, color: HEX.accent2, breakLine: true } }, { text: '군 기지와 민가에 170여 발', options: {} }], { x: 8.5, y: 1.8, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
step(2);
arrow(s, 3.8, 4.7, 2.3, 2.7, 'E0A526'); arrow(s, 4.4, 4.7, 5.5, 2.7, 'E0A526');
T(s, [{ text: '② 14:47  K-9 대응사격', options: { bold: true, color: HEX.accent1, breakLine: true } }, { text: '13분 뒤, 80여 발 — 같은 무기로', options: {} }], { x: 8.5, y: 3.1, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
step(3);
icon(s, 'GiJetFighter', 'E0A526', 6.2, 4.6, 1.1, '전투기');
T(s, '공대공 무장', { x: 5.9, y: 5.75, w: 1.8, h: 0.4, fontSize: 13, color: HEX.accent1, align: 'center' });
T(s, [{ text: '③ 하늘의 전투기', options: { bold: true, color: HEX.accent1, breakLine: true } }, { text: '공대공 무장뿐 — 원점을 칠 수 없었다', options: {} }], { x: 8.5, y: 4.4, w: 4.2, h: 1.1, fontSize: 17, color: HEX.lt1 });
T(s, '※ 위치 관계를 보여주기 위한 개략도이며 축척·실제 배치와 다르다.', { x: 0.7, y: 6.7, w: 11.0, h: 0.3, fontSize: 11, color: HEX.accent6 });
s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '장면 ② 연평도 2010' });
s.addText('당신이 합참 작전본부장이라면? — 전투기로 칠 것인가', { placeholder: 'title' });
step(1); card(s, 0.7, 1.8, 5.8, 3.6, '2A3140', 'rule');
icon(s, 'GiScales', 'E0A526', 5.5, 1.95, 0.8);
T(s, '손에 쥔 규칙', { x: 1.0, y: 2.0, w: 5.2, h: 0.6, fontSize: 22, bold: true, color: HEX.accent1 });
T(s, [
  { text: '유엔사 정전시 교전규칙', options: { bullet: true, breakLine: true } },
  { text: '비례성 = 동종·동량 대응', options: { bullet: true, breakLine: true } },
  { text: '포격에는 포격으로, 같은 양으로', options: { bullet: true } },
], { x: 1.0, y: 2.7, w: 5.2, h: 2.4, fontSize: 18, color: HEX.lt1, paraSpaceAfter: 8 });
step(2); card(s, 6.85, 1.8, 5.8, 3.6, '2A3140', 'reality');
icon(s, 'GiHouse', 'B5482B', 11.65, 1.95, 0.8);
T(s, '눈앞의 현실', { x: 7.15, y: 2.0, w: 5.2, h: 0.6, fontSize: 22, bold: true, color: HEX.accent2 });
T(s, [
  { text: '민가에 포탄이 떨어지고 있다', options: { bullet: true, breakLine: true } },
  { text: '북한 포는 해안 갱도 속에 있다', options: { bullet: true, breakLine: true } },
  { text: 'K-9만으로 원점을 제압할 수 있나?', options: { bullet: true } },
], { x: 7.15, y: 2.7, w: 5.2, h: 2.4, fontSize: 18, color: HEX.lt1, paraSpaceAfter: 8 });
step(3); T(s, '토의 2분 — "같은 무기, 같은 양"으로 대응하면 비례적이다. 그런데 그것으로 충분한가?', { x: 0.7, y: 5.8, w: 11.9, h: 0.7, fontSize: 18, bold: true, color: HEX.accent1 });
s.addNotes('합참이 공중 대응에 대해 유엔군사령관의 승인을 구하자 "이것은 한국군의 자위권 문제"라는 답을 들었다는 이야기가 전해집니다. 공식 기록이 아니므로 "전해지는 이야기"로만 소개합니다.');

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ② 연평도 2010' });
s.addText('결과: 비례했지만, 충분했는가?', { placeholder: 'title' });
const stats2 = [['13분', '대응까지 걸린 시간'], ['170 : 80', '북한 포탄 대 아군 대응 (발)'], ['4명', '사망 (해병 2·민간인 2)']];
stats2.forEach(([n, l], i) => {
  step(i + 1);
  const x = 0.7 + i * 4.05;
  card(s, x, 1.7, 3.8, 2.2, i === 1 ? HEX.dk2 : HEX.lt2, 'st2' + i);
  icon(s, ['GiStopwatch', 'GiCannon', 'GiCandleFlame'][i], i === 1 ? 'E0A526' : '13294B', x + 1.55, 1.82, 0.7);
  T(s, n, { x, y: 2.5, w: 3.8, h: 0.8, fontSize: 38, bold: true, color: i === 1 ? HEX.accent1 : HEX.dk2, align: 'center' });
  T(s, l, { x, y: 3.3, w: 3.8, h: 0.45, fontSize: 15, color: i === 1 ? HEX.lt1 : HEX.dk1, align: 'center' });
});
step(4); T(s, [
  { text: '규칙이 정한 비례성(동종·동량)은 지켰지만, 도발을 멈추게 하고 다시 못 하게 하는 데 충분했는지는 논란으로 남았다', options: { bullet: true, breakLine: true } },
  { text: '국민 여론은 "왜 더 강하게 대응하지 않았나"를 물었다', options: { bullet: true } },
], { x: 0.8, y: 4.3, w: 11.8, h: 1.6, fontSize: 18, color: HEX.dk1, paraSpaceAfter: 8 });
step(5); T(s, '판문점은 비례성을 넘어 이겼고, 연평도는 비례성에 묶여 아쉬움을 남겼다.', { x: 0.8, y: 6.0, w: 11.8, h: 0.6, fontSize: 18, bold: true, color: HEX.accent2 });

s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '장면 ② 연평도 2010' });
s.addText('그 뒤 우리는 비례성을 다시 정의했다', { placeholder: 'title' });
const steps = [
  ['자위권', '국지도발 대응을 교전규칙이 아닌 자위권 차원의 개념으로 정립'],
  ['신속·정확·충분', '비례성을 "같은 양"이 아니라 "도발을 멈추게 하는 데 충분한가"로 해석'],
  ['선조치 후보고', '현장 지휘관의 자위권 대응을 강조, 민간인 피해 시 대응도 규칙에 반영'],
  ['한·미 공동계획', '확전을 우려한 미국과 공동 국지도발 대비계획 마련: 한국군 주도·미군 지원'],
];
steps.forEach(([h, d], i) => {
  step(i + 1, 'wipe');
  const y = 1.65 + i * 1.2;
  s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.8, h: 0.8, fill: { color: HEX.dk2 }, line: { color: HEX.dk2 }, objectName: nm('n') });
  T(s, String(i + 1), { x: 0.8, y, w: 0.8, h: 0.8, fontSize: 24, bold: true, color: HEX.accent1, align: 'center', valign: 'middle' });
  T(s, h, { x: 1.9, y, w: 2.9, h: 0.8, fontSize: 20, bold: true, color: HEX.dk2, valign: 'middle' });
  T(s, d, { x: 4.9, y, w: 7.7, h: 0.8, fontSize: 17, color: HEX.dk1, valign: 'middle' });
});
T(s, '※ 2021년 서해수호의 날에 "연평도 포격전"이 공식 명칭으로 처음 쓰였다.', { x: 0.8, y: 6.55, w: 11.8, h: 0.4, fontSize: 13, color: HEX.accent4 });

// ===== 종합
pres.addSection({ title: '종합' });
s = (step(0), pres).addSlide({ masterName: 'LIGHT_CONTENT', sectionTitle: '종합' });
s.addText('두 장면을 나란히 놓으면', { placeholder: 'title' });
step(1); s.addTable([
  [{ text: '', options: { fill: { color: HEX.dk2 } } },
   { text: '판문점 1976', options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } },
   { text: '연평도 2010', options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } }],
  [{ text: '비례성의 저울', options: { bold: true, fill: { color: HEX.lt2 } } }, '과잉 동원, 과소 사용 — 저울을 벗어남', '동종·동량 — 저울을 지킴'],
  [{ text: '결과', options: { bold: true, fill: { color: HEX.lt2 } } }, '북한 유감 표명, 위기완화', '충분성 논란, 억제 효과 의문'],
  [{ text: '무엇이 갈랐나', options: { bold: true, fill: { color: HEX.lt2 } } }, '능력과 의사를 동시에 보여준 메시지', '규칙이 대응의 수준을 묶음'],
  [{ text: '남긴 것', options: { bold: true, fill: { color: HEX.lt2 } } }, '억제의 성공조건 = 능력 + 의사전달', '목적 중심 비례성, 자위권 대응 개념'],
], { x: 0.7, y: 1.7, w: 11.9, colW: [2.6, 4.65, 4.65], fontFace: FONT, fontSize: 16, color: HEX.dk1, border: { type: 'solid', pt: 0.75, color: 'BFC8D6' }, rowH: 0.75, valign: 'middle', margin: 0.1, objectName: nm('tbl') });
step(2); T(s, '비례성은 판단의 출발점이지, 위기를 끝내는 정답 그 자체는 아니다.', { x: 0.7, y: 6.0, w: 11.9, h: 0.6, fontSize: 20, bold: true, color: HEX.accent2 });

s = (step(0), pres).addSlide({ masterName: 'DARK_SCENE', sectionTitle: '종합' });
s.addText('여러분의 판단서에 다시 묻는다', { placeholder: 'title' });
const qs = [
  ['Q1', '우리 방책에서 "우리는 준비되어 있다"와 "전쟁은 원하지 않는다"는 각각 어디에 있는가?'],
  ['Q2', '우리가 쓴 비례성은 2010년의 기준(같은 양)인가, 그 이후의 기준(충분한가)인가?'],
  ['Q3', '비례성 표가 "충족"인데도 위기가 끝나지 않는다면, 무엇이 빠진 것인가?'],
];
qs.forEach(([k, q], i) => {
  step(i + 1);
  const y = 1.8 + i * 1.45;
  card(s, 0.7, y, 11.9, 1.2, '2A3140', 'q' + i);
  T(s, k, { x: 1.0, y, w: 1.0, h: 1.2, fontSize: 28, bold: true, color: HEX.accent1, valign: 'middle' });
  T(s, q, { x: 2.1, y, w: 10.2, h: 1.2, fontSize: 19, color: HEX.lt1, valign: 'middle' });
});
s.addNotes('각 조의 판단서를 펴 놓고 1분씩 답하게 합니다. 정답을 고르는 질문이 아니라, 자기 방책의 빈칸을 찾는 질문입니다.');

s = (step(0), pres).addSlide({ masterName: 'DARK_TITLE', sectionTitle: '종합' });
s.addText('위기를 끝내는 것은 저울이 아니라\n상대의 계산을 바꾸는 것이다', { placeholder: 'title' });
icon(s, 'GiChessKnight', 'E0A526', 10.4, 0.6, 2.0, '체스 나이트');
s.addText('비례성은 지키되, 그 안에 능력과 의사를 함께 담아라 — 그것이 합참의 위기관리다', { placeholder: 'body' });

(async () => {
  const out = process.argv[2];
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('ok', out);
})();
