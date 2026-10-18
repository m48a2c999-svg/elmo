// 졸업연구 웹 포트폴리오 — 구축·게시 강의안 (Framer)
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const THEME = {
  name: "Space Directing Green",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "1A1F1C", lt1: "FFFFFF", dk2: "1F3A2E", lt2: "EEF3EF",
    accent1: "2F6B4F", accent2: "8DB89A", accent3: "D9A441",
    accent4: "5E6B64", accent5: "C9D9CD", accent6: "B5543C",
    hlink: "2F6B4F", folHlink: "5E6B64",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "웹 포트폴리오 구축과 게시";
pres.author = "호서대학교 실내건축학과 졸업연구";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6;

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.1, w: 11.5, h: 1.9, fontSize: 46, bold: true, color: C.background1, align: "left", valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.35, w: 11.5, h: 1.0, fontSize: 20, color: C.accent5, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.45, w: 10.6, h: 0.8, fontSize: 32, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "졸업연구 · Web Portfolio", options: { x: M, y: 7.0, w: 5, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent4, align: "right" },
});

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 15, color: C.text1, valign: "top" }, o));
const box = (s, x, y, w, h, fill, name) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill }, objectName: name });
const kicker = (s, text) => T(s, text, { x: M, y: 0.2, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
const content = (sec, title, k) => {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText(title, { placeholder: "title" });
  if (k) kicker(s, k);
  return s;
};
const num = (s, n, x, y, d = 0.55, fill = C.accent1) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: "num" + n });
  T(s, String(n).padStart(2, "0"), { x, y, w: d, h: d, align: "center", valign: "middle", bold: true, fontSize: 14, color: C.background1 });
};

// 1 표지
pres.addSection({ title: "도입" });
let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "도입" });
T(s, "WEB PORTFOLIO · 구축과 게시", { x: M, y: 1.6, w: 8, h: 0.4, fontSize: 14, bold: true, color: C.accent3, charSpacing: 2 });
s.addText("내 졸업작품 사이트를\n만들고, 올린다", { placeholder: "title" });
s.addText("Framer 템플릿 선택 → 6개 섹션 구성 → 모바일 점검 → 게시 · QR", { placeholder: "body" });
T(s, "호서대학교 실내건축학과 · 2026학년도 2학기 졸업연구 · 검토용", { x: M, y: 6.6, w: 10, h: 0.3, fontSize: 11, color: C.accent2 });
s.addNotes("지난 차시에 Framer 가입까지 마쳤습니다. 오늘부터 실제로 사이트를 만들고, 마지막에는 누구나 접속할 수 있는 주소로 올립니다. 끝나면 각자 공개 URL 하나와 판넬용 QR 코드 하나를 갖게 됩니다.");

// 2 오늘의 흐름
s = content("도입", "오늘의 흐름", "ROADMAP");
const flow = [
  ["준비", "브랜드 시트 · 이미지 · 영상 링크 모으기"],
  ["템플릿", "작품에 맞는 템플릿을 골라 내 프로젝트로 복제"],
  ["구성", "Hero부터 About까지 6개 섹션 채우기"],
  ["점검", "Desktop · Tablet · Phone 화면 확인"],
  ["게시", "Publish로 주소 발급, QR 코드 만들기"],
];
flow.forEach(([h, d], i) => {
  const x = M + i * 2.45, y = 2.0;
  box(s, x, y, 2.2, 3.6, i === 4 ? C.accent1 : C.background2, "flow" + i);
  num(s, i + 1, x + 0.25, y + 0.3, 0.55, i === 4 ? C.text2 : C.accent1);
  T(s, h, { x: x + 0.25, y: y + 1.1, w: 1.8, h: 0.5, fontSize: 22, bold: true, color: i === 4 ? C.background1 : C.text2 });
  T(s, d, { x: x + 0.25, y: y + 1.75, w: 1.75, h: 1.6, fontSize: 14, color: i === 4 ? C.background1 : C.text1 });
});
T(s, "1~3은 이번 차시, 4~5는 다음 차시에 마무리합니다. 집에서 끝내도 됩니다.", { x: M, y: 6.0, w: 12, h: 0.4, fontSize: 14, color: C.accent4 });
s.addNotes("다섯 단계입니다. 준비, 템플릿, 구성은 오늘 수업 안에서 시작하고, 점검과 게시는 다음 시간에 같이 확인합니다. 중요한 건 순서입니다. 자료를 모으기 전에 템플릿부터 고치기 시작하면 두 번 일하게 됩니다.");

// 3 완성 기준
s = content("도입", "완성 기준: 이 6개 섹션과 5개 조건", "DONE = ");
const secs = ["Hero", "Video", "Concept", "Drawing", "Render", "About"];
secs.forEach((n, i) => {
  const x = M + (i % 3) * 2.1, y = 1.7 + Math.floor(i / 3) * 1.55;
  box(s, x, y, 1.9, 1.3, C.background2, "sec" + i);
  T(s, String(i + 1).padStart(2, "0"), { x: x + 0.2, y: y + 0.18, w: 1, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, n, { x: x + 0.2, y: y + 0.55, w: 1.6, h: 0.5, fontSize: 22, bold: true, color: C.text2 });
});
T(s, "필수 섹션", { x: M, y: 4.95, w: 6, h: 0.35, fontSize: 13, color: C.accent4 });
box(s, 7.2, 1.7, 5.5, 4.3, C.text2, "rules");
T(s, [
  { text: "실제 접속되는 공개 URL", options: { bullet: true, breakLine: true } },
  { text: "모바일에서 읽히는 반응형", options: { bullet: true, breakLine: true } },
  { text: "영상 자동재생 금지", options: { bullet: true, breakLine: true } },
  { text: "첫 화면 로딩 5초 이내", options: { bullet: true, breakLine: true } },
  { text: "판넬에 붙일 QR 코드", options: { bullet: true } },
], { x: 7.55, y: 2.05, w: 4.9, h: 3.6, fontSize: 18, color: C.background1, paraSpaceAfter: 14 });
T(s, "출처: 2026-2학기 수업 안내 · 산출물 ③", { x: M, y: 6.4, w: 8, h: 0.3, fontSize: 11, color: C.accent4 });
s.addNotes("평가할 때 보는 기준을 먼저 보여줍니다. 섹션 여섯 개는 학기 초 수업 안내에 있던 그대로입니다. 오른쪽 다섯 조건 중 하나라도 빠지면 미완성으로 봅니다. 특히 영상 자동재생은 끄세요. 모바일 데이터로 열었을 때 바로 소리가 나는 사이트는 감점입니다.");

// 4 준비물
pres.addSection({ title: "준비" });
s = content("준비", "시작 전에 모을 것", "STEP 01 · 준비");
const prep = [
  ["브랜드 시트", "컬러 3색 HEX · 폰트 2종 · 포지셔닝 한 문장"],
  ["대표 이미지", "Hero에 쓸 가로 렌더 1장. 가장 자신 있는 컷"],
  ["렌더 6~12장", "공간 순서대로 번호를 붙여 한 폴더에"],
  ["도면 2~4장", "평면 · 단면 · 다이어그램을 이미지로 내보내기"],
  ["영상 링크", "5분 영상을 YouTube에 일부 공개로 업로드"],
  ["About 문장", "자기소개 2~3줄 + 연락처(이메일 · 인스타그램)"],
];
prep.forEach(([h, d], i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const x = M + col * 6.15, y = 1.65 + row * 1.6;
  box(s, x, y, 5.9, 1.35, C.background2, "prep" + i);
  num(s, i + 1, x + 0.3, y + 0.4);
  T(s, h, { x: x + 1.1, y: y + 0.22, w: 4.6, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  T(s, d, { x: x + 1.1, y: y + 0.68, w: 4.6, h: 0.5, fontSize: 14 });
});
s.addNotes("템플릿을 열기 전에 이 여섯 가지를 한 폴더에 모읍니다. 브랜드 시트의 HEX 코드와 폰트 이름은 그대로 Framer에 입력합니다. 영상은 Framer에 직접 올리지 않고 YouTube 링크로 연결합니다. 공개 범위는 '일부 공개'로 두면 링크를 아는 사람만 볼 수 있습니다.");

// 5 이미지 규격
s = content("준비", "이미지는 올리기 전에 줄인다", "STEP 01 · 파일 규격");
const spec = [["가로 2400px", "이하로 리사이즈"], ["1MB 안팎", "한 장 기준 권장 용량"], ["JPG · WebP", "PNG는 도면 · 로고만"]];
spec.forEach(([b, l], i) => {
  const x = M + i * 4.1;
  T(s, b, { x, y: 1.9, w: 3.9, h: 0.9, fontSize: 40, bold: true, color: C.accent1 });
  T(s, l, { x, y: 2.85, w: 3.9, h: 0.4, fontSize: 15, color: C.accent4 });
});
box(s, M, 3.8, 12.1, 2.6, C.background2, "nameRule");
T(s, "파일 이름 규칙", { x: M + 0.35, y: 4.05, w: 5, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
T(s, "01_hero.jpg   02_lobby.jpg   03_corridor.jpg …", { x: M + 0.35, y: 4.6, w: 11.4, h: 0.5, fontSize: 20, color: C.text2, fontFace: "Courier New" });
T(s, "영문 소문자 · 숫자 · 밑줄만 사용합니다. 한글과 공백이 들어간 이름은 링크에서 깨질 수 있습니다. 클로드 코드에 \"이 폴더 이미지를 가로 2400px, WebP로 바꾸고 번호순으로 이름 바꿔줘\"라고 요청하면 한 번에 처리됩니다.", { x: M + 0.35, y: 5.2, w: 11.4, h: 1.0, fontSize: 14 });
s.addNotes("트윈모션 렌더는 보통 4K라 한 장에 5~10MB입니다. 그대로 올리면 첫 화면이 늦게 뜨고 로딩 5초 조건을 못 지킵니다. 가로 2400px 정도면 큰 모니터에서도 충분합니다. 5주차에 클로드 코드로 에셋 정리를 했던 학생은 같은 방법을 쓰면 됩니다.");

// 6 템플릿 고르기
pres.addSection({ title: "템플릿" });
s = content("템플릿", "템플릿은 세 가지만 보고 고른다", "STEP 02 · 템플릿");
const crit = [
  ["사진이 크게 들어가는가", "첫 화면에 렌더 한 장이 화면을 채우는 구성"],
  ["메뉴가 단순한가", "메뉴 4~5개 이하. 6개 섹션과 그대로 맞출 수 있는지"],
  ["무료인가", "Free 표시 확인. 유료 템플릿은 구매가 필요"],
];
crit.forEach(([h, d], i) => {
  const y = 1.7 + i * 1.25;
  num(s, i + 1, M, y + 0.1);
  T(s, h, { x: M + 0.8, y, w: 5.6, h: 0.45, fontSize: 19, bold: true, color: C.text2 });
  T(s, d, { x: M + 0.8, y: y + 0.5, w: 5.6, h: 0.5, fontSize: 14 });
});
box(s, 7.3, 1.6, 5.4, 4.6, C.background2, "ref");
T(s, "지난 차시에 본 무료 템플릿", { x: 7.65, y: 1.85, w: 4.8, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
[["Forma Interior", "공간 사진을 크게 보여주는 첫 화면"], ["VOLUM", "건축 사진 + 절제된 메뉴"], ["ARCHIOR", "여러 프로젝트를 갤러리로 정리"]].forEach(([n, d], i) => {
  T(s, n, { x: 7.65, y: 2.5 + i * 1.15, w: 4.8, h: 0.45, fontSize: 20, bold: true, color: C.text2 });
  T(s, d, { x: 7.65, y: 2.95 + i * 1.15, w: 4.8, h: 0.4, fontSize: 14, color: C.accent4 });
});
T(s, "템플릿 페이지에서 Remix(복제)를 누르면 내 프로젝트로 복사됩니다. 원본은 바뀌지 않습니다.", { x: M, y: 5.6, w: 6.4, h: 0.7, fontSize: 13, color: C.accent4 });
s.addNotes("템플릿은 오래 고르지 않습니다. 사진이 크게, 메뉴는 단순하게, 그리고 무료인지. 이 세 가지만 확인하세요. 지난 시간에 본 세 개 중 하나를 써도 됩니다. Remix를 누르면 내 계정에 복사본이 생기고 그걸 고칩니다. 버튼 이름은 화면 업데이트로 바뀔 수 있으니 수업 당일 화면을 기준으로 안내합니다.");

// 7 정보구조
s = content("템플릿", "한 페이지를 위에서 아래로 걷게 만든다", "STEP 03 · 구조");
const ia = [["Hero", "대표 렌더 + 작품명 + 한 줄 컨셉"], ["Video", "5분 영상 (YouTube)"], ["Concept", "컨셉 문장 + 다이어그램"], ["Drawing", "평면 · 단면"], ["Render", "공간 순서대로 렌더"], ["About", "나 · 연락처 · 사용 도구"]];
ia.forEach(([n, d], i) => {
  const y = 1.55 + i * 0.86;
  box(s, M, y, 2.4, 0.68, i === 0 ? C.accent1 : C.background2, "ia" + i);
  T(s, n, { x: M + 0.25, y, w: 2.0, h: 0.68, valign: "middle", fontSize: 17, bold: true, color: i === 0 ? C.background1 : C.text2 });
  T(s, d, { x: M + 2.7, y, w: 4.2, h: 0.68, valign: "middle", fontSize: 14 });
});
box(s, 7.6, 1.55, 5.1, 4.9, C.text2, "why");
T(s, "왜 한 페이지인가", { x: 7.95, y: 1.85, w: 4.5, h: 0.45, fontSize: 18, bold: true, color: C.accent3 });
T(s, "판넬 QR로 들어온 사람은 휴대폰으로 스크롤만 합니다. 메뉴를 눌러 다른 페이지로 가는 일은 거의 없습니다.\n\n6주차에 다룬 '스크롤은 곧 걷기'를 그대로 적용합니다. 영상과 컨셉으로 들어가 공간을 둘러보고, About에서 나옵니다.\n\n메뉴는 같은 페이지의 섹션으로 이동하는 링크로 둡니다.", { x: 7.95, y: 2.45, w: 4.5, h: 3.8, fontSize: 14, color: C.background1 });
s.addNotes("템플릿에 About, Projects, Contact처럼 페이지가 여러 개 있어도, 홈 한 페이지에 여섯 섹션을 순서대로 쌓는 것을 기본으로 합니다. 메뉴 링크는 페이지 대신 섹션으로 연결합니다. 프로젝트가 두 개 이상인 학생만 상세 페이지를 추가합니다.");

// 8 브랜드 적용
pres.addSection({ title: "구성" });
s = content("구성", "색과 글꼴은 스타일로 한 번에 바꾼다", "STEP 03 · 브랜드 적용");
const steps8 = [
  ["Assets 패널 열기", "왼쪽 패널에서 Colors · Text 스타일 목록을 확인"],
  ["Color Style 교체", "템플릿의 주조 · 보조 · 강조 색을 브랜드 시트 HEX로 수정"],
  ["Text Style 교체", "제목 · 본문 스타일의 글꼴을 브랜드 폰트 2종으로"],
  ["전체 확인", "스타일을 쓰는 모든 요소가 함께 바뀌었는지 스크롤로 확인"],
];
steps8.forEach(([h, d], i) => {
  const y = 1.65 + i * 1.15;
  num(s, i + 1, M, y + 0.05);
  T(s, h, { x: M + 0.8, y, w: 6.0, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
  T(s, d, { x: M + 0.8, y: y + 0.48, w: 6.0, h: 0.5, fontSize: 14 });
});
// 60-30-10 swatch
T(s, "브랜드 컬러 비율", { x: 7.6, y: 1.65, w: 5, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
s.addShape(pres.shapes.RECTANGLE, { x: 7.6, y: 2.2, w: 3.06, h: 2.2, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: "c60" });
s.addShape(pres.shapes.RECTANGLE, { x: 10.66, y: 2.2, w: 1.53, h: 2.2, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "c30" });
s.addShape(pres.shapes.RECTANGLE, { x: 12.19, y: 2.2, w: 0.51, h: 2.2, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: "c10" });
T(s, "주조 60", { x: 7.75, y: 3.9, w: 2, h: 0.35, fontSize: 13, color: C.background1, bold: true });
T(s, "보조 30", { x: 10.75, y: 3.9, w: 1.4, h: 0.35, fontSize: 13, color: C.text2, bold: true });
T(s, "강조 10", { x: 11.0, y: 4.5, w: 1.7, h: 0.35, fontSize: 12, color: C.accent4, align: "right" });
T(s, "강조색은 버튼 · 링크 · 숫자에만. 한글 웹폰트가 목록에 없으면 Pretendard · Noto Sans KR처럼 Google Fonts에 있는 글꼴로 대신합니다.", { x: 7.6, y: 5.0, w: 5.1, h: 1.2, fontSize: 13, color: C.accent4 });
s.addNotes("요소 하나하나 색을 바꾸지 마세요. 템플릿은 색과 글꼴을 스타일로 묶어 두었기 때문에 스타일만 바꾸면 사이트 전체가 바뀝니다. 브랜드 시트의 HEX 코드를 그대로 입력합니다. 영상 자막, 숏폼과 같은 색이 되는지가 14주차 브랜드 점검 항목입니다.");

// 9 섹션별 내용 표
s = content("구성", "섹션마다 넣을 것과 분량", "STEP 03 · 내용");
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
const rows = [
  [hdr("섹션"), hdr("넣을 것"), hdr("분량"), hdr("자주 하는 실수")],
  ["Hero", "대표 렌더, 작품명, 한 줄 컨셉, 이름", "이미지 1 · 글 2줄", "글이 사진을 가림"],
  ["Video", "YouTube 영상 임베드, 러닝타임", "영상 1", "자동재생 켜 둠"],
  ["Concept", "컨셉 문장, 컨셉 · 동선 다이어그램", "글 3~4줄 · 이미지 1~2", "설명이 너무 김"],
  ["Drawing", "평면도, 단면도", "이미지 2~4", "선이 가늘어 모바일에서 안 보임"],
  ["Render", "공간 순서대로 렌더, 짧은 캡션", "이미지 6~12", "같은 앵글 반복"],
  ["About", "자기소개, 연락처, 사용 도구 표기", "글 3줄 · 링크 2~3", "AI 도구 표기 누락"],
];
s.addTable(rows, {
  x: M, y: 1.55, w: 12.1, colW: [1.7, 4.6, 2.5, 3.3], rowH: 0.62,
  fontSize: 14, color: C.text1, valign: "middle", margin: [0, 0.12, 0, 0.12],
  border: { type: "solid", pt: 0.75, color: "C9D9CD" }, fill: { color: C.background1 },
});
T(s, "쓰지 않는 템플릿 섹션(Testimonials, Pricing 등)은 숨기지 말고 삭제합니다.", { x: M, y: 6.2, w: 12, h: 0.4, fontSize: 13, color: C.accent4 });
s.addNotes("분량을 정해 두면 끝이 보입니다. 렌더는 많을수록 좋은 게 아니라 공간 순서대로 6~12장이면 충분합니다. 도면은 모바일에서 선이 안 보이는 경우가 많으니 선 두께를 올려서 내보내세요. About에 사용한 AI 도구를 적는 것은 수업 안내의 저작권 항목입니다.");

// 10 교체 방법
s = content("구성", "이미지와 글은 이렇게 바꾼다", "STEP 03 · 편집");
const edits = [
  ["이미지", ["이미지 요소 선택", "오른쪽 속성 패널의 Fill에서 이미지 교체", "필요하면 Fit · Fill로 잘림 조정"]],
  ["글", ["글 상자 더블클릭", "내용 입력 후 바깥 클릭", "스타일은 Text Style로 유지"]],
  ["섹션", ["왼쪽 Layers에서 섹션 선택", "복제: Ctrl/⌘ + D", "삭제: Delete"]],
];
edits.forEach(([h, items], i) => {
  const x = M + i * 4.1;
  box(s, x, 1.65, 3.85, 4.5, C.background2, "edit" + i);
  T(s, h, { x: x + 0.3, y: 1.9, w: 3.2, h: 0.55, fontSize: 24, bold: true, color: C.accent1 });
  T(s, items.map((t, j) => ({ text: t, options: { bullet: { type: "number" }, breakLine: j < items.length - 1 } })), { x: x + 0.3, y: 2.65, w: 3.3, h: 3.2, fontSize: 15, paraSpaceAfter: 12 });
});
T(s, "되돌리기는 Ctrl/⌘ + Z. 작업 내용은 자동 저장되지만, 방문자에게는 Publish를 누르기 전까지 보이지 않습니다.", { x: M, y: 6.35, w: 12.1, h: 0.4, fontSize: 13, color: C.accent4 });
s.addNotes("가장 많이 쓰는 동작 세 가지입니다. 메뉴 위치와 이름은 Framer 업데이트로 조금씩 바뀌므로 수업 당일 화면을 같이 보면서 진행합니다. 자동 저장된다는 것과 게시된다는 것은 다릅니다. 이 차이는 게시 단계에서 다시 이야기합니다.");

// 11 영상
s = content("구성", "영상은 YouTube로 올리고 연결만 한다", "STEP 03 · 영상");
const v = [["YouTube 업로드", "공개 범위 '일부 공개'. 퍼가기 허용 확인"], ["Framer에 삽입", "Insert의 Media에서 YouTube 선택"], ["주소 붙여넣기", "영상 URL 입력, 가로 16:9 비율 유지"], ["자동재생 끄기", "Autoplay 해제. 소리는 방문자가 켜게"]];
v.forEach(([h, d], i) => {
  const x = M + i * 3.05;
  num(s, i + 1, x, 1.8, 0.6, i === 3 ? C.accent6 : C.accent1);
  T(s, h, { x, y: 2.6, w: 2.8, h: 0.45, fontSize: 18, bold: true, color: i === 3 ? C.accent6 : C.text2 });
  T(s, d, { x, y: 3.1, w: 2.75, h: 1.0, fontSize: 14 });
});
box(s, M, 4.5, 12.1, 1.7, C.background2, "vnote");
T(s, "왜 YouTube인가", { x: M + 0.35, y: 4.7, w: 4, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
T(s, "5분 영상 파일은 수백 MB라 사이트에 직접 올리면 로딩 5초 조건을 지킬 수 없습니다. YouTube는 화질을 자동으로 맞춰 줍니다. 무료 영상 앞에 광고가 붙을 수 있다는 점은 지난 차시에 안내했습니다.", { x: M + 0.35, y: 5.15, w: 11.4, h: 0.9, fontSize: 14 });
s.addNotes("영상은 사이트에 파일로 올리지 않습니다. YouTube에 일부 공개로 올리고 링크만 붙입니다. 반드시 자동재생을 끄세요. 수업 안내의 필수 조건입니다. 퍼가기가 꺼져 있으면 사이트에서 '재생할 수 없음'이 뜹니다.");

// 12 모바일 점검
pres.addSection({ title: "점검·게시" });
s = content("점검·게시", "PC에서 끝난 게 아니다. 휴대폰이 기준이다", "STEP 04 · 반응형");
const bps = [["Desktop", 2.8, 1.8], ["Tablet", 2.0, 2.6], ["Phone", 1.1, 2.9]];
let bx = M;
bps.forEach(([n, w, h], i) => {
  const y = 4.6 - h;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y, w, h, rectRadius: 0.1, fill: { color: i === 2 ? C.accent1 : C.background2 }, line: { color: C.accent1, width: 1.5 }, objectName: "bp" + i });
  T(s, n, { x: bx, y: 4.75, w: Math.max(w, 1.3), h: 0.4, fontSize: 15, bold: true, color: C.text2 });
  bx += w + 0.4;
});
T(s, "화면 상단의 Breakpoint 탭을 바꿔 가며 크기별로 고칩니다.", { x: M, y: 5.35, w: 6.6, h: 0.7, fontSize: 13, color: C.accent4 });
T(s, "Phone 화면에서 볼 것", { x: 7.8, y: 1.6, w: 4.9, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
T(s, [
  { text: "제목이 두 줄 넘게 꺾이지 않는가", options: { bullet: true, breakLine: true } },
  { text: "Hero 사진의 중심이 잘리지 않는가", options: { bullet: true, breakLine: true } },
  { text: "도면 선과 글씨가 읽히는가", options: { bullet: true, breakLine: true } },
  { text: "섹션 순서가 PC와 같은가", options: { bullet: true, breakLine: true } },
  { text: "메뉴가 햄버거 버튼으로 열리는가", options: { bullet: true } },
], { x: 7.8, y: 2.15, w: 4.9, h: 3.2, fontSize: 15, paraSpaceAfter: 10 });
T(s, "마지막 확인은 반드시 내 휴대폰으로 미리보기 링크를 열어서.", { x: 7.8, y: 5.6, w: 4.9, h: 0.6, fontSize: 13, bold: true, color: C.accent6 });
s.addNotes("심사위원과 채용담당자는 판넬의 QR을 휴대폰으로 찍어서 들어옵니다. 그러니 Phone 화면이 실제 기준입니다. Framer는 Desktop, Tablet, Phone 크기를 따로 조정할 수 있습니다. 한 크기에서 고친 글자 크기가 다른 크기에는 적용되지 않을 수 있으니 세 화면을 모두 봅니다.");

// 13 게시 전 설정
s = content("점검·게시", "게시 전에 사이트 정보를 채운다", "STEP 05 · Site Settings");
const meta = [
  ["사이트 제목", "작품명 | 이름", "브라우저 탭과 검색 결과에 보이는 이름"],
  ["설명", "한 줄 컨셉", "카카오톡 · 검색 미리보기의 설명 문장"],
  ["파비콘", "정사각 PNG", "탭 아이콘. 브랜드 시그니처를 작게"],
  ["소셜 이미지", "1200×630", "링크를 공유할 때 뜨는 대표 이미지"],
];
meta.forEach(([a, b, c], i) => {
  const y = 1.6 + i * 1.15;
  box(s, M, y, 12.1, 0.95, i % 2 ? C.background1 : C.background2, "meta" + i);
  T(s, a, { x: M + 0.35, y, w: 2.6, h: 0.95, valign: "middle", fontSize: 18, bold: true, color: C.text2 });
  T(s, b, { x: M + 3.1, y, w: 2.8, h: 0.95, valign: "middle", fontSize: 17, bold: true, color: C.accent1 });
  T(s, c, { x: M + 6.1, y, w: 5.8, h: 0.95, valign: "middle", fontSize: 14 });
});
T(s, "프로젝트 이름 옆 설정(Site Settings)에서 입력합니다. 권장 규격은 Framer 화면의 안내를 우선합니다.", { x: M, y: 6.35, w: 12, h: 0.4, fontSize: 13, color: C.accent4 });
s.addNotes("이걸 빼먹으면 카톡으로 링크를 보냈을 때 템플릿 제작자 이름과 샘플 사진이 뜹니다. 실제로 가장 많이 나오는 실수입니다. 제목, 설명, 파비콘, 소셜 이미지 네 가지를 꼭 바꾸고 게시하세요.");

// 14 게시
s = content("점검·게시", "Publish를 눌러야 세상에 나간다", "STEP 05 · 게시");
const pub = [["Publish 클릭", "오른쪽 위 버튼"], ["주소 이름 정하기", "영문 이름 또는 작품명으로"], ["게시 완료", "발급된 링크를 휴대폰으로 열어 확인"], ["고치면 다시 Publish", "누르기 전까지 방문자는 이전 버전을 봄"]];
pub.forEach(([h, d], i) => {
  const x = M + i * 3.05;
  box(s, x, 1.7, 2.8, 2.3, i === 3 ? C.text2 : C.background2, "pub" + i);
  T(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: 1.9, w: 1, h: 0.4, fontSize: 14, bold: true, color: i === 3 ? C.accent3 : C.accent1 });
  T(s, h, { x: x + 0.3, y: 2.35, w: 2.3, h: 0.8, fontSize: 18, bold: true, color: i === 3 ? C.background1 : C.text2 });
  T(s, d, { x: x + 0.3, y: 3.15, w: 2.3, h: 0.7, fontSize: 13, color: i === 3 ? C.accent5 : C.text1 });
});
box(s, M, 4.4, 12.1, 1.9, C.background2, "free");
T(s, "무료 플랜으로 충분합니다", { x: M + 0.35, y: 4.6, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
T(s, "무료 게시는 Framer 주소(예: 이름.framer.website 형식)와 'Made in Framer' 배지가 붙습니다. 개인 도메인 연결과 배지 제거는 유료 플랜이 필요하며, 이 수업의 필수 조건은 아닙니다. 주소 형식과 가격은 바뀔 수 있으니 게시 화면의 안내를 기준으로 합니다.", { x: M + 0.35, y: 5.05, w: 11.4, h: 1.1, fontSize: 14 });
s.addNotes("자동 저장과 게시는 다릅니다. 고치고 나서 Publish를 다시 누르지 않으면 판넬 QR로 들어온 사람은 옛날 버전을 봅니다. 무료로 게시하면 Framer 주소와 배지가 붙는데, 수업 과제로는 문제없습니다. 개인 도메인은 졸업 후 취업용으로 원하는 학생만 따로 하세요.");

// 15 QR
s = content("점검·게시", "주소를 QR로 바꿔 판넬에 붙인다", "STEP 05 · QR");
const qr = [["주소 복사", "게시된 사이트를 Chrome으로 열기"], ["QR 만들기", "주소창 공유 메뉴의 'QR 코드 만들기' → 다운로드"], ["판넬 배치", "가로 3cm 이상, 주변에 여백"], ["찍어 보기", "출력 후 휴대폰 3대 이상으로 인식 확인"]];
qr.forEach(([h, d], i) => {
  const y = 1.6 + i * 1.18;
  num(s, i + 1, M, y + 0.05);
  T(s, h, { x: M + 0.8, y, w: 5.8, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
  T(s, d, { x: M + 0.8, y: y + 0.48, w: 5.8, h: 0.5, fontSize: 14 });
});
// QR motif
const qx = 8.6, qy = 1.7, cell = 0.36;
box(s, qx - 0.35, qy - 0.35, 3.95, 3.95, C.background2, "qrbg");
const pat = ["1111101011", "1000101101", "1011101010", "1011100111", "1000100010", "1111101011", "0001010100", "1101111011", "0110100110", "1011010111"];
pat.forEach((row, r) => [...row].forEach((c, k) => {
  if (c === "1") s.addShape(pres.shapes.RECTANGLE, { x: qx + k * cell, y: qy + r * cell, w: cell, h: cell, fill: { color: C.text2 }, line: { color: C.text2 }, objectName: `q${r}_${k}` });
}));
T(s, "그림은 예시입니다. 실제 QR은 각자 생성합니다.", { x: 8.25, y: 5.75, w: 4.4, h: 0.35, fontSize: 11, color: C.accent4 });
s.addNotes("QR 생성 사이트는 광고가 많고, 일부는 몇 달 뒤 유료로 바뀌어 QR이 작동하지 않게 만듭니다. Chrome 브라우저에 있는 QR 만들기 기능을 쓰면 주소가 그대로 들어간 QR이 나옵니다. 출력하면 반드시 여러 휴대폰으로 찍어 보세요.");

// 16 최종 체크리스트
pres.addSection({ title: "정리" });
s = content("정리", "제출 전 체크리스트", "CHECK");
const checks = [
  "6개 섹션이 순서대로 있다", "템플릿 샘플 글 · 사진이 하나도 남지 않았다", "브랜드 컬러 3색 · 폰트 2종 안에 있다",
  "영상 자동재생이 꺼져 있다", "휴대폰에서 첫 화면이 5초 안에 뜬다", "제목 · 설명 · 파비콘 · 소셜 이미지를 바꿨다",
  "마지막 수정 후 다시 Publish 했다", "QR이 휴대폰 3대에서 인식된다", "About에 사용한 AI 도구를 적었다",
];
checks.forEach((t, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  const x = M + col * 4.1, y = 1.6 + row * 1.35;
  box(s, x, y, 3.85, 1.1, C.background2, "chk" + i);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: y + 0.35, w: 0.4, h: 0.4, rectRadius: 0.06, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.5 }, objectName: "cb" + i });
  T(s, t, { x: x + 0.85, y, w: 2.85, h: 1.1, valign: "middle", fontSize: 14, color: C.text2 });
});
box(s, M, 5.75, 12.1, 0.8, C.text2, "submit");
T(s, "제출: 공개 URL · QR 이미지(PNG) · PC와 휴대폰 첫 화면 캡처 각 1장", { x: M + 0.35, y: 5.75, w: 11.4, h: 0.8, valign: "middle", fontSize: 16, bold: true, color: C.background1 });
s.addNotes("이 아홉 개가 모두 체크되면 제출합니다. 특히 템플릿 샘플 글이 페이지 아래쪽 푸터에 남아 있는 경우가 많습니다. 끝까지 스크롤해서 확인하세요.");

// 17 자주 막히는 곳
s = content("정리", "자주 막히는 곳", "TROUBLESHOOTING");
const tr = [
  ["고쳤는데 사이트에 안 보인다", "Publish를 다시 누르지 않았다. 다시 게시 후 새로고침"],
  ["이미지가 안 올라간다", "용량이 크다. 2400px · 1MB 안팎으로 줄여서 다시"],
  ["휴대폰에서 글자가 넘친다", "Phone 탭에서 해당 글의 크기를 따로 줄인다"],
  ["영상이 재생되지 않는다", "YouTube 공개 범위와 퍼가기 허용 설정을 확인"],
  ["카톡 미리보기가 템플릿 그대로다", "소셜 이미지 · 설명 교체 후 재게시. 반영까지 시간이 걸릴 수 있음"],
];
tr.forEach(([q, a], i) => {
  const y = 1.55 + i * 0.98;
  box(s, M, y, 12.1, 0.82, i % 2 ? C.background1 : C.background2, "tr" + i);
  T(s, q, { x: M + 0.35, y, w: 4.9, h: 0.82, valign: "middle", fontSize: 15, bold: true, color: C.text2 });
  T(s, a, { x: M + 5.4, y, w: 6.5, h: 0.82, valign: "middle", fontSize: 14 });
});
s.addNotes("질문이 가장 많이 나오는 다섯 가지입니다. 첫 번째가 압도적으로 많습니다. 고쳤으면 다시 Publish.");

// 18 마무리
s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "정리" });
T(s, "이번 주 과제", { x: M, y: 1.6, w: 8, h: 0.4, fontSize: 14, bold: true, color: C.accent3 });
s.addText("템플릿 복제 + Hero · Video\n두 섹션까지 채워 오기", { placeholder: "title" });
s.addText("다음 차시 · 나머지 섹션, 휴대폰 점검, 게시와 QR까지", { placeholder: "body" });
T(s, "자세한 단계는 학생용 구축 워크북을 참고합니다.", { x: M, y: 6.6, w: 10, h: 0.3, fontSize: 11, color: C.accent2 });
s.addNotes("이번 주는 템플릿을 내 프로젝트로 복제하고 Hero와 Video 두 섹션만 채워 옵니다. 나머지는 다음 시간에 같이 합니다. 단계별 설명은 워크북에 모두 있습니다.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
