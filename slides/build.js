// NGS 외래성 바이러스 세미나 - 슬라이드 초안 생성
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "목록에 없는 바이러스 - NGS 기반 외래성 바이러스 부정시험";

const NAVY = "000080";
const ORANGE = "FFA500";
const WHITE = "FFFFFF";
const LIGHT = "F2F2F2";
const LINE = "D9D9D9";
const GRAY = "595959";
const MUTED = "7F7F7F";
const FONT = "Malgun Gothic";
const NUMF = "Arial";
const W = 13.333;

let page = 0;

// ---------- 공통 ----------
function base(opts) {
  const s = pres.addSlide();
  page += 1;
  s.background = { color: opts.bg || WHITE };
  if (opts.section) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.4, w: 0.12, h: 0.12, fill: { color: ORANGE }, line: { color: ORANGE, width: 0 } });
    s.addText(opts.section, { x: 0.82, y: 0.3, w: 8, h: 0.32, fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, valign: "middle", isTextBox: true });
  }
  if (opts.title) {
    s.addText(opts.title, { x: 0.6, y: 0.62, w: 12.1, h: 0.8, fontFace: FONT, fontSize: 32, bold: true, color: NAVY, margin: 0, valign: "middle", isTextBox: true });
  }
  if (opts.cite) {
    s.addText(opts.cite, { x: 0.6, y: 6.98, w: 11.3, h: 0.34, fontFace: FONT, fontSize: 8.5, color: MUTED, margin: 0, valign: "middle", isTextBox: true });
  }
  if (!opts.noPage) {
    s.addText(String(page), { x: 12.23, y: 6.98, w: 0.5, h: 0.34, fontFace: NUMF, fontSize: 9, color: MUTED, align: "right", margin: 0, valign: "middle", isTextBox: true });
  }
  if (opts.notes) s.addNotes(opts.notes);
  return s;
}

function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: FONT, fontSize: 14, color: NAVY, margin: 0, valign: "top", isTextBox: true }, o));
}

function rect(s, x, y, w, h, fill, lineColor) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: lineColor || fill, width: lineColor ? 1 : 0 } });
}

function hline(s, x1, x2, y, color, width) {
  s.addShape(pres.shapes.LINE, { x: x1, y, w: x2 - x1, h: 0, line: { color: color || LINE, width: width || 1 } });
}

function vline(s, x, y1, y2, color, width) {
  s.addShape(pres.shapes.LINE, { x, y: y1, w: 0, h: y2 - y1, line: { color: color || LINE, width: width || 1 } });
}

function arrow(s, x1, y1, x2, y2, color) {
  s.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: color || NAVY, width: 1.5, endArrowType: "triangle" },
  });
}

function dot(s, cx, cy, d, color) {
  s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color }, line: { color, width: 0 } });
}

// 키워드: 세로 스택 (번호 + 굵은 키워드 + 선택적 설명), 사이에 얇은 구분선
function kwStack(s, items, x, y, w, gap) {
  gap = gap || 0.95;
  items.forEach((it, i) => {
    const yy = y + i * gap;
    if (i > 0) hline(s, x, x + w, yy - 0.14, LINE, 0.75);
    txt(s, String(i + 1).padStart(2, "0"), { x, y: yy, w: 0.5, h: 0.4, fontFace: NUMF, fontSize: 12, bold: true, color: it.hot ? ORANGE : MUTED });
    txt(s, it.k, { x: x + 0.55, y: yy - 0.04, w: w - 0.55, h: 0.45, fontSize: 18, bold: true });
    if (it.d) txt(s, it.d, { x: x + 0.55, y: yy + 0.38, w: w - 0.55, h: 0.4, fontSize: 11.5, color: GRAY });
  });
}

// 키워드: 가로 칩 (둥근 사각형, 연회색 / 강조는 오렌지)
function kwChips(s, items, x, y, w, h) {
  h = h || 0.55;
  const gap = 0.25;
  const cw = (w - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const hot = typeof it === "object" && it.hot;
    const label = typeof it === "object" ? it.k : it;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + i * (cw + gap), y, w: cw, h, rectRadius: 0.08, fill: { color: hot ? ORANGE : LIGHT }, line: { color: hot ? ORANGE : LIGHT, width: 0 } });
    txt(s, label, { x: x + i * (cw + gap), y, w: cw, h, fontSize: 14, bold: true, align: "center", valign: "middle" });
  });
}

// 키워드: 가로 3단 (작은 라벨 + 큰 값)
function kwTriple(s, items, x, y, w) {
  const cw = w / items.length;
  items.forEach((it, i) => {
    const xx = x + i * cw;
    if (i > 0) vline(s, xx - 0.1, y, y + 1.15, LINE, 0.75);
    txt(s, it.label, { x: xx + 0.1, y, w: cw - 0.3, h: 0.35, fontSize: 11.5, color: GRAY });
    txt(s, it.value, { x: xx + 0.1, y: y + 0.38, w: cw - 0.3, h: 0.75, fontSize: it.size || 24, bold: true, color: NAVY, fontFace: it.num ? NUMF : FONT });
  });
}

// 논문 카드: 서지 + 첫 페이지/대표 Figure 삽입 영역(점선)
function paperCard(s, x, y, w, h, p) {
  rect(s, x, y, w, h, LIGHT);
  txt(s, p.journal, { x: x + 0.22, y: y + 0.2, w: w - 0.44, h: 0.3, fontSize: 11, bold: true, fontFace: NUMF });
  txt(s, p.title, { x: x + 0.22, y: y + 0.52, w: w - 0.44, h: 0.95, fontSize: 10.5, italic: true, fontFace: NUMF, color: NAVY });
  txt(s, p.authors, { x: x + 0.22, y: y + 1.45, w: w - 0.44, h: 0.3, fontSize: 9, color: GRAY, fontFace: NUMF });
  const sy = y + 1.85;
  const sh = h - 2.05;
  s.addShape(pres.shapes.RECTANGLE, { x: x + 0.22, y: sy, w: w - 0.44, h: sh, fill: { color: WHITE }, line: { color: "A6A6A6", width: 1, dashType: "dash" } });
  txt(s, p.slot || "원문 첫 페이지 캡처 삽입", { x: x + 0.3, y: sy, w: w - 0.6, h: sh, fontSize: 10.5, color: MUTED, align: "center", valign: "middle" });
}

// 수평 타임라인
function timeline(s, y, x1, x2, nodes) {
  hline(s, x1, x2, y, NAVY, 2);
  nodes.forEach((n) => {
    dot(s, n.x, y, n.hot ? 0.26 : 0.2, n.hot ? ORANGE : NAVY);
    const bw = n.w || 1.9;
    if (n.up) {
      txt(s, n.year, { x: n.x - bw / 2, y: y - 1.35, w: bw, h: 0.35, fontSize: 14, bold: true, align: "center", fontFace: NUMF });
      txt(s, n.label, { x: n.x - bw / 2, y: y - 1.0, w: bw, h: 0.75, fontSize: 11, align: "center", color: GRAY });
    } else {
      txt(s, n.year, { x: n.x - bw / 2, y: y + 0.28, w: bw, h: 0.35, fontSize: 14, bold: true, align: "center", fontFace: NUMF });
      txt(s, n.label, { x: n.x - bw / 2, y: y + 0.63, w: bw, h: 0.75, fontSize: 11, align: "center", color: GRAY });
    }
  });
}

// 표 셀 헬퍼
const H = (t) => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } });
const C = (t, o) => ({ text: t, options: Object.assign({ color: NAVY, valign: "middle" }, o || {}) });
const HOT = (t) => ({ text: t, options: { fill: { color: ORANGE }, color: NAVY, bold: true, align: "center", valign: "middle" } });
const SOFT = (t) => ({ text: t, options: { fill: { color: LIGHT }, color: NAVY, align: "center", valign: "middle" } });
const NA = () => ({ text: "—", options: { color: MUTED, align: "center", valign: "middle" } });
const TB = { type: "solid", pt: 0.75, color: LINE };

// 섹션 구분 슬라이드
function divider(num, title, kws, minutes, notes) {
  const s = base({ bg: LIGHT, notes });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.05, w: 0.18, h: 0.18, fill: { color: ORANGE }, line: { color: ORANGE, width: 0 } });
  txt(s, num, { x: 0.9, y: 2.35, w: 3, h: 1.3, fontSize: 80, bold: true, fontFace: NUMF });
  txt(s, title, { x: 0.9, y: 3.75, w: 11, h: 0.9, fontSize: 40, bold: true });
  txt(s, kws.join("   ·   "), { x: 0.9, y: 4.8, w: 11, h: 0.5, fontSize: 18, color: GRAY });
  if (minutes) txt(s, minutes, { x: 0.9, y: 5.45, w: 4, h: 0.4, fontSize: 12, color: MUTED });
  return s;
}

// =====================================================================
// 1. 표지
{
  const s = base({
    bg: LIGHT, noPage: true,
    notes: "표지. 40분 발표, Q&A 별도.\n발표자·소속·일자는 확정 후 기입.",
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 1.75, w: 0.18, h: 0.18, fill: { color: ORANGE }, line: { color: ORANGE, width: 0 } });
  txt(s, "목록에 없는 바이러스", { x: 0.9, y: 2.1, w: 11.5, h: 1.2, fontSize: 54, bold: true });
  txt(s, "NGS 기반 외래성 바이러스 부정시험", { x: 0.9, y: 3.35, w: 11.5, h: 0.7, fontSize: 28, color: NAVY });
  txt(s, "ICH Q5A(R2) · Ph. Eur. 2.6.41 이후의 규제 좌표", { x: 0.9, y: 4.15, w: 11.5, h: 0.5, fontSize: 16, color: GRAY });
  txt(s, "2026", { x: 0.9, y: 6.2, w: 4, h: 0.4, fontSize: 14, color: MUTED, fontFace: NUMF });
}

// =====================================================================
// 01 인트로
divider("01", "인트로", ["오염 사건", "규제 전환", "정의"], "7분",
  "파트 1 (7분). 왜 지금 이 주제인가: 사건 → 규제 → 대체 범위 → 용어 정의 순.");

// 1-2 반복되는 오염
{
  const s = base({
    section: "01 인트로", title: "반복되는 오염",
    cite: "CDC Historical concerns (SV40) · IOM 2002 · Garnick RL. Dev Biol Stand 1996;88:49 · Nims RW et al. BioPharm Int 2008 · Genzyme Form 10-K FY2009 · Barone PW et al. Nat Biotechnol 2020;38:563",
    notes: [
      "요지: 외래성 바이러스 오염은 과거형이 아니라 반복되는 산업 리스크다.",
      "- SV40: 1955–1963 폴리오 백신, 원숭이 신장세포 유래 (CDC; IOM 2002).",
      "- MVM: Genentech CHO 공정 2건 (Garnick 1996). ⚠ 정확한 발생 연도 미확인 → '1990년대 보고'로만 표기.",
      "- Cache Valley virus: 2000–2004 CHO 생산 의약품 (Nims 2008).",
      "- Vesivirus 2117: 2009 Genzyme Allston, Cerezyme·Fabrazyme 공급 중단 (Genzyme 10-K).",
      "- PCV1: 2010 Rotarix (다음 장).",
      "- 하단 수치: CAACB 조사 18건, 원인 바이러스 9종, 13건 상업 생산 중 (Barone 2020).",
    ].join("\n"),
  });
  timeline(s, 2.95, 0.9, 12.45, [
    { x: 1.6, year: "1955–63", label: "SV40\n폴리오 백신 · 원숭이 신장세포", up: false, w: 2.2 },
    { x: 4.2, year: "1990년대 보고", label: "MVM\nCHO 공정", up: false, w: 2.2 },
    { x: 6.8, year: "2000–04", label: "Cache Valley virus\nCHO 생산 의약품", up: false, w: 2.2 },
    { x: 9.4, year: "2009", label: "Vesivirus 2117\nCHO · 공급 중단", up: false, w: 2.2 },
    { x: 11.8, year: "2010", label: "PCV1\n허가 백신 (Rotarix)", up: false, w: 1.9, hot: true },
  ]);
  txt(s, "CAACB 산업계 조사", { x: 0.9, y: 4.75, w: 5, h: 0.3, fontSize: 11.5, color: GRAY });
  const stats = [
    { n: "18", u: "건", l: "바이러스 오염 사건" },
    { n: "9", u: "종", l: "원인 바이러스" },
    { n: "13", u: "건", l: "상업 생산 단계 발생", hot: true },
  ];
  stats.forEach((st, i) => {
    const x = 0.9 + i * 3.9;
    if (st.hot) rect(s, x - 0.05, 5.2, 1.25, 0.95, ORANGE);
    txt(s, [
      { text: st.n, options: { fontFace: NUMF, fontSize: 48, bold: true } },
      { text: " " + st.u, options: { fontSize: 18, bold: true } },
    ], { x, y: 5.15, w: 2.4, h: 1.0, valign: "middle" });
    txt(s, st.l, { x: x + 1.55, y: 5.35, w: 2.2, h: 0.7, fontSize: 13, color: GRAY, valign: "middle" });
  });
}

// 1-3 Rotarix
{
  const s = base({
    section: "01 인트로", title: "Rotarix, 2010",
    cite: "Victoria JG et al. J Virol 2010;84:6033 · FDA 발표 2010.03.22 / 05.14 (Medscape 보도) · CIDRAP 2010.05.06 · Dubin G et al. Hum Vaccin Immunother 2013;9:2398",
    notes: [
      "요지: 허가·출하시험을 모두 통과한 제품에서 NGS(메타게놈)가 PCV1을 찾았다. 못 찾은 이유는 '검사 목록에 없었기 때문'.",
      "- 발견: 허가 생백신 8종 메타게놈 + 범미생물 마이크로어레이 재검 (Victoria 2010).",
      "- 2010.03.22 FDA 사용 일시 중단 권고 → 05.06 RotaTeq에서도 PCV1·PCV2 DNA 단편 → 05.07 VRBPAC → 05.14 사용 재개.",
      "- 원인: 1983년 MCB 제작 시 사용한 비조사(non-irradiated) 돼지 trypsin 추정. Vero 세포은행 포함 전 단계에 존재. 인체 감염 근거 없음 (Dubin 2013).",
      "⚠ FDA 원 보도자료 URL 확보 필요 (현재 보도 기사 경유 확인).",
    ].join("\n"),
  });
  txt(s, [
    { text: "출하시험 통과", options: { breakLine: true } },
    { text: "그러나 ", options: {} },
    { text: "PCV1", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 1.75, w: 6.4, h: 1.8, fontSize: 40, bold: true, valign: "top" });
  kwStack(s, [
    { k: "메타게놈 재검", d: "허가 생백신 8종 · 2010" },
    { k: "비조사 돼지 trypsin", d: "1983년 MCB 제작 단계 추정" },
    { k: "인체 감염 근거 없음", d: "위해성 평가 후 사용 재개" },
  ], 0.6, 4.0, 6.0, 0.95);
  // 세로 타임라인
  const x0 = 8.1;
  vline(s, x0, 1.95, 6.35, NAVY, 2);
  const ev = [
    { y: 2.05, d: "2010.03.22", t: "FDA, Rotarix 사용 일시 중단 권고", hot: true },
    { y: 3.45, d: "05.06", t: "RotaTeq에서도 PCV1·PCV2 DNA 단편" },
    { y: 4.85, d: "05.07", t: "FDA 자문위원회(VRBPAC) 심의" },
    { y: 6.2, d: "05.14", t: "사용 재개 권고" },
  ];
  ev.forEach((e) => {
    dot(s, x0, e.y, e.hot ? 0.26 : 0.2, e.hot ? ORANGE : NAVY);
    txt(s, e.d, { x: x0 + 0.35, y: e.y - 0.22, w: 4.2, h: 0.35, fontSize: 15, bold: true, fontFace: NUMF });
    txt(s, e.t, { x: x0 + 0.35, y: e.y + 0.12, w: 4.4, h: 0.4, fontSize: 12, color: GRAY });
  });
}

// 1-4 ICH Q5A(R2)
{
  const s = base({
    section: "01 인트로", title: "ICH Q5A(R2)",
    cite: "ICH Q5A(R2) Step 4, 2023-11-01, §3.2.5.2 · FDA Federal Register 89 FR 1925 (2024-01-11) · EMA/CHMP/ICH/804363/2022 (2024-06-14) · MHLW 2024-02-14",
    notes: [
      "요지: 2023년 개정으로 non-targeted NGS가 in vivo 시험을 정면 비교 없이 대체할 수 있게 되었다.",
      "- 위치: 3.2.5 Molecular Methods › 3.2.5.2 Next generation sequencing.",
      "- 정면 비교를 요구하지 않는 근거: 시험계마다 종말점이 다르고, in vitro·in vivo의 검출 폭이 NGS보다 좁다.",
      "- 시행: FDA 2024-01-11, EU 2024-06-14, 일본 2024-02-14.",
      "⚠ 인용문은 검색 발췌 기준. 원문 PDF와 글자 단위 대조 필요. 절 제목 표기도 확인.",
      "⚠ 식약처 시행 여부·일자 미확인 → 발표 시 '확인 중'으로 언급하거나 확인 후 추가.",
    ].join("\n"),
  });
  txt(s, "“", { x: 0.55, y: 1.45, w: 1, h: 1.1, fontSize: 96, bold: true, color: ORANGE, fontFace: NUMF });
  txt(s, "Non-targeted NGS can be used to replace the in vivo assays and supplement or replace the in vitro cell culture assays, without a head-to-head comparison …", {
    x: 1.45, y: 1.75, w: 11.1, h: 2.3, fontSize: 26, italic: true, fontFace: NUMF, color: NAVY, valign: "top",
  });
  txt(s, "ICH Q5A(R2) · 3.2.5.2 Next generation sequencing", { x: 1.45, y: 4.05, w: 10, h: 0.35, fontSize: 12, color: GRAY });
  kwTriple(s, [
    { label: "Step 4", value: "2023. 11", num: true },
    { label: "미국 · EU · 일본 시행", value: "2024" , num: true },
    { label: "기존 시험과의 비교", value: "정면 비교 불요" },
  ], 1.45, 4.95, 11.1);
}

// 1-5 국제 동향
{
  const s = base({
    section: "01 인트로", title: "이미 현행 규제",
    cite: "WHO TRS 978 Annex 3 (2013) · IABS NGS 컨퍼런스 보고 Biologicals 2018;55:1 / 2020;67:94 / 2023;83:101696 / 2025;92:101859 · 식약처 NGS 부정시험법 정보집 (2021) · WHO/BS/2024.2471 · EDQM 2025",
    notes: [
      "요지: NGS는 전망이 아니라 현행 규제 선택지다. Ph. Eur. 2.6.41은 2026-04-01부터 시행 중.",
      "- 2013 WHO TRS 978 Annex 3: deep sequencing의 가능성과 Rotarix 사례를 이미 언급.",
      "- 2013.11 PDA/FDA 워크숍(Bethesda) → IABS 컨퍼런스 1차 2017 Rockville, 2차 2019 Ghent, 3차 2022 Rockville, 4차 2024 Frankfurt.",
      "- 2017 Ph. Eur. 5.2.3·2.6.16 개정: 위해성 평가 도입, 일부 동물시험 삭제 (⚠ 삭제 항목 원문 확인).",
      "- 2021 식약처(식품의약품안전평가원) 「백신 생산용 세포주의 NGS 기반 외래성 바이러스 부정시험법 정보집」 (⚠ 발간일 확인).",
      "- 2024 WHO ECBS 1st International Reference Panel (HTS 외래성 바이러스 검출, 7종).",
      "- 2025.03 Ph. Eur. 2.6.41 채택, 12.2 수재 → 2026.04.01 시행.",
    ].join("\n"),
  });
  const X = (yr) => 1.3 + (yr - 2010) * (10.9 / 16.25);
  timeline(s, 3.95, 0.9, 12.6, [
    { x: X(2010), year: "2010", label: "Rotarix PCV1", up: true, w: 1.3 },
    { x: X(2013), year: "2013", label: "WHO TRS 978", up: false, w: 1.3 },
    { x: X(2017), year: "2017", label: "IABS 1차\nPh. Eur. 5.2.3 개정", up: true, w: 1.3 },
    { x: X(2019), year: "2019", label: "IABS 2차", up: false, w: 1.3 },
    { x: X(2021), year: "2021", label: "식약처\nNGS 정보집", up: true, w: 1.3 },
    { x: X(2022), year: "2022", label: "IABS 3차", up: false, w: 1.3 },
    { x: X(2023), year: "2023", label: "ICH\nQ5A(R2)", up: true, w: 1.3 },
    { x: X(2024), year: "2024", label: "IABS 4차\nWHO 표준패널", up: false, w: 1.3 },
    { x: X(2025), year: "2025", label: "2.6.41\n채택", up: true, w: 1.3 },
    { x: X(2026.25), year: "2026.04", label: "2.6.41\n시행", up: false, w: 1.3, hot: true },
  ]);
  kwChips(s, [{ k: "Ph. Eur. 2.6.41 시행", hot: true }, "WHO 국제 표준 패널", "IABS 4차 합의"], 0.9, 6.05, 11.5, 0.55);
}

// 1-6 무엇이 대체되나
{
  const s = base({
    section: "01 인트로", title: "무엇이 대체되나",
    cite: "ICH Q5A(R2) §3.2.5.2 · Ph. Eur. 2.6.41 (EDQM 2025) · Khan AS et al. Biologicals 2025;92:101859 (4차 IABS 컨퍼런스 보고)",
    notes: [
      "요지: 문서 세 곳이 같은 방향. In vivo와 표적 PCR은 '대체', in vitro 세포배양은 '보완 또는 대체'.",
      "- ICH Q5A(R2): non-targeted NGS로 in vivo 대체, in vitro 보완·대체 / Table 3 동물(항체생산) 시험은 NAT·NGS로 대체.",
      "- Ph. Eur. 2.6.41: in vivo·NAT 대체, 세포배양 in vitro 보완·대체.",
      "- 4차 컨퍼런스 합의: 'readiness of NGS to replace the in vivo ... and PCR assays, and to supplement or replace the in vitro cell-based assays, based on a suitable validation package.'",
      "- '—'는 해당 문서에서 명시 여부를 확인하지 못한 칸.",
      "⚠ ICH의 종특이 PCR 대체 문구는 검색 발췌 기준 → 원문 대조 필요.",
      "- 3Rs는 부수 효과로만 언급. 1차 논거는 검출 폭(성능).",
    ].join("\n"),
  });
  const rows = [
    [H("기존 시험"), H("ICH Q5A(R2)"), H("Ph. Eur. 2.6.41"), H("IABS 4차 합의")],
    [C("In vivo 시험", { bold: true }), HOT("대체"), HOT("대체"), HOT("대체")],
    [C("항체생산시험", { bold: true }), HOT("대체"), NA(), NA()],
    [C("종특이 PCR", { bold: true }), HOT("대체"), HOT("대체"), HOT("대체")],
    [C("In vitro 지시세포", { bold: true }), SOFT("보완 · 대체"), SOFT("보완 · 대체"), SOFT("보완 · 대체")],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 8.4, colW: [2.4, 2.0, 2.0, 2.0], rowH: 0.78, fontFace: FONT, fontSize: 15, border: TB });
  txt(s, "— 해당 문서에서 명시 확인 전", { x: 0.6, y: 5.85, w: 6, h: 0.3, fontSize: 10, color: MUTED });
  kwStack(s, [
    { k: "동물시험 대체", d: "in vivo · 항체생산시험", hot: true },
    { k: "표적 PCR 대체", d: "종특이 NAT" },
    { k: "세포배양 보완", d: "in vitro 지시세포 시험" },
  ], 9.6, 1.95, 3.1, 1.25);
}

// 1-7 정의
{
  const s = base({
    section: "01 인트로", title: "외래성의 기준점",
    cite: "ICH Q5A(R1) Glossary (R2에 승계) · ICH Q5A(R2) 2023",
    notes: [
      "요지: '외래(外來)'의 기준은 환자가 아니라 세포기질 유전체.",
      "- Adventitious: 'Unintentionally introduced contaminant viruses.'",
      "- Endogenous: 'Viral entity whose genome is part of the germ line of the species of origin of the cell line and is covalently integrated into the genome of animal from which the parental cell line was derived.'",
      "- 예: CHO의 레트로바이러스 유사입자(RVLP)는 바이러스이지만 외래성이 아니다. FBS 유래 BVDV는 외래성이다.",
      "- 규제상 취급: 외래성 = 부정시험(없음 확인) / 내인성 = 존재 전제, 정량 + 제거·불활화 공정 검증.",
      "- 예고: 파트 4의 EVE 사례에서 이 경계로 돌아온다.",
      "⚠ Glossary 문구는 R1 기준. R2 원문 동일 여부 대조.",
    ].join("\n"),
  });
  txt(s, [
    { text: "기준점 = ", options: {} },
    { text: "세포기질 유전체", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 1.65, w: 12, h: 0.75, fontSize: 30, bold: true });
  const cols = [
    { x: 0.6, en: "Adventitious", ko: "외래성", key: "비의도적 혼입",
      q: "“Unintentionally introduced contaminant viruses.”", ex: "FBS 유래 BVDV · trypsin 유래 PCV" },
    { x: 6.85, en: "Endogenous", ko: "내인성", key: "생식계열 통합",
      q: "“…genome is part of the germ line of the species of origin of the cell line …”", ex: "CHO 레트로바이러스 유사입자 (RVLP)" },
  ];
  cols.forEach((c) => {
    rect(s, c.x, 2.7, 5.9, 2.85, LIGHT);
    txt(s, c.ko + "  ", { x: c.x + 0.3, y: 2.9, w: 2, h: 0.5, fontSize: 22, bold: true });
    txt(s, c.en, { x: c.x + 1.55, y: 2.98, w: 2.5, h: 0.4, fontSize: 13, color: GRAY, fontFace: NUMF, italic: true });
    txt(s, c.key, { x: c.x + 0.3, y: 3.5, w: 5.3, h: 0.45, fontSize: 18, bold: true });
    txt(s, c.q, { x: c.x + 0.3, y: 4.0, w: 5.3, h: 0.8, fontSize: 12, italic: true, color: GRAY, fontFace: NUMF });
    txt(s, "예  " + c.ex, { x: c.x + 0.3, y: 4.9, w: 5.3, h: 0.4, fontSize: 12.5 });
  });
  txt(s, [
    { text: "부정시험", options: { bold: true } },
    { text: "  없음을 확인          ", options: { color: GRAY } },
    { text: "vs", options: { color: MUTED } },
    { text: "          정량 · 제거 검증", options: { bold: true } },
    { text: "  존재를 전제", options: { color: GRAY } },
  ], { x: 0.6, y: 5.85, w: 12.1, h: 0.5, fontSize: 16, align: "center", valign: "middle" });
}

// =====================================================================
// 02 부정시험
divider("02", "부정시험", ["In vitro", "In vivo", "분자적 방법"], "7분",
  "파트 2 (7분). 청중이 발표자보다 이 분야의 실무 전문가임을 전제. 각 장은 가이드라인 원문 요약 + 출처만. 발표자 해석은 마지막 '세 가지 한계' 한 장에 모은다.");

// 2-1 단계별 시험
{
  const s = base({
    section: "02 부정시험", title: "단계별 시험",
    cite: "ICH Q5A(R2) Table 1 “Examples of virus tests to be performed once at various cell levels” · 미가공 벌크 시험 조항",
    notes: [
      "요지: 세포은행과 LIVCA는 1회, 미가공 벌크는 배치마다.",
      "- Table 1 열: MCB, WCB, LIVCA(시험관내 세포 연령 한계의 세포). 행: 레트로·내인성 시험(감염성, EM, RT), 외래성 시험(in vitro, in vivo, 항체생산, 분자법).",
      "- 미가공 벌크: 배치별 외래성 바이러스 시험 (in vitro 지시세포 또는 non-targeted NGS).",
      "⚠ 하단 점선 영역에 ICH Q5A(R2) Table 1 원문을 캡처해 넣을 것. 셀별 +/− 값은 검증 단계에서 읽지 못했으므로 원문 그대로 사용.",
      "⚠ 미가공 벌크 조항의 절 번호 확인.",
    ].join("\n"),
  });
  const steps = [
    { k: "세포은행", d: "MCB · WCB", f: "1회" },
    { k: "LIVCA", d: "시험관내 세포 연령 한계", f: "1회" },
    { k: "미가공 벌크", d: "in vitro 또는 NGS", f: "배치마다", hot: true },
  ];
  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.15;
    rect(s, x, 1.75, 3.7, 1.55, st.hot ? ORANGE : LIGHT);
    txt(s, st.k, { x: x + 0.25, y: 1.9, w: 2.3, h: 0.45, fontSize: 20, bold: true });
    txt(s, st.f, { x: x + 2.3, y: 1.92, w: 1.2, h: 0.45, fontSize: 16, bold: true, align: "right" });
    txt(s, st.d, { x: x + 0.25, y: 2.5, w: 3.3, h: 0.4, fontSize: 12.5, color: st.hot ? NAVY : GRAY });
    if (i < 2) arrow(s, x + 3.75, 2.52, x + 4.1, 2.52);
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.6, w: 12.1, h: 3.1, fill: { color: WHITE }, line: { color: "A6A6A6", width: 1, dashType: "dash" } });
  txt(s, "ICH Q5A(R2) Table 1 원문 캡처 삽입", { x: 0.6, y: 3.6, w: 12.1, h: 3.1, fontSize: 13, color: MUTED, align: "center", valign: "middle" });
}

// 2-2 In vitro
{
  const s = base({
    section: "02 부정시험", title: "In vitro: 감염성의 직접 증거",
    cite: "FDA. Characterization and Qualification of Cell Substrates … for Viral Vaccines (2010) · ICH Q5A(R2) · WHO TRS 978 Annex 3",
    notes: [
      "요지: 지시세포 시험은 '복제 가능한 바이러스'를 직접 보는 시험. NGS가 주지 않는 정보(감염성)를 준다.",
      "- 지시세포 (FDA 2010): 인간 이배체 세포(예: MRC-5), 원숭이 신장세포(예: Vero), 생산세포와 동일 종·조직의 세포.",
      "- 판독: 배양 중 세포변성효과(CPE) 관찰 → 종료 시점 혈구흡착(hemadsorption)·혈구응집(hemagglutination).",
      "⚠ 관찰 기간 '28일(2주 계대)'은 출처(Q5A(R2) vs Ph. Eur.)가 미확정 → 슬라이드에 기재하지 않음.",
      "- 파트 4 EVE 사례에서 '서열 ≠ 감염성'으로 다시 연결.",
    ].join("\n"),
  });
  const cols = [
    { k: "지시세포 3종", lines: ["인간 이배체 (MRC-5)", "원숭이 신장 (Vero)", "생산세포와 동일 종 · 조직"] },
    { k: "CPE 관찰", lines: ["배양 기간 중", "세포변성효과 판독"] },
    { k: "혈구흡착 · 혈구응집", lines: ["종료 시점 판독", "비세포변성 바이러스 포착"] },
  ];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    txt(s, String(i + 1).padStart(2, "0"), { x, y: 2.0, w: 1, h: 0.5, fontSize: 28, bold: true, fontFace: NUMF, color: i === 0 ? ORANGE : MUTED });
    txt(s, c.k, { x, y: 2.65, w: 3.7, h: 0.55, fontSize: 21, bold: true });
    hline(s, x, x + 3.6, 3.35, LINE, 0.75);
    txt(s, c.lines.map((l, j) => ({ text: l, options: { breakLine: j < c.lines.length - 1 } })), { x, y: 3.55, w: 3.7, h: 1.6, fontSize: 14, color: GRAY, paraSpaceAfter: 6 });
    if (i < 2) arrow(s, x + 3.7, 2.92, x + 4.05, 2.92, MUTED);
  });
  txt(s, "복제 가능한 바이러스만 신호를 낸다", { x: 0.6, y: 5.6, w: 12.1, h: 0.6, fontSize: 20, bold: true, color: NAVY });
}

// 2-3 In vivo
{
  const s = base({
    section: "02 부정시험", title: "In vivo: 문서마다 다르다",
    cite: "ICH Q5A(R2) §3.2.3 · FDA. Characterization and Qualification of Cell Substrates … for Viral Vaccines (2010)",
    notes: [
      "요지: 동물 구성은 문서마다 다르다. 재조합 단백질(ICH)과 백신 세포기질(FDA 2010)의 차이.",
      "- ICH Q5A(R2): 유약 마우스, 성숙 마우스, 발육란 — 위해성 평가에 근거. CHO·NS0·SP2/0처럼 광범위하게 사용되고 특성이 잘 분석된 세포주는 in vivo 불필요.",
      "- FDA 2010: 위 시험 + 기니피그.",
      "- Ph. Eur.: 2017년 5.2.3·2.6.16 개정으로 일부 동물시험 삭제 (⚠ 삭제 항목 원문 확인).",
      "⚠ 동물 수·관찰 일수는 원문 대조 전 게재 금지 → 확인 후 Supplementary 표로.",
    ].join("\n"),
  });
  const Y = () => ({ text: "●", options: { color: NAVY, align: "center", valign: "middle", fontSize: 16 } });
  const rows = [
    [H("시험"), H("ICH Q5A(R2)"), H("FDA 2010 (백신)")],
    [C("유약 마우스", { bold: true }), Y(), Y()],
    [C("성숙 마우스", { bold: true }), Y(), Y()],
    [C("발육란", { bold: true }), Y(), Y()],
    [C("기니피그", { bold: true }), NA(), HOT("●")],
    [C("면제 세포주", { bold: true }), SOFT("CHO · NS0 · SP2/0"), NA()],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 7.6, colW: [2.4, 2.6, 2.6], rowH: 0.68, fontFace: FONT, fontSize: 15, border: TB });
  kwStack(s, [
    { k: "마우스 · 발육란", d: "ICH · FDA 공통" },
    { k: "기니피그", d: "FDA 2010 백신 세포기질", hot: true },
    { k: "위해성 평가 기반", d: "특성 분석된 세포주 면제" },
  ], 8.9, 1.95, 3.8, 1.3);
}

// 2-4 항체생산시험 · PCR
{
  const s = base({
    section: "02 부정시험", title: "항체생산시험 · PCR",
    cite: "ICH Q5A(R2) Table 3 · §3.2.5",
    notes: [
      "요지: 표적 기반 시험의 검출 범위는 표적 목록 그 자체.",
      "- 설치류 세포주이거나 설치류 유래 원료에 노출된 경우: MAP(마우스)·RAP(랫드)·HAP(햄스터). 대상 바이러스는 Q5A(R2) Table 3.",
      "- Q5A(R2): Table 3의 동물 시험은 PCR 등 NAT 또는 NGS로 정면 비교 없이 대체 가능 (⚠ 문구 원문 대조).",
      "- 종특이 PCR은 알려진 표적에 대해 빠르고 민감. 목록 밖은 원리적으로 검출하지 않는다.",
      "- Supplementary의 Bas-Congo 슬라이드: 질의응답에서 'PCR 패널을 넓히면 되지 않나' 질문 대비.",
    ].join("\n"),
  });
  txt(s, [
    { text: "표적 목록 = ", options: {} },
    { text: "검출 범위", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 2.1, w: 12.1, h: 1.3, fontSize: 48, bold: true, valign: "middle" });
  kwTriple(s, [
    { label: "설치류 세포 · 원료 노출 시", value: "MAP · HAP · RAP", num: true },
    { label: "대상 바이러스 목록", value: "Q5A(R2) Table 3", num: true },
    { label: "알려진 표적 전용", value: "종특이 PCR" },
  ], 0.6, 4.35, 12.1);
}

// 2-5 내인성 시험
{
  const s = base({
    section: "02 부정시험", title: "내인성: 정량의 대상",
    cite: "ICH Q5A(R2) Table 1 (Tests for retroviruses and other endogenous viruses) · FDA 2010 Cell Substrates Guidance",
    notes: [
      "요지: 내인성 바이러스는 '없음'을 증명하는 대상이 아니라 '얼마나 있는가'를 재고 공정으로 관리하는 대상.",
      "- 시험: 역전사효소 활성(PERT/PBRT), 투과전자현미경(TEM), 감염성 시험.",
      "- CHO의 RVLP: 존재를 전제로 TEM 정량 → 공정의 제거·불활화 능력으로 안전역 확보.",
      "- 1-7(외래성/내인성 정의)이 시험 설계에서 어떻게 달라지는지 보여주는 장.",
    ].join("\n"),
  });
  txt(s, [
    { text: "존재를", options: { breakLine: true } },
    { text: "전제로 한다", options: {} },
  ], { x: 0.6, y: 2.0, w: 5.0, h: 2.0, fontSize: 40, bold: true });
  txt(s, "예  CHO 레트로바이러스 유사입자 → 정량 + 공정 제거 · 불활화", { x: 0.6, y: 4.4, w: 5.2, h: 0.9, fontSize: 13, color: GRAY });
  const items = [
    { k: "PERT", d: "역전사효소 활성" },
    { k: "TEM", d: "입자 형태 · 정량", hot: true },
    { k: "감염성 시험", d: "복제능 확인" },
  ];
  items.forEach((it, i) => {
    const y = 1.85 + i * 1.5;
    rect(s, 6.4, y, 6.3, 1.25, it.hot ? ORANGE : LIGHT);
    txt(s, it.k, { x: 6.7, y, w: 3, h: 1.25, fontSize: 26, bold: true, fontFace: it.k === "감염성 시험" ? FONT : NUMF, valign: "middle" });
    txt(s, it.d, { x: 9.6, y, w: 3.0, h: 1.25, fontSize: 15, valign: "middle" });
  });
}

// 2-6 세 가지 한계
{
  const s = base({
    section: "02 부정시험", title: "전통법의 세 가지 한계",
    cite: "ICH Q5A(R2) §3.2.5.2 (in vitro·in vivo의 검출 폭 한계 언급) · 사례는 파트 4 참조",
    notes: [
      "요지(발표자 해석): 전통법의 한계는 세 유형으로 정리된다. 각 유형이 파트 4의 사례와 연결된다.",
      "- 표적 기반: 목록에 없거나 프라이머가 닿지 않으면 음성 → FBS·trypsin, Rotarix.",
      "- 증식 기반: 지시세포에서 자라지 않거나 잠복 상태면 음성 → Vero·High Five 전사체.",
      "- 검출 폭: 동물이 반응하는 바이러스군이 제한적 → in vivo 9종 중 5종.",
      "- 균형 문장: 이 시험들은 수십 년간 실제 오염을 잡아냈다(Kerr & Nims 2010). 이 발표는 폐기론이 아니라 보완·대체의 근거를 다룬다.",
    ].join("\n"),
  });
  const rows = [
    [H("한계"), H("원리"), H("관련 사례")],
    [C("표적 기반", { bold: true, fontSize: 18 }), C("목록 · 프라이머 밖은 음성"), C("FBS · trypsin / Rotarix")],
    [C("증식 기반", { bold: true, fontSize: 18 }), C("비증식 · 잠복 상태는 음성"), C("Vero · High Five 전사체")],
    [{ text: "검출 폭", options: { bold: true, fontSize: 18, color: NAVY, fill: { color: ORANGE }, valign: "middle" } }, C("동물 반응 바이러스군 제한"), C("In vivo 9종 중 5종")],
  ];
  s.addTable(rows, { x: 0.6, y: 1.85, w: 12.1, colW: [2.8, 4.9, 4.4], rowH: 0.95, fontFace: FONT, fontSize: 15, border: TB });
  txt(s, "전통법의 실제 검출 기록도 함께 다룬다 — 10년간 bulk harvest 시험 (파트 4)", { x: 0.6, y: 5.95, w: 12.1, h: 0.4, fontSize: 12.5, color: GRAY });
}

// =====================================================================
// 03 NGS
divider("03", "NGS 원리와 분석", ["Chemistry", "전처리", "In-silico"], "11분",
  "파트 3 (11분). 화학 원리는 3장·약 3분만 가볍게. 나머지는 전처리 설계와 in-silico 분석, 특히 host removal.");

// 3-1 질문이 다르다
{
  const s = base({
    section: "03 NGS", title: "질문이 다르다",
    cite: "Chiu CY, Miller SA. Nat Rev Genet 2019;20:341 · Ph. Eur. 2.6.41 (EDQM 2025)",
    notes: [
      "요지: PCR은 '이것이 있는가'를, non-targeted NGS는 '무엇이 있는가'를 묻는다.",
      "- 흐름: 검체 → 전처리(뉴클레이스·추출) → 라이브러리 → 시퀀싱 → in-silico(QC, host removal, 분류, assembly, 판정) → 확증시험.",
      "- Ph. Eur. 2.6.41은 전처리부터 보고·후속조치까지 전 과정을 다룬다.",
    ].join("\n"),
  });
  txt(s, "PCR", { x: 0.6, y: 1.8, w: 2.0, h: 0.9, fontSize: 30, bold: true, color: MUTED, fontFace: NUMF, valign: "middle" });
  txt(s, "이것이 있는가?", { x: 2.6, y: 1.8, w: 9, h: 0.9, fontSize: 34, bold: true, color: MUTED, valign: "middle" });
  txt(s, "NGS", { x: 0.6, y: 2.85, w: 2.0, h: 0.9, fontSize: 30, bold: true, fontFace: NUMF, valign: "middle" });
  txt(s, [{ text: "무엇이", options: { highlight: ORANGE } }, { text: " 있는가?" }], { x: 2.6, y: 2.85, w: 9, h: 0.9, fontSize: 34, bold: true, valign: "middle" });
  const flow = ["검체", "전처리", "라이브러리", "시퀀싱", "In-silico", "확증시험"];
  const fw = 1.8, fg = 0.26, fy = 4.95;
  flow.forEach((f, i) => {
    const x = 0.6 + i * (fw + fg);
    const hot = f === "In-silico";
    rect(s, x, fy, fw, 0.75, hot ? ORANGE : LIGHT);
    txt(s, f, { x, y: fy, w: fw, h: 0.75, fontSize: 15, bold: true, align: "center", valign: "middle" });
    if (i < flow.length - 1) arrow(s, x + fw + 0.02, fy + 0.375, x + fw + fg - 0.02, fy + 0.375, MUTED);
  });
  // 구간 라벨 (Wet / Dry / 확증)
  const seg = (x1, x2, label) => {
    hline(s, x1, x2, 4.62, NAVY, 1);
    vline(s, x1, 4.62, 4.8, NAVY, 1);
    vline(s, x2, 4.62, 4.8, NAVY, 1);
    txt(s, label, { x: x1, y: 4.15, w: x2 - x1, h: 0.4, fontSize: 14, bold: true, align: "center" });
  };
  const xs = (i) => 0.6 + i * (fw + fg);
  seg(xs(0), xs(3) + fw, "Wet lab");
  seg(xs(4), xs(4) + fw, "Dry lab");
  seg(xs(5), xs(5) + fw, "확증");
}

// 3-2 라이브러리 제작
{
  const s = base({
    section: "03 NGS", title: "라이브러리 제작",
    cite: "Illumina. An Introduction to Next-Generation Sequencing Technology · Illumina. DNA Library Preparation · Ng SHS et al. Viruses 2018;10:566",
    notes: [
      "요지: 핵산을 짧게 자르고, 양 끝에 어댑터를 붙이고, 검체마다 index(바코드)를 달아 한 번에 분석한다.",
      "- 단편화 + 말단 수선·어댑터 결합(ligation) 또는 tagmentation(트랜스포좀이 단편화와 어댑터 부착을 한 번에).",
      "- RNA 바이러스·전사체: random priming으로 cDNA 합성 후 동일 과정.",
      "- Index 멀티플렉싱 → 뒤의 '대조군 설계'에서 index 오배정(교차오염) 문제로 이어짐. 도식에서 오렌지 블록이 index.",
    ].join("\n"),
  });
  // 도식: 3 단계
  const y0 = 2.2;
  // 1 단편화
  rect(s, 0.6, y0 + 0.2, 3.3, 0.22, NAVY);
  arrow(s, 2.25, y0 + 0.6, 2.25, y0 + 1.15, MUTED);
  [0.6, 1.5, 2.35, 3.15].forEach((x, i) => rect(s, x, y0 + 1.4 + (i % 2) * 0.45, 0.7, 0.22, NAVY));
  // 2 어댑터
  arrow(s, 4.1, y0 + 1.3, 4.7, y0 + 1.3, MUTED);
  [0, 1, 2].forEach((i) => {
    const yy = y0 + 0.75 + i * 0.5;
    rect(s, 4.95, yy, 0.3, 0.22, "A6A6A6");
    rect(s, 5.25, yy, 1.5, 0.22, NAVY);
    rect(s, 6.75, yy, 0.3, 0.22, "A6A6A6");
  });
  // 3 index · pool
  arrow(s, 7.25, y0 + 1.3, 7.85, y0 + 1.3, MUTED);
  [0, 1, 2, 3, 4].forEach((i) => {
    const yy = y0 + 0.45 + i * 0.4;
    rect(s, 8.1 + (i % 2) * 0.35, yy, 0.22, 0.22, ORANGE);
    rect(s, 8.32 + (i % 2) * 0.35, yy, 0.25, 0.22, "A6A6A6");
    rect(s, 8.57 + (i % 2) * 0.35, yy, 1.5, 0.22, NAVY);
    rect(s, 10.07 + (i % 2) * 0.35, yy, 0.25, 0.22, "A6A6A6");
  });
  txt(s, "■ 삽입 서열   ■ 어댑터   ■ Index", { x: 8.1, y: y0 + 2.6, w: 4.6, h: 0.3, fontSize: 10, color: MUTED });
  // 키워드 3단
  const kws = [
    { k: "단편화", d: "DNA · cDNA (RNA는 random priming)" },
    { k: "어댑터 · Tagmentation", d: "결합 또는 트랜스포좀 일괄 부착" },
    { k: "Index 멀티플렉싱", d: "검체별 바코드 → 한 run에 다중 검체", hot: true },
  ];
  kws.forEach((k, i) => {
    const x = [0.6, 4.6, 8.1][i];
    txt(s, k.k, { x, y: 5.3, w: 3.9, h: 0.45, fontSize: 18, bold: true, color: NAVY });
    txt(s, k.d, { x, y: 5.8, w: 3.9, h: 0.5, fontSize: 12, color: GRAY });
  });
}

// 3-3 클러스터 생성
{
  const s = base({
    section: "03 NGS", title: "클러스터 생성",
    cite: "Illumina. Patterned Flow Cell Technology · Illumina. An Introduction to NGS Technology · Bentley DR et al. Nature 2008;456:53",
    notes: [
      "요지: 한 분자를 제자리에서 수천 복제해 '보이는' 신호로 만든다.",
      "- Bridge amplification: flow cell 표면 올리고에 결합한 단일 가닥이 제자리에서 증폭 → 클론 클러스터.",
      "- Patterned flow cell: 위치가 고정된 나노웰 배열. ExAmp(Exclusion Amplification)는 시딩과 증폭을 동시에 진행해 한 웰에 두 분자가 자랄 확률을 낮춘다.",
      "- 도식은 개념도(축척 무관). Illumina 공식 도식으로 교체 가능.",
    ].join("\n"),
  });
  // 왼쪽: bridge amplification 3단계 (표면 + 가닥 수)
  const sx = 0.8, sy = 4.3;
  [1, 2, 8].forEach((n, i) => {
    const x = sx + i * 1.9;
    hline(s, x, x + 1.5, sy, NAVY, 2.5);
    for (let j = 0; j < n; j++) {
      const lx = n === 1 ? x + 0.75 : x + 0.15 + (j * 1.2) / (n - 1);
      vline(s, lx, sy - 1.1, sy, j === 0 && i === 2 ? ORANGE : NAVY, 2);
    }
    txt(s, ["결합", "브리지", "클러스터"][i], { x, y: sy + 0.15, w: 1.5, h: 0.35, fontSize: 12, color: GRAY, align: "center" });
  });
  arrow(s, sx + 1.55, sy - 0.55, sx + 1.85, sy - 0.55, MUTED);
  arrow(s, sx + 3.45, sy - 0.55, sx + 3.75, sy - 0.55, MUTED);
  txt(s, "Bridge amplification", { x: 0.8, y: 2.0, w: 5.5, h: 0.5, fontSize: 20, bold: true });
  // 오른쪽: patterned flow cell 나노웰 격자
  txt(s, "Patterned flow cell", { x: 7.3, y: 2.0, w: 5.4, h: 0.5, fontSize: 20, bold: true });
  const gx = 7.4, gy = 2.8;
  for (let r = 0; r < 5; r++) {
    for (let c = 0; c < 9; c++) {
      const empty = (r * 9 + c) % 7 === 3;
      const hot = r === 2 && c === 4;
      const d = 0.34;
      const cx = gx + c * 0.55 + d / 2, cy = gy + r * 0.45 + d / 2;
      if (empty) s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: WHITE }, line: { color: "A6A6A6", width: 1 } });
      else dot(s, cx, cy, d, hot ? ORANGE : NAVY);
    }
  }
  txt(s, "고정 위치 나노웰 · 웰당 한 클론", { x: 7.3, y: 5.15, w: 5.4, h: 0.35, fontSize: 12, color: GRAY });
  kwChips(s, ["Bridge amplification", "Patterned flow cell", { k: "ExAmp", hot: true }], 0.6, 5.95, 12.1, 0.55);
}

// 3-4 SBS
{
  const s = base({
    section: "03 NGS", title: "SBS: 사이클당 한 염기",
    cite: "Bentley DR et al. Nature 2008;456:53 · Illumina. Sequencing Technology (SBS) · Illumina. 2-Channel SBS · Hirai T et al. Biologicals 2024;85:101739",
    notes: [
      "요지: 형광 표지된 가역적 종결자(reversible terminator) dNTP로 한 사이클에 한 염기씩 합성·촬영·제거를 반복한다.",
      "- 2-channel(NextSeq·NovaSeq 등): 적·녹 두 색으로 4염기 판독. A=적+녹, C=적, T=녹, G=신호 없음(dark G).",
      "- Paired-end: 한 단편의 양 끝을 읽어 정렬 정확도 향상.",
      "- 곁가지: long-read(Oxford Nanopore)도 동일 목적에 사용 가능. Hirai 2024에서 short-read와 검출한계가 거의 동일 (파트 4).",
    ].join("\n"),
  });
  // 사이클 3단계 (가로)
  const cyc = [{ k: "결합", d: "종결자 dNTP 1개" }, { k: "촬영", d: "형광 판독" }, { k: "절단", d: "형광 · 종결기 제거" }];
  cyc.forEach((c, i) => {
    const x = 0.6 + i * 2.2;
    s.addShape(pres.shapes.OVAL, { x, y: 2.0, w: 1.75, h: 1.75, fill: { color: i === 0 ? ORANGE : LIGHT }, line: { color: i === 0 ? ORANGE : LIGHT, width: 0 } });
    txt(s, c.k, { x, y: 2.35, w: 1.75, h: 0.5, fontSize: 20, bold: true, align: "center" });
    txt(s, c.d, { x: x + 0.1, y: 2.85, w: 1.55, h: 0.6, fontSize: 10.5, align: "center", color: GRAY });
    if (i < 2) arrow(s, x + 1.8, 2.875, x + 2.15, 2.875, MUTED);
  });
  txt(s, "× 읽기 길이만큼 반복", { x: 0.6, y: 3.95, w: 6.4, h: 0.4, fontSize: 13, color: GRAY });
  // 2-channel 표
  const on = () => ({ text: "●", options: { color: NAVY, align: "center", valign: "middle", fontSize: 18 } });
  const off = () => ({ text: "○", options: { color: "A6A6A6", align: "center", valign: "middle", fontSize: 18 } });
  const B = (t, hot) => ({ text: t, options: { bold: true, fontFace: NUMF, fontSize: 20, align: "center", valign: "middle", color: NAVY, fill: hot ? { color: ORANGE } : undefined } });
  const rows = [
    [H("염기"), H("적색"), H("녹색")],
    [B("A"), on(), on()],
    [B("C"), on(), off()],
    [B("T"), off(), on()],
    [B("G", true), off(), off()],
  ];
  txt(s, "2-channel 판독", { x: 7.6, y: 1.75, w: 5, h: 0.4, fontSize: 14, bold: true });
  s.addTable(rows, { x: 7.6, y: 2.2, w: 5.1, colW: [1.7, 1.7, 1.7], rowH: 0.52, fontFace: FONT, fontSize: 14, border: TB });
  txt(s, "G = 신호 없음 (dark G)", { x: 7.6, y: 4.95, w: 5.1, h: 0.3, fontSize: 11, color: MUTED });
  kwChips(s, [{ k: "가역적 종결자", hot: true }, "2-channel 판독", "Paired-end"], 0.6, 5.95, 12.1, 0.55);
}

// 3-5 무엇을 시퀀싱할 것인가
{
  const s = base({
    section: "03 NGS", title: "무엇을 시퀀싱할 것인가",
    cite: "Allander T et al. PNAS 2001;98:11609 · Ng SHS et al. Viruses 2018;10:566 · Onions D et al. Vaccine 2011;29:7117",
    notes: [
      "요지: 전처리 선택이 '무엇을 볼 수 있는가'를 결정한다.",
      "- 뉴클레이스 처리·여과: 캡시드 밖 숙주 핵산을 분해해 바이러스 입자 핵산을 농축 (Allander 2001). 대가: 입자를 형성하지 않은 잠복·통합 바이러스는 놓친다.",
      "- 총 핵산·전사체: 잠복·발현 바이러스까지 포착 (Onions 2011, 파트 4). 대가: 숙주 read가 압도적 → host removal 부담.",
      "- DNA/RNA 병행, rRNA 제거로 숙주 배경 감소, 저투입량 증폭 (Ng 2018).",
    ].join("\n"),
  });
  const cols = [
    { x: 0.6, k: "뉴클레이스 처리 · 여과", sub: "입자 내 핵산 농축", plus: "숙주 배경 감소", minus: "잠복 · 통합 바이러스 누락" },
    { x: 6.85, k: "총 핵산 · 전사체", sub: "세포 전체 분석", plus: "잠복 · 발현 바이러스 포착", minus: "숙주 read 과다", hot: true },
  ];
  cols.forEach((c) => {
    rect(s, c.x, 1.75, 5.9, 3.55, c.hot ? WHITE : LIGHT, c.hot ? ORANGE : null);
    txt(s, c.k, { x: c.x + 0.35, y: 1.95, w: 5.3, h: 0.5, fontSize: 22, bold: true });
    txt(s, c.sub, { x: c.x + 0.35, y: 2.5, w: 5.3, h: 0.4, fontSize: 13, color: GRAY });
    txt(s, "+", { x: c.x + 0.35, y: 3.2, w: 0.4, h: 0.5, fontSize: 26, bold: true, fontFace: NUMF });
    txt(s, c.plus, { x: c.x + 0.85, y: 3.25, w: 4.8, h: 0.5, fontSize: 17 });
    txt(s, "−", { x: c.x + 0.35, y: 4.05, w: 0.4, h: 0.5, fontSize: 26, bold: true, fontFace: NUMF, color: MUTED });
    txt(s, c.minus, { x: c.x + 0.85, y: 4.1, w: 4.8, h: 0.5, fontSize: 17, color: GRAY });
  });
  txt(s, [
    { text: "공통  ", options: { color: MUTED, bold: true } },
    { text: "DNA · RNA 병행   ·   rRNA 제거   ·   저투입량 증폭", options: { bold: true } },
  ], { x: 0.6, y: 5.65, w: 12.1, h: 0.5, fontSize: 16, valign: "middle" });
}

// 3-6 대조군 설계
{
  const s = base({
    section: "03 NGS", title: "대조군 설계",
    cite: "Mee ET et al. Vaccine 2016;34:2035 · WHO/BS/2024.2471 (1st International Reference Panel) · Illumina white paper 770-2017-004 (index misassignment) · Salter SJ et al. BMC Biol 2014;12:87",
    notes: [
      "요지: NGS 결과의 신뢰도는 대조군 설계가 정한다.",
      "- 6/25: NIBSC 후보 표준물질(25종 바이러스)을 16개 기관이 분석 → 전 기관 검출은 6종, 나머지는 4–14개 기관 (Mee 2016). 기관 간 편차가 크다.",
      "- 7종: WHO 1st International Reference Panel (2024) — EBV, reovirus 1, RSV, FeLV, PCV1, HCoV-OC43, MVM.",
      "- ~2%: patterned flow cell에서 index 오배정 최대 약 2% (Illumina). 강한 양성 대조가 같은 run의 시험 검체로 번질 수 있다 → Unique dual index(UDI)로 차단.",
      "- 음성 대조(추출 blank, no-template): 시약 유래 핵산 'kitome' (Salter 2014).",
    ].join("\n"),
  });
  const st = [
    { n: "6 / 25", k: "기관 간 편차", d: "16개 기관 전원이 검출한 바이러스 수\nNIBSC 후보 표준물질", hot: true },
    { n: "7종", k: "국제 표준 패널", d: "WHO 1st International\nReference Panel (2024)" },
    { n: "~2%", k: "Index 오배정", d: "Patterned flow cell 최대치\n→ Unique dual index" },
  ];
  st.forEach((x, i) => {
    const xx = 0.6 + i * 4.15;
    txt(s, x.k, { x: xx, y: 1.85, w: 3.8, h: 0.45, fontSize: 17, bold: true, color: NAVY });
    if (x.hot) rect(s, xx - 0.05, 2.45, 2.95, 1.25, ORANGE);
    txt(s, x.n, { x: xx, y: 2.4, w: 3.8, h: 1.35, fontSize: 60, bold: true, fontFace: NUMF, valign: "middle" });
    txt(s, x.d, { x: xx, y: 3.95, w: 3.8, h: 0.8, fontSize: 12.5, color: GRAY });
  });
  txt(s, "음성 대조 필수 — 추출 blank · no-template (시약 유래 핵산, kitome)", { x: 0.6, y: 5.6, w: 12.1, h: 0.45, fontSize: 14, color: NAVY });
}

// 3-7 In-silico 파이프라인
{
  const s = base({
    section: "03 NGS", title: "In-silico 파이프라인",
    cite: "Bolger 2014 (Trimmomatic) · Chen 2018 (fastp) · Li & Durbin 2009 (BWA) · Langmead 2012 (Bowtie2) · Buchfink 2015 (DIAMOND) · Wood 2019 (Kraken2) · Bankevich 2012 (SPAdes) · Li 2015 (MEGAHIT) · Simonyan 2016 (HIVE)",
    notes: [
      "요지: NGS의 성능은 시퀀서가 아니라 이 파이프라인과 그 밸리데이션이 결정한다.",
      "① QC·트리밍: 어댑터·저품질·저복잡도 read 제거 (FastQC, Trimmomatic, fastp).",
      "② Host removal: 생산세포주 유전체에 정렬해 제거 (BWA, Bowtie2) — 다음 두 장.",
      "③ 분류: BLASTn/x, DIAMOND, Kraken2 — DB 선택이 핵심.",
      "④ De novo assembly: 신종·발산 바이러스 contig 복원 (SPAdes, MEGAHIT).",
      "⑤ 판정·확증: 역치 판정 → 특이 PCR, 감염성 시험.",
      "- 규제기관 플랫폼 예: FDA CBER HIVE (Simonyan 2016).",
    ].join("\n"),
  });
  const steps = [
    { k: "QC · 트리밍", t: "FastQC\nTrimmomatic · fastp" },
    { k: "Host removal", t: "BWA · Bowtie2", hot: true },
    { k: "분류", t: "BLAST · DIAMOND\nKraken2" },
    { k: "Assembly", t: "SPAdes · MEGAHIT" },
    { k: "판정 · 확증", t: "역치 · PCR\n감염성 시험" },
  ];
  const bw = 2.15, gp = 0.34;
  steps.forEach((st, i) => {
    const x = 0.6 + i * (bw + gp);
    txt(s, String(i + 1), { x, y: 1.8, w: bw, h: 0.5, fontSize: 22, bold: true, fontFace: NUMF, color: st.hot ? ORANGE : MUTED });
    rect(s, x, 2.35, bw, 1.25, st.hot ? ORANGE : NAVY);
    txt(s, st.k, { x, y: 2.35, w: bw, h: 1.25, fontSize: 17, bold: true, color: st.hot ? NAVY : WHITE, align: "center", valign: "middle" });
    txt(s, st.t, { x, y: 3.8, w: bw, h: 0.8, fontSize: 11.5, color: GRAY, align: "center", fontFace: NUMF });
    if (i < steps.length - 1) arrow(s, x + bw + 0.03, 2.975, x + bw + gp - 0.03, 2.975, MUTED);
  });
  kwTriple(s, [
    { label: "생산세포주 유전체 기준", value: "Host removal" },
    { label: "범위와 오분류의 균형", value: "DB 선택" },
    { label: "서열 신호의 검증", value: "직교 확증" },
  ], 0.6, 5.1, 12.1);
}

// 3-8 호스트 = 생산세포주
{
  const s = base({
    section: "03 NGS", title: "호스트 = 생산세포주",
    cite: "Xu X et al. Nat Biotechnol 2011;29:735 · Lewis NE et al. Nat Biotechnol 2013;31:759 · Rupp O et al. Biotechnol Bioeng 2018;115:2087 · Osada N et al. DNA Res 2014;21:673 · Lin YC et al. Nat Commun 2014;5:4767 · Nandakumar S et al. Genome Announc 2017;5:e00829-17",
    notes: [
      "요지: host removal의 '호스트'는 환자가 아니라 생산세포주. 참조 유전체의 품질이 위양성률을 좌우한다.",
      "- CHO: Xu 2011(CHO-K1), Lewis 2013(C. griseus), Rupp 2018 PICR(PacBio+Illumina 하이브리드). PICR 저자들은 이전 어셈블리가 단편적이고 gap·오조립이 많다고 명시.",
      "- Vero: Osada 2014 (Chlorocebus sabaeus 유래). SRV 근연 베타레트로바이러스 발현 보고 (Onions 2011).",
      "- HEK293: Lin 2014, 계통(lineage)별 게놈 차이, pseudotriploid.",
      "- Sf9: Nandakumar 2017 (FDA CBER, PacBio). Sf 세포는 rhabdovirus 유사 EVE 보유 (Geisler 2016).",
      "- 쟁점: 불완전 참조 / 실제 생산 계대주와의 계통 차이 / ERV·EVE는 숙주 서열이면서 바이러스 서열.",
    ].join("\n"),
  });
  const I = (t) => ({ text: t, options: { italic: true, color: NAVY, valign: "middle", fontFace: NUMF } });
  const rows = [
    [H("세포주"), H("종"), H("참조 유전체"), H("주의점")],
    [C("CHO", { bold: true }), I("Cricetulus griseus"), C("Xu 2011 · Lewis 2013 · PICR 2018"), C("이수성 · 재배열")],
    [C("Vero", { bold: true }), I("Chlorocebus sabaeus"), C("Osada 2014"), C("베타레트로바이러스 발현")],
    [C("HEK293", { bold: true }), I("Homo sapiens"), C("Lin 2014"), C("계통별 게놈 차이")],
    [C("Sf9", { bold: true }), I("Spodoptera frugiperda"), C("Nandakumar 2017"), { text: "Rhabdovirus 유사 EVE", options: { fill: { color: ORANGE }, color: NAVY, bold: true, valign: "middle" } }],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 12.1, colW: [1.6, 2.9, 4.0, 3.6], rowH: 0.62, fontFace: FONT, fontSize: 13.5, border: TB });
  kwChips(s, ["불완전 참조", "계통 차이", { k: "ERV · EVE", hot: true }], 0.6, 5.55, 12.1, 0.6);
}

// 3-9 지우면 안 되는 것
{
  const s = base({
    section: "03 NGS", title: "지우면 안 되는 것",
    cite: "Geisler C, Jarvis DL. Biologicals 2016;44:219 · Goodacre N et al. mSphere 2018;3:e00069-18 · ICH Q5A(R2)",
    notes: [
      "요지: ERV·EVE는 숙주 서열이면서 바이러스 서열. 기계적으로 지우면 안 되고, 식별해서 해석해야 한다.",
      "- 잔류 숙주 read: 참조 유전체가 불완전하거나 계통이 다르면 숙주 read가 남아 바이러스 DB에 오정렬 → 위양성.",
      "- ERV·EVE: 판정 단계에서 '내인성 요소'로 해석 (1-7의 정의로 귀결).",
      "- DNA+RNA 병행: 유전체와 전사체를 함께 보면 전사되는 EVE를 식별할 수 있다 (Geisler 2016).",
      "⚠ '과도한 제거가 숙주 유사 서열을 가진 바이러스를 함께 지운다'는 일반 원리 수준. 특정 문헌 미확인 → 발표 시 '원리적 위험'으로만 언급.",
    ].join("\n"),
  });
  txt(s, [
    { text: "숙주 서열이면서 ", options: {} },
    { text: "바이러스 서열", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 1.7, w: 12.1, h: 0.8, fontSize: 32, bold: true, valign: "middle" });
  // 분기 도식
  rect(s, 0.6, 3.55, 2.6, 1.0, NAVY);
  txt(s, "Host 정렬", { x: 0.6, y: 3.55, w: 2.6, h: 1.0, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle" });
  const br = [
    { y: 2.75, k: "잔류 숙주 read", d: "DB 오정렬 → 위양성" },
    { y: 3.9, k: "ERV · EVE", d: "제거가 아닌 해석 대상", hot: true },
    { y: 5.05, k: "DNA + RNA 병행", d: "전사되는 EVE 식별" },
  ];
  br.forEach((b) => {
    arrow(s, 3.25, 4.05, 4.35, b.y + 0.35, MUTED);
    rect(s, 4.45, b.y, 3.6, 0.72, b.hot ? ORANGE : LIGHT);
    txt(s, b.k, { x: 4.45, y: b.y, w: 3.6, h: 0.72, fontSize: 17, bold: true, align: "center", valign: "middle" });
    txt(s, b.d, { x: 8.35, y: b.y, w: 4.4, h: 0.72, fontSize: 14, color: GRAY, valign: "middle" });
  });
}

// 3-10 DB 선택
{
  const s = base({
    section: "03 NGS", title: "DB 선택",
    cite: "Goodacre N et al. mSphere 2018;3:e00069-18 · Chin P et al. mSphere 2025;10:e00286-25 · MacDonald ML et al. mSphere 2021;6:e01336-20 · Buchfink B et al. Nat Methods 2015;12:59 · Wood DE et al. Genome Biol 2019;20:257",
    notes: [
      "요지: 같은 read라도 DB에 따라 결과가 달라진다.",
      "- 좁은 DB(RefSeq viral): 발산한 바이러스를 놓침.",
      "- 넓은 DB(GenBank nt 전체): 세포 서열이 다수 → 느리고 오분류 증가.",
      "- RVDB (FDA CBER): 바이러스·바이러스 관련 서열 망라, 파지 제외, 세포 서열 축소, 98% 동일성 군집(CD-HIT-EST). 2025년 개정(U-RVDB/C-RVDB).",
      "- 분류기: 정렬 기반(BLAST, DIAMOND) vs k-mer 기반(Kraken2) — 속도와 민감도의 교환. 생물의약품 적용 벤치마크: MacDonald 2021.",
    ].join("\n"),
  });
  const rows = [
    [H("DB"), H("범위"), H("위험")],
    [C("RefSeq viral", { bold: true, fontFace: NUMF }), C("좁음"), C("발산 바이러스 누락")],
    [C("GenBank nt", { bold: true, fontFace: NUMF }), C("넓음 · 세포 서열 포함"), C("속도 저하 · 오분류")],
    [{ text: "RVDB (FDA CBER)", options: { bold: true, fill: { color: ORANGE }, color: NAVY, valign: "middle" } }, C("바이러스 전용 · 파지 제외 · 98% 군집"), C("세포 서열 축소")],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 8.3, colW: [2.4, 3.3, 2.6], rowH: 0.85, fontFace: FONT, fontSize: 14, border: TB });
  kwStack(s, [
    { k: "정렬 vs k-mer", d: "BLAST · DIAMOND / Kraken2" },
    { k: "DB의 폭", d: "누락과 오분류의 교환" },
    { k: "RVDB", d: "규제기관 개발 · 2025 개정", hot: true },
  ], 9.5, 1.95, 3.2, 1.25);
}

// 3-11 판정과 확증
{
  const s = base({
    section: "03 NGS", title: "판정과 확증",
    cite: "Ph. Eur. 2.6.41 (EDQM 2025) · ICH Q5A(R2) · Khan AS et al. mSphere 2017;2:e00307-17",
    notes: [
      "요지: NGS 양성 신호는 '추정 검출'. 직교 시험으로 확증해야 판정이 된다.",
      "- 판정 역치(read 수, 커버리지, 독립 영역 수, e-value 등)는 각 시험기관이 spike-in 밸리데이션으로 설정. 규제상 공통 수치는 확인되지 않음 → 특정 숫자를 표준처럼 제시하지 말 것.",
      "- 확증: 특이 PCR → 감염성 시험으로 복제 바이러스 / 불활화 핵산 / EVE·ERV 구분.",
      "- 다기관 근거: Khan 2017 — 3개 기관이 서로 다른 플랫폼·분석법으로 비교 가능한 결과 (⚠ 수치 원문 대조).",
    ].join("\n"),
  });
  const fl = [
    { k: "NGS 신호", d: "추정 검출", hot: true },
    { k: "특이 PCR", d: "서열 확인" },
    { k: "감염성 시험", d: "복제능 확인" },
  ];
  fl.forEach((f, i) => {
    const x = 0.6 + i * 2.75;
    rect(s, x, 2.05, 2.35, 1.3, f.hot ? ORANGE : LIGHT);
    txt(s, f.k, { x, y: 2.15, w: 2.35, h: 0.6, fontSize: 18, bold: true, align: "center", valign: "middle" });
    txt(s, f.d, { x, y: 2.72, w: 2.35, h: 0.45, fontSize: 12, align: "center", color: GRAY });
    if (i < 2) arrow(s, x + 2.4, 2.7, x + 2.7, 2.7, MUTED);
  });
  arrow(s, 8.1, 2.7, 8.6, 2.7, MUTED);
  const out = ["복제 바이러스", "불활화 핵산", "EVE · ERV"];
  out.forEach((o, i) => {
    const y = 1.85 + i * 0.6;
    rect(s, 8.75, y, 3.95, 0.5, WHITE, NAVY);
    txt(s, o, { x: 8.75, y, w: 3.95, h: 0.5, fontSize: 14, bold: true, align: "center", valign: "middle" });
  });
  kwTriple(s, [
    { label: "공통 수치 없음", value: "기관별 역치" },
    { label: "PCR · 감염성", value: "직교 확증" },
    { label: "Ph. Eur. 2.6.41 요구", value: "후속 조사" },
  ], 0.6, 4.55, 12.1);
}

// =====================================================================
// 04 사례
divider("04", "사례 13편", ["오염의 경로", "NGS의 한계", "규제 전환"], "13분",
  "파트 4 (13분). 13편을 7개 묶음으로. 슬라이드마다 숫자 하나 또는 문장 하나. 원 노트 번호가 아니라 저자·연도로 부른다.");

// 4-0 지도
{
  const s = base({
    section: "04 사례", title: "13편, 7개 묶음",
    cite: "문헌 전체 목록은 References 참조",
    notes: [
      "요지: 청중이 지금 어디를 듣는지 알 수 있도록 지도를 먼저 보여준다.",
      "순서: 원료 → 세포주(전사체, 음성 입증, Sf9) → 함정(EVE) → 완제품(Rotarix) → 현장(10년 기록, 저항성 세포주) → 성능(in vivo 비교, LoD) → 구현(LAIV 교체, 4차 합의).",
      "Sf9(Ma 2014)와 EVE(Geisler 2016)를 연속 배치: '같은 세포주, 같은 바이러스 계통, 정반대 결론'.",
    ].join("\n"),
  });
  const groups = [
    { g: "원료", p: ["Gagnieur 2014", "Zhang 2022"] },
    { g: "세포주", p: ["Onions 2011", "Shabram 2014", "Ma 2014"] },
    { g: "함정", p: ["Geisler 2016"], hot: true },
    { g: "완제품", p: ["Victoria 2010"] },
    { g: "현장", p: ["Kerr 2010", "Mascarenhas 2017"] },
    { g: "성능", p: ["Beurdeley 2023", "Hirai 2024"] },
    { g: "구현", p: ["Alston 2025", "Khan 2025"] },
  ];
  const cw = 12.1 / 7;
  // 상단 묶음 괄호: 경로 / 사건 / 전환
  const br = (i1, i2, label) => {
    const x1 = 0.6 + i1 * cw + 0.08, x2 = 0.6 + (i2 + 1) * cw - 0.08;
    hline(s, x1, x2, 2.25, NAVY, 1);
    vline(s, x1, 2.25, 2.42, NAVY, 1);
    vline(s, x2, 2.25, 2.42, NAVY, 1);
    txt(s, label, { x: x1, y: 1.75, w: x2 - x1, h: 0.42, fontSize: 16, bold: true, align: "center" });
  };
  br(0, 2, "오염의 경로 · 한계");
  br(3, 4, "사건과 대응");
  br(5, 6, "규제 전환");
  groups.forEach((g, i) => {
    const x = 0.6 + i * cw;
    rect(s, x + 0.08, 2.65, cw - 0.16, 0.7, g.hot ? ORANGE : NAVY);
    txt(s, g.g, { x: x + 0.08, y: 2.65, w: cw - 0.16, h: 0.7, fontSize: 17, bold: true, color: g.hot ? NAVY : WHITE, align: "center", valign: "middle" });
    txt(s, g.p.map((p, j) => ({ text: p, options: { breakLine: j < g.p.length - 1 } })), { x: x + 0.08, y: 3.55, w: cw - 0.16, h: 1.6, fontSize: 12, align: "center", color: NAVY, fontFace: NUMF, paraSpaceAfter: 8 });
  });
}

// 4-1 원료
{
  const s = base({
    section: "04 사례 · 원료", title: "원료가 경로다",
    cite: "Gagnieur L et al. Biologicals 2014;42:145 (doi:10.1016/j.biologicals.2014.02.002) · Zhang P et al. Zool Res 2022;43:756 (doi:10.24272/j.issn.2095-8137.2022.093)",
    notes: [
      "요지: 세포배양 원료(FBS·trypsin) 자체가 다양한 바이러스를 운반한다. Rotarix의 원인(돼지 trypsin)과 같은 경로.",
      "- Gagnieur 2014: FBS에서 BVDV 1–3형·BPV3 주종, 신규 후보종(bovine parvovirus, bovine pegivirus). Trypsin에서 PCV2 주종.",
      "- Zhang 2022: 8년 뒤 다른 지역 상용 제품 재확인. BPV3·bosavirus 고빈도·고농도. Bovine norovirus·BVDV-1은 신규 유전형 시사.",
      "⚠ 원문 대조 전 게재 금지: 바이러스 read 비율 배치 간 0.002–22.7%, FBS 유래 BVDV 서열과 통용 스크리닝 프라이머의 불일치. 확인되면 이 장의 핵심 숫자로 추가.",
      "메시지: 'BVDV 검사를 하고 있다'와 '우리 FBS에 BVDV가 없다'는 다른 문장.",
    ].join("\n"),
  });
  const I2 = (t) => ({ text: t, options: { color: NAVY, valign: "middle" } });
  const rows = [
    [H("원료"), H("주요 검출"), H("신규")],
    [C("FBS (2014)", { bold: true }), HOT("BVDV 1–3형 · BPV3"), I2("bovine parvovirus · pegivirus 후보종")],
    [C("Trypsin (2014)", { bold: true }), HOT("PCV2"), NA()],
    [C("FBS · trypsin (2022)", { bold: true }), I2("BPV3 · bosavirus 고빈도"), I2("bovine norovirus · BVDV-1 신규 유전형")],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 8.2, colW: [2.2, 2.9, 3.1], rowH: 0.85, fontFace: FONT, fontSize: 13.5, border: TB });
  kwChips(s, ["BVDV · BPV3", "PCV2", "신규 유전형"], 0.6, 5.55, 8.2, 0.6);
  paperCard(s, 9.25, 1.8, 3.45, 4.35, {
    journal: "Biologicals · 2014", title: "Unbiased analysis by high throughput sequencing of the viral diversity in fetal bovine serum and trypsin …",
    authors: "Gagnieur L … Eloit M", slot: "원문 첫 페이지 캡처\n(Zhang 2022는 발표자 노트 참조)",
  });
}

// 4-2 발현을 본다 (Onions)
{
  const s = base({
    section: "04 사례 · 세포주", title: "발현을 본다",
    cite: "Onions D et al. Vaccine 2011;29:7117 (doi:10.1016/j.vaccine.2011.05.071)",
    notes: [
      "요지: 유전체에 '있는가'와 지금 '발현되는가'는 다른 질문. 전사체 시퀀싱(MP-Seq)은 후자를 묻는다.",
      "- Vero ATCC CCL-81 세포은행: PCV 및 기타 외래성 인자 배제. 동시에 SRV 근연 베타레트로바이러스 전장이 발현될 수 있음을 확인.",
      "- Trichoplusia ni(High Five) 세포: 오염된 잠복 nodavirus 검출, 발현 중인 errantivirus 유전체 확인.",
      "- 잠복 바이러스는 통상 배양 조건의 증식 기반 시험에서 신호가 없을 수 있다 → 파트 2의 '증식 기반' 한계.",
    ].join("\n"),
  });
  paperCard(s, 0.6, 1.8, 3.45, 4.6, {
    journal: "Vaccine · 2011", title: "Ensuring the safety of vaccine cell substrates by massively parallel sequencing of the transcriptome",
    authors: "Onions D, Côté C, Love B, et al.",
  });
  const blk = [
    { cell: "Vero CCL-81", a: "PCV 배제", b: "SRV 근연 베타레트로바이러스 전장 발현" },
    { cell: "High Five", a: "잠복 nodavirus", b: "errantivirus 유전체 발현", hot: true },
  ];
  blk.forEach((b, i) => {
    const y = 1.85 + i * 1.95;
    txt(s, b.cell, { x: 4.6, y, w: 8, h: 0.5, fontSize: 22, bold: true, fontFace: NUMF });
    if (b.hot) rect(s, 4.6, y + 0.62, 2.95, 0.55, ORANGE);
    txt(s, b.a, { x: 4.65, y: y + 0.62, w: 3.2, h: 0.55, fontSize: 17, bold: true, valign: "middle" });
    txt(s, b.b, { x: 7.9, y: y + 0.62, w: 4.8, h: 0.55, fontSize: 15, color: GRAY, valign: "middle" });
    hline(s, 4.6, 12.7, y + 1.55, LINE, 0.75);
  });
  kwChips(s, ["전사체 시퀀싱", { k: "잠복 바이러스", hot: false }, "레트로바이러스 발현"], 4.6, 5.8, 8.1, 0.55);
}

// 4-3 없음의 증명 (Shabram)
{
  const s = base({
    section: "04 사례 · 세포주", title: "없음의 증명",
    cite: "Shabram P, Kolman JL. PDA J Pharm Sci Technol 2014;68:639 (doi:10.5731/pdajpst.2014.01027)",
    notes: [
      "요지: NGS는 오염을 찾는 도구이자, '없음'을 입증해 새 세포기질의 규제 문턱을 넘는 도구.",
      "- A549(종양원성 인간 폐암 세포주)를 백신 세포기질로 쓸 수 있는가: 2012년 9월 FDA VRBPAC 심의.",
      "- 전 세포 전사체를 시퀀싱해 큐레이션된 바이러스 서열 DB와 대조.",
      "- 결론: 'A549 cells pose no more risk than any other cell substrate for vaccine manufacture.'",
      "- 신규 모달리티 공정 개발 청중에게 가장 실무적인 각도.",
    ].join("\n"),
  });
  txt(s, "A549", { x: 0.6, y: 1.8, w: 7.8, h: 0.8, fontSize: 30, bold: true, fontFace: NUMF, color: MUTED });
  txt(s, [
    { text: "다른 세포기질 대비", options: { breakLine: true } },
    { text: "추가 위험 없음", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 2.6, w: 7.8, h: 1.9, fontSize: 40, bold: true });
  kwTriple(s, [
    { label: "세포기질", value: "종양원성 세포주", size: 18 },
    { label: "FDA 자문위원회", value: "VRBPAC 2012.09", size: 18 },
    { label: "분석", value: "전사체 vs 바이러스 DB", size: 18 },
  ], 0.6, 5.0, 8.0);
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "PDA J Pharm Sci Technol · 2014", title: "Evaluation of A549 as a new vaccine cell substrate: digging deeper with massively parallel sequencing",
    authors: "Shabram P, Kolman JL",
  });
}

// 4-4 통과 ≠ 무바이러스 (Ma 2014)
{
  const s = base({
    section: "04 사례 · 세포주", title: "통과 ≠ 무바이러스",
    cite: "Ma H, Galvin TA, Glasner DR, Shaheduzzaman S, Khan AS. J Virol 2014;88:6576 (doi:10.1128/JVI.00780-14)",
    notes: [
      "요지: 'Sf9에서 보고된 바이러스가 없다'와 'Sf9에 바이러스가 없다'는 다른 말.",
      "- FDA CBER(Khan) 그룹. Degenerate PCR + massively parallel sequencing으로 Mononegavirales 목의 신종 랩도바이러스 동정.",
      "- ORF: N·P·M·G·L + G–L 사이 미지 ORF(111 aa). L 유전자 일부가 Taastrup virus 및 식물 cytorhabdovirus와 근연.",
      "- TEM으로 랩도바이러스 형태 확인. 모세포주 Sf21에도 존재, 다른 곤충세포주에는 없음. 인체 세포주에서 진입·복제 근거 없음.",
      "⚠ '광범위 시험에서 바이러스 보고 없음'은 초록 명시 문장이 아님 → 발표자 서술로 표현.",
      "- 게놈 도식은 유전자 순서만 표시(길이 비율 무관).",
      "- 다음 장(EVE)과 연속: 같은 세포주·같은 바이러스 계통·정반대 결론.",
    ].join("\n"),
  });
  paperCard(s, 0.6, 1.8, 3.45, 4.6, {
    journal: "J Virol · 2014", title: "Identification of a novel rhabdovirus in Spodoptera frugiperda cell lines",
    authors: "Ma H … Khan AS (FDA CBER)", slot: "원문 대표 Figure 삽입\n(TEM 이미지)",
  });
  txt(s, "Sf-rhabdovirus 게놈", { x: 4.6, y: 1.85, w: 8, h: 0.45, fontSize: 17, bold: true });
  const genes = [{ g: "N", w: 1.0 }, { g: "P", w: 0.85 }, { g: "M", w: 0.8 }, { g: "G", w: 1.2 }, { g: "X", w: 0.6, hot: true }, { g: "L", w: 3.0 }];
  let gx = 4.6;
  hline(s, 4.45, 12.7, 2.92, NAVY, 1.5);
  genes.forEach((g) => {
    rect(s, gx, 2.6, g.w - 0.08, 0.64, g.hot ? ORANGE : NAVY);
    txt(s, g.g, { x: gx, y: 2.6, w: g.w - 0.08, h: 0.64, fontSize: 18, bold: true, fontFace: NUMF, color: g.hot ? NAVY : WHITE, align: "center", valign: "middle" });
    gx += g.w;
  });
  txt(s, "X = G–L 사이 미지 ORF (111 aa)", { x: 4.6, y: 3.35, w: 8, h: 0.35, fontSize: 11.5, color: GRAY });
  kwStack(s, [
    { k: "신종 랩도바이러스", d: "Degenerate PCR + 대규모 병렬 시퀀싱", hot: true },
    { k: "Sf21 모세포주에도 존재", d: "다른 곤충세포주에는 없음" },
    { k: "인체 세포 복제 근거 없음", d: "위해성 평가 결론" },
  ], 4.6, 4.05, 8.1, 0.82);
}

// 4-5 서열 ≠ 감염성 (Geisler 2016)
{
  const s = base({
    section: "04 사례 · 함정", title: "서열 ≠ 감염성",
    cite: "Geisler C, Jarvis DL. Biologicals 2016;44:219 (doi:10.1016/j.biologicals.2016.04.004) · Ma H et al. J Virol 2014;88:6576",
    notes: [
      "요지: NGS는 서열을 볼 뿐 감염성을 보지 않는다. 같은 세포주·같은 바이러스 계통에서 정반대 결론.",
      "- Sf 세포 유전체·전사체에서 Sf-rhabdovirus 유사 서열 발견 → 전사체와 유전체 서열 일치 → 내인성 바이러스 유사 요소(EVE).",
      "- N·P 유사 ORF는 온전, G·L 유사 ORF는 부분적.",
      "- 저자 권고 요지: MPS 기반 외래성 바이러스 탐색은 유전체와 전사체를 함께 봐야 전사되는 EVE를 식별하고 위양성을 피할 수 있다.",
      "⚠ '4개 유전자 자리', 'M ORF 없음'은 원문 확인 전 게재 금지.",
      "- 1-7 정의(외래성 vs 내인성)로 귀결. 전사 여부 행이 핵심: 둘 다 전사되므로 RNA만으로는 구분 불가.",
    ].join("\n"),
  });
  const rows = [
    [H(""), H("Sf-rhabdovirus  (Ma 2014)"), H("Sf EVE  (Geisler 2016)")],
    [C("정체", { bold: true }), C("바이러스 입자", { align: "center" }), C("숙주 유전체 내 통합 요소", { align: "center" })],
    [C("ORF", { bold: true }), C("N · P · M · G · X · L", { align: "center", fontFace: NUMF }), C("N · P 온전 / G · L 부분", { align: "center" })],
    [C("전사", { bold: true }), HOT("예"), HOT("예")],
    [C("판정", { bold: true }), C("외래성 바이러스", { align: "center", bold: true }), C("내인성 요소", { align: "center", bold: true })],
  ];
  s.addTable(rows, { x: 0.6, y: 1.8, w: 12.1, colW: [2.1, 5.0, 5.0], rowH: 0.72, fontFace: FONT, fontSize: 15, border: TB });
  kwChips(s, [{ k: "EVE 전사", hot: true }, "불완전 ORF", "유전체 + 전사체"], 0.6, 5.75, 12.1, 0.6);
}

// 4-6 한 번에 보인 것 (Victoria 2010)
{
  const s = base({
    section: "04 사례 · 완제품", title: "한 번에 보인 것",
    cite: "Victoria JG, Wang C, Jones MS, Jaing C, McLoughlin K, Gardner S, Delwart EL. J Virol 2010;84:6033 (doi:10.1128/JVI.02690-09)",
    notes: [
      "요지: 같은 메타게놈 데이터가 외래성 오염체, 약독화주 변이체, 내인성 레트로바이러스를 한 번에 보여줬다.",
      "- 허가 생백신 8종(OPV, 풍진, 홍역, 황열, 수두, MMR, 로타바이러스 2종) 분석.",
      "- 외래성: Rotarix에서 PCV1.",
      "- Minority variant: OPV·볼거리·수두 백신.",
      "- 내인성 레트로바이러스: 조류 ALV(입자 내 RNA), 원숭이 SRV(결손 DNA).",
      "- 범미생물 마이크로어레이로 교차 확인.",
      "- 사건 경과는 인트로에서 다뤘으므로 여기서는 1분 이내.",
    ].join("\n"),
  });
  txt(s, "허가 생백신 8종 메타게놈", { x: 0.6, y: 1.8, w: 8, h: 0.45, fontSize: 16, color: GRAY });
  const st = [
    { n: "1", u: "외래성", k: "PCV1 (Rotarix)", hot: true },
    { n: "3", u: "백신", k: "Minority variant" },
    { n: "2", u: "종", k: "내인성 ALV · SRV" },
  ];
  st.forEach((x, i) => {
    const xx = 0.6 + i * 2.85;
    if (x.hot) rect(s, xx - 0.05, 2.55, 1.1, 1.4, ORANGE);
    txt(s, x.n, { x: xx, y: 2.5, w: 1.1, h: 1.5, fontSize: 80, bold: true, fontFace: NUMF, valign: "middle" });
    txt(s, x.u, { x: xx + 1.2, y: 3.35, w: 1.5, h: 0.5, fontSize: 15, bold: true });
    txt(s, x.k, { x: xx, y: 4.2, w: 2.7, h: 0.5, fontSize: 15, bold: true });
  });
  txt(s, "OPV · 볼거리 · 수두 백신의 변이체 / ALV = 입자 내 RNA · SRV = 결손 DNA", { x: 0.6, y: 5.3, w: 8.2, h: 0.7, fontSize: 12, color: GRAY });
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "J Virol · 2010", title: "Viral nucleic acids in live-attenuated vaccines: detection of minority variants and an adventitious virus",
    authors: "Victoria JG … Delwart EL",
  });
}

// 4-7 전통법도 잡았다 (Kerr & Nims)
{
  const s = base({
    section: "04 사례 · 현장", title: "전통법도 잡았다",
    cite: "Kerr A, Nims R. PDA J Pharm Sci Technol 2010;64:481 (PMID 21502056)",
    notes: [
      "요지: 전통적 in vitro 시험의 실제 작동 기록. 이 장을 빼면 발표가 선전물이 된다.",
      "- 10년간 다양한 포유류 세포배양 생물의약품 bulk harvest를 in vitro virus screening assay로 평가.",
      "- 검출: reovirus 2형, Cache Valley virus — 둘 다 CHO 세포 생산품.",
      "- 희소성은 곧 '한 건의 무게'를 뜻한다 → 다음 장(저항성 세포주)으로 연결.",
      "⚠ '문헌상 CHO에서 MVM 보고'는 본문 맥락 — 인용 시 원문 확인.",
    ].join("\n"),
  });
  txt(s, [
    { text: "10", options: { fontFace: NUMF, fontSize: 88, bold: true } },
    { text: " 년", options: { fontSize: 26, bold: true } },
  ], { x: 0.6, y: 1.8, w: 4.0, h: 1.7, valign: "middle" });
  rect(s, 4.55, 1.95, 1.55, 1.45, ORANGE);
  txt(s, [
    { text: "2", options: { fontFace: NUMF, fontSize: 88, bold: true } },
    { text: " 종", options: { fontSize: 26, bold: true } },
  ], { x: 4.6, y: 1.8, w: 3.6, h: 1.7, valign: "middle" });
  txt(s, "Reovirus 2형 · Cache Valley virus", { x: 0.6, y: 3.65, w: 8, h: 0.5, fontSize: 18, bold: true });
  kwStack(s, [
    { k: "In vitro 실제 검출", d: "Bulk harvest 선별 시험" },
    { k: "CHO 생산품", d: "검출 2종 모두", hot: true },
    { k: "희소한 사건", d: "10년간 2종" },
  ], 0.6, 4.35, 8.0, 0.78);
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "PDA J Pharm Sci Technol · 2010", title: "Adventitious viruses detected in biopharmaceutical bulk harvest samples over a 10 year period",
    authors: "Kerr A, Nims R",
  });
}

// 4-8 설계로 막는다 (Mascarenhas)
{
  const s = base({
    section: "04 사례 · 현장", title: "설계로 막는다",
    cite: "Mascarenhas JX et al. Biotechnol Bioeng 2017;114:576 (doi:10.1002/bit.26186)",
    notes: [
      "요지: 검출 다음 단계 — 세포주 설계로 위험 자체를 줄인다.",
      "- ZFN 유전자 편집으로 MVM 진입에 필요한 세포 표면 시알산 경로 제거.",
      "- Slc35a1(CMP-시알산 수송체) KO: 시알산 완전 소실 → MVM 감염에 완전 저항.",
      "- Cosmc·Mgat1 KO: 유의한 억제.",
      "⚠ 재조합 단백질 생산성·품질 유지, '파국적(catastrophic)' 인용은 본문 대조 필요.",
    ].join("\n"),
  });
  txt(s, [
    { text: "시알산 제거 → ", options: {} },
    { text: "MVM 진입 차단", options: { highlight: ORANGE } },
  ], { x: 0.6, y: 1.8, w: 8.3, h: 0.8, fontSize: 28, bold: true, valign: "middle" });
  const I3 = (t) => ({ text: t, options: { italic: true, bold: true, color: NAVY, fontFace: NUMF, valign: "middle" } });
  const rows = [
    [H("ZFN KO 표적"), H("MVM 감염")],
    [I3("Slc35a1"), HOT("완전 저항")],
    [I3("Cosmc"), SOFT("유의한 억제")],
    [I3("Mgat1"), SOFT("유의한 억제")],
  ];
  s.addTable(rows, { x: 0.6, y: 2.95, w: 8.2, colW: [4.1, 4.1], rowH: 0.62, fontFace: FONT, fontSize: 16, border: TB });
  kwChips(s, ["ZFN 편집", "시알산 경로", { k: "MVM 저항성", hot: false }], 0.6, 5.75, 8.2, 0.55);
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "Biotechnol Bioeng · 2017", title: "Genetic engineering of CHO cells for viral resistance to minute virus of mice",
    authors: "Mascarenhas JX, et al.",
  });
}

// 4-9 동물시험 5/9 (Beurdeley-Fehlbaum 2023)
{
  const s = base({
    section: "04 사례 · 성능", title: "동물시험 5 / 9",
    cite: "Beurdeley-Fehlbaum P et al. Vaccine 2023;41:5383 (doi:10.1016/j.vaccine.2023.07.010)",
    notes: [
      "요지: 동물시험은 윤리 때문만이 아니라 성능 때문에 대체되고 있다.",
      "- 세포배양에 9종 바이러스 접종 → in vivo(동물·발육란) 시험은 5종만 검출, NGS 전사체 분석은 9종 모두 검출.",
      "- 민감도: 감염세포 1개를 비감염세포 10^3–10^7개 중에서 검출. (기존 노트의 10^5–10^6은 오기, 정정 완료)",
      "- 결론: 세포기질 바이러스 안전성 시험에서 in vivo를 NGS로 대체하는 것을 지지.",
    ].join("\n"),
  });
  txt(s, "검출 폭 · 바이러스 9종 접종", { x: 0.6, y: 1.75, w: 6, h: 0.4, fontSize: 15, bold: true });
  s.addChart(pres.charts.BAR, [
    { name: "검출 바이러스 수", labels: ["In vivo", "NGS"], values: [5, 9] },
  ], {
    x: 0.5, y: 2.15, w: 6.3, h: 3.2, barDir: "bar",
    chartColors: [NAVY, ORANGE],
    valAxisMinVal: 0, valAxisMaxVal: 9, valAxisMajorUnit: 3,
    valAxisLabelColor: MUTED, catAxisLabelColor: NAVY, catAxisLabelFontSize: 14, valAxisLabelFontSize: 10,
    catAxisLabelFontFace: NUMF, valAxisLabelFontFace: NUMF,
    valGridLine: { color: "E6E6E6", size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelColor: NAVY,
    showLegend: false, barGapWidthPct: 60,
    catAxisOrientation: "maxMin",
  });
  // 오른쪽: 민감도 로그 축
  txt(s, "민감도 · 감염세포 1개 대비 비감염세포", { x: 7.3, y: 1.75, w: 5.4, h: 0.4, fontSize: 15, bold: true });
  const ax1 = 7.5, ax2 = 12.5, ay = 3.8;
  hline(s, ax1, ax2, ay, NAVY, 1.5);
  const pos = (e) => ax1 + (e / 8) * (ax2 - ax1);
  rect(s, pos(3), ay - 0.32, pos(7) - pos(3), 0.28, ORANGE);
  for (let e = 0; e <= 8; e++) {
    vline(s, pos(e), ay, ay + 0.1, NAVY, 1);
    if (e % 2 === 1 || e === 0 || e === 8) {
      txt(s, [{ text: "10" }, { text: String(e), options: { superscript: true } }], { x: pos(e) - 0.3, y: ay + 0.15, w: 0.6, h: 0.3, fontSize: 11, fontFace: NUMF, align: "center", color: GRAY });
    }
  }
  txt(s, [{ text: "1 : 10" }, { text: "3", options: { superscript: true } }, { text: " – 10" }, { text: "7", options: { superscript: true } }], { x: 7.3, y: 4.45, w: 5.4, h: 0.8, fontSize: 30, bold: true, fontFace: NUMF });
  kwChips(s, ["검출 폭", "민감도", { k: "In vivo 대체 지지", hot: true }], 0.6, 5.85, 12.1, 0.55);
}

// 4-10 1 copy (Hirai 2024)
{
  const s = base({
    section: "04 사례 · 성능", title: "1 copy / assay",
    cite: "Hirai T, Kataoka K, Yuan Y, Yusa K, Sato Y, Uchida K, Kono K. Biologicals 2024;85:101739 (doi:10.1016/j.biologicals.2023.101739)",
    notes: [
      "요지: 플랫폼(short vs long read)보다 설계가 중요. 둘 다 assay당 1 copy 수준까지 검출.",
      "- 일본 국립의약품식품위생연구소. 외래성 바이러스 in vitro 시험의 표준 지시세포인 Vero에 adenovirus 5 접종 → 감염세포 RNA를 비감염세포 RNA로 연속 희석.",
      "- Short-read vs long-read(Oxford Nanopore): 검출한계 거의 동일. 'sensitive enough to detect viral sequences as long as there was at least one copy in one assay.'",
      "⚠ 멀티플렉싱에 따른 검체 간 교차오염 위험 지적은 본문 대조 필요 → 확인되면 파트 3의 index 오배정과 연결.",
    ].join("\n"),
  });
  rect(s, 0.55, 1.95, 3.4, 1.9, ORANGE);
  txt(s, "1", { x: 0.6, y: 1.8, w: 3.3, h: 2.2, fontSize: 150, bold: true, fontFace: NUMF, align: "center", valign: "middle" });
  txt(s, "copy / assay", { x: 4.2, y: 2.4, w: 4.2, h: 0.7, fontSize: 30, bold: true, fontFace: NUMF });
  txt(s, "검출 한계 · Vero + adenovirus 5 RNA 희석계열", { x: 4.2, y: 3.15, w: 8.4, h: 0.45, fontSize: 14, color: GRAY });
  // short vs long
  const cmp = [{ k: "Short-read", d: "짧은 read 플랫폼" }, { k: "Long-read", d: "Oxford Nanopore" }];
  cmp.forEach((c, i) => {
    const x = 0.6 + i * 4.6;
    rect(s, x, 4.3, 3.6, 1.2, LIGHT);
    txt(s, c.k, { x, y: 4.38, w: 3.6, h: 0.6, fontSize: 22, bold: true, fontFace: NUMF, align: "center" });
    txt(s, c.d, { x, y: 4.95, w: 3.6, h: 0.4, fontSize: 12, color: GRAY, align: "center" });
  });
  txt(s, "≈", { x: 4.2, y: 4.3, w: 1.0, h: 1.2, fontSize: 44, bold: true, align: "center", valign: "middle", fontFace: NUMF });
  txt(s, "검출 한계 거의 동일", { x: 9.4, y: 4.3, w: 3.3, h: 1.2, fontSize: 18, bold: true, valign: "middle" });
  kwChips(s, ["지시세포 RNA 희석", "플랫폼 간 동등", "In vitro 시험 보완 근거"], 0.6, 5.9, 12.1, 0.55);
}

// 4-11 이미 교체했다 (Alston 2025)
{
  const s = base({
    section: "04 사례 · 구현", title: "이미 교체했다",
    cite: "Alston A, Bova RA, Hasson B. Biologicals 2025;90:101828 (doi:10.1016/j.biologicals.2025.101828) · ICH Q5A(R2) §3.2.5.2",
    notes: [
      "요지: '가능하다'(Beurdeley 2023)에서 '실제로 했다'로. 상용 제품 출하시험에 적용된 구현 사례.",
      "- 생백신 인플루엔자(LAIV) 상용 제품의 in vivo 외래성 인자 시험을 NGS 외래성 인자 검출법으로 교체.",
      "- 제품 매트릭스에 spike하여 검출한계(LoD) 설정·밸리데이션.",
      "- 규제 근거: ICH Q5A(R2) 3.2.5.2.",
      "⚠ AstraZeneca·MilliporeSigma 협업 명시 여부, '정면 비교 생략' 서술은 본문 대조 필요.",
      "- 계란 기반 백신 생산 맥락은 국내 백신 개발과도 가깝다.",
    ].join("\n"),
  });
  txt(s, "상용 LAIV · 로트 출하시험", { x: 0.6, y: 1.8, w: 8, h: 0.45, fontSize: 16, color: GRAY });
  rect(s, 0.6, 2.5, 3.5, 1.5, LIGHT);
  txt(s, "In vivo\n외래성 인자 시험", { x: 0.6, y: 2.5, w: 3.5, h: 1.5, fontSize: 18, bold: true, color: MUTED, align: "center", valign: "middle" });
  arrow(s, 4.25, 3.25, 5.05, 3.25, NAVY);
  rect(s, 5.2, 2.5, 3.5, 1.5, ORANGE);
  txt(s, "NGS\n외래성 인자 검출", { x: 5.2, y: 2.5, w: 3.5, h: 1.5, fontSize: 18, bold: true, align: "center", valign: "middle" });
  kwStack(s, [
    { k: "상용 생백신", d: "LAIV 로트별 시험" },
    { k: "매트릭스 spike · LoD", d: "제품 기반 밸리데이션" },
    { k: "ICH Q5A(R2) 근거", d: "§3.2.5.2", hot: true },
  ], 0.6, 4.45, 8.1, 0.72);
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "Biologicals · 2025", title: "Validation of a Next Generation Sequencing Method for adventitious agents detection in a live vaccine matrix",
    authors: "Alston A, Bova RA, Hasson B",
  });
}

// 4-12 합의에 이르기까지 (Khan 2025)
{
  const s = base({
    section: "04 사례 · 구현", title: "합의에 이르기까지",
    cite: "Mallet L, Gisonni-Lex L. PDA J Pharm Sci Technol 2014;68:556 · IABS 보고 Biologicals 2018;55:1 / 2020;67:94 / 2023;83:101696 · Khan AS et al. Biologicals 2025;92:101859",
    notes: [
      "요지: 규제기관이 공동 주재한 4차 컨퍼런스에서 대체 준비 합의 — '전망'이 아니라 '현재 좌표'.",
      "- 2013.11 PDA/FDA 'Advanced Technologies for Virus Detection' 워크숍(Bethesda).",
      "- IABS: 1차 2017.10 Rockville(IABS·FDA), 2차 2019.11 Ghent, 3차 2022.09 Rockville, 4차 2024.12.4–5 Frankfurt(FDA·EDQM 공동 주재).",
      "- 4차 합의: 'readiness of NGS to replace the in vivo adventitious virus detection assays and PCR assays, and to supplement or replace the in vitro cell-based assays, based on a suitable validation package.'",
      "⚠ 4차 보고의 PMID(40945333) 연결, 2차 보고 PMID 원문 확인.",
    ].join("\n"),
  });
  const X = (yr) => 1.3 + (yr - 2013) * (10.4 / 11.5);
  timeline(s, 2.95, 0.9, 12.4, [
    { x: X(2013.85), year: "2013", label: "PDA/FDA 워크숍\nBethesda", up: false, w: 1.9 },
    { x: X(2017.8), year: "2017", label: "IABS 1차\nRockville", up: false, w: 1.9 },
    { x: X(2019.85), year: "2019", label: "IABS 2차\nGhent", up: false, w: 1.9 },
    { x: X(2022.7), year: "2022", label: "IABS 3차\nRockville", up: false, w: 1.9 },
    { x: X(2024.9), year: "2024", label: "IABS 4차\nFrankfurt", up: false, w: 1.9, hot: true },
  ]);
  const cons = [
    { k: "In vivo · PCR", v: "대체 준비", hot: true },
    { k: "In vitro 세포배양", v: "보완 · 대체" },
    { k: "전제 조건", v: "밸리데이션 패키지" },
  ];
  cons.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    rect(s, x, 4.95, 3.8, 1.45, c.hot ? ORANGE : LIGHT);
    txt(s, c.k, { x: x + 0.25, y: 5.05, w: 3.4, h: 0.4, fontSize: 13, color: c.hot ? NAVY : GRAY });
    txt(s, c.v, { x: x + 0.25, y: 5.5, w: 3.4, h: 0.7, fontSize: 24, bold: true });
  });
}

// =====================================================================
// 맺음
{
  const s = base({
    section: "정리", title: "정리",
    cite: "ICH Q5A(R2) 2023 · Ph. Eur. 2.6.41 (2026-04-01 시행) · Geisler C, Jarvis DL. Biologicals 2016;44:219",
    notes: [
      "Take-home (2분).",
      "1. 정의: 외래성의 기준점은 세포기질 유전체. 내인성과의 경계를 NGS 해석에서도 유지.",
      "2. 규제: ICH Q5A(R2)(2023)와 Ph. Eur. 2.6.41(2026-04 시행)로 non-targeted NGS는 이미 현행 규제 선택지.",
      "3. 기술: 성능은 시퀀서가 아니라 전처리 설계, host removal·DB·판정 기준을 포함한 in-silico 파이프라인과 그 밸리데이션이 결정.",
      "4. 한계: 서열은 감염성이 아니다. 확증시험과 전통법의 역할은 남는다.",
    ].join("\n"),
  });
  txt(s, [
    { text: "NGS는 ", options: {} },
    { text: "목록 밖", options: { highlight: ORANGE } },
    { text: "을 본다", options: {} },
  ], { x: 0.6, y: 1.8, w: 12.1, h: 1.1, fontSize: 44, bold: true, valign: "middle" });
  txt(s, "판정은 여전히 확증시험이 한다", { x: 0.6, y: 3.05, w: 12.1, h: 0.7, fontSize: 26, color: GRAY, valign: "middle" });
  const items = [
    { k: "정의", v: "기준점은 세포기질 유전체" },
    { k: "규제", v: "ICH Q5A(R2) · Ph. Eur. 2.6.41 현행" },
    { k: "기술", v: "파이프라인과 밸리데이션" },
  ];
  items.forEach((it, i) => {
    const x = 0.6 + i * 4.15;
    txt(s, it.k, { x, y: 4.55, w: 3.8, h: 0.45, fontSize: 14, bold: true, color: MUTED });
    hline(s, x, x + 3.7, 5.05, NAVY, 1);
    txt(s, it.v, { x, y: 5.2, w: 3.8, h: 0.9, fontSize: 18, bold: true });
  });
}

// =====================================================================
// Supplementary
divider("S", "Supplementary", ["Bas-Congo virus", "In vivo 상세", "References"], "발표하지 않음",
  "부록. 본 발표에서는 넘기고, 질의응답 대비용으로만 사용.");

// S1 Bas-Congo
{
  const s = base({
    section: "Supplementary", title: "PCR이 닿지 않는 거리",
    cite: "Grard G et al. PLoS Pathog 2012;8:e1002924 (doi:10.1371/journal.ppat.1002924) · 수치 원문 대조 필요",
    notes: [
      "용도: 질의응답에서 'PCR 패널을 넓히면 되지 않나'에 대한 답.",
      "- 2009년 콩고민주공화국 급성 출혈열 3례. 고도로 발산한 신종 랩도바이러스(Bas-Congo virus).",
      "- 발산도가 이 정도면 degenerate PCR 프라이머 설계 자체가 원리적으로 어렵다 → unbiased NGS + de novo assembly.",
      "⚠ 원문 대조 전: 기존 랩도바이러스와 아미노산 동일성 34% 미만, 혈청 1.09×10^6 copies/mL, 약 1억 4천만 read, 유전체 98.2% de novo assembly.",
    ].join("\n"),
  });
  rect(s, 0.55, 2.0, 4.3, 1.75, ORANGE);
  txt(s, "< 34%", { x: 0.6, y: 1.85, w: 4.2, h: 2.05, fontSize: 80, bold: true, fontFace: NUMF, align: "center", valign: "middle" });
  txt(s, "기존 랩도바이러스와의 아미노산 동일성", { x: 5.2, y: 2.35, w: 3.6, h: 1.1, fontSize: 18, bold: true, valign: "middle" });
  kwStack(s, [
    { k: "신종 랩도바이러스", d: "2009 콩고민주공화국 · 급성 출혈열" },
    { k: "Degenerate PCR 한계", d: "프라이머 설계가 닿지 않는 발산도" },
    { k: "De novo assembly", d: "unbiased NGS로 유전체 복원" },
  ], 0.6, 4.2, 8.1, 0.78);
  paperCard(s, 9.25, 1.8, 3.45, 4.6, {
    journal: "PLoS Pathog · 2012", title: "A novel rhabdovirus associated with acute hemorrhagic fever in central Africa",
    authors: "Grard G, et al.",
  });
}

// S2 In vivo 상세
{
  const s = base({
    section: "Supplementary", title: "In vivo 시험 상세",
    cite: "FDA. Characterization and Qualification of Cell Substrates and Other Biological Materials Used in the Production of Viral Vaccines for Infectious Disease Indications (2010)",
    notes: "FDA 2010 가이던스의 in vivo 시험 표(동물 수, 투여 경로, 관찰 기간, 표적 바이러스군)를 원문에서 캡처하거나 옮겨 적을 것. 검증 단계에서 원문 PDF를 열지 못해 수치를 싣지 않았다.",
  });
  kwChips(s, ["유약 · 성숙 마우스", "발육란", "기니피그"], 0.6, 1.8, 12.1, 0.55);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.65, w: 12.1, h: 4.0, fill: { color: WHITE }, line: { color: "A6A6A6", width: 1, dashType: "dash" } });
  txt(s, "FDA 2010 가이던스 in vivo 시험 표 원문 삽입\n(동물 수 · 투여 경로 · 관찰 기간 · 표적 바이러스군)", { x: 0.6, y: 2.65, w: 12.1, h: 4.0, fontSize: 13, color: MUTED, align: "center", valign: "middle" });
}

// References
const REFS = [
  ["사례 문헌", [
    "Ma H et al. J Virol 2014;88:6576. doi:10.1128/JVI.00780-14",
    "Onions D et al. Vaccine 2011;29:7117. doi:10.1016/j.vaccine.2011.05.071",
    "Shabram P, Kolman JL. PDA J Pharm Sci Technol 2014;68:639. doi:10.5731/pdajpst.2014.01027",
    "Gagnieur L et al. Biologicals 2014;42:145. doi:10.1016/j.biologicals.2014.02.002",
    "Zhang P et al. Zool Res 2022;43:756. doi:10.24272/j.issn.2095-8137.2022.093",
    "Victoria JG et al. J Virol 2010;84:6033. doi:10.1128/JVI.02690-09",
    "Kerr A, Nims R. PDA J Pharm Sci Technol 2010;64:481. PMID 21502056",
    "Mascarenhas JX et al. Biotechnol Bioeng 2017;114:576. doi:10.1002/bit.26186",
    "Beurdeley-Fehlbaum P et al. Vaccine 2023;41:5383. doi:10.1016/j.vaccine.2023.07.010",
    "Alston A et al. Biologicals 2025;90:101828. doi:10.1016/j.biologicals.2025.101828",
    "Hirai T et al. Biologicals 2024;85:101739. doi:10.1016/j.biologicals.2023.101739",
    "Khan AS et al. Biologicals 2025;92:101859. doi:10.1016/j.biologicals.2025.101859",
    "Geisler C, Jarvis DL. Biologicals 2016;44:219. doi:10.1016/j.biologicals.2016.04.004",
    "Grard G et al. PLoS Pathog 2012;8:e1002924. doi:10.1371/journal.ppat.1002924",
  ]],
  ["규제 · 사건", [
    "ICH Q5A(R2) Step 4, 2023-11-01 · FDA 89 FR 1925 (2024) · EMA/CHMP/ICH/804363/2022",
    "EDQM. Ph. Eur. 2.6.41 채택 공지 (2025) · Ph. Eur. 12.2",
    "FDA. Cell Substrates Guidance for Viral Vaccines (2010)",
    "WHO TRS 978 Annex 3 (2013) · WHO/BS/2024.2471",
    "식약처. 생물의약품 외래성 바이러스 부정시험 가이드라인 (2010)",
    "식품의약품안전평가원. 백신 생산용 세포주의 NGS 기반 외래성 바이러스 부정시험법 정보집 (2021)",
    "IABS NGS 컨퍼런스 보고: Biologicals 2018;55:1 · 2020;67:94 · 2023;83:101696",
    "Mallet L, Gisonni-Lex L. PDA J Pharm Sci Technol 2014;68:556",
    "Dubin G et al. Hum Vaccin Immunother 2013;9:2398. doi:10.4161/hv.25973",
    "Barone PW et al. Nat Biotechnol 2020;38:563. doi:10.1038/s41587-020-0507-2",
    "Garnick RL. Dev Biol Stand 1996;88:49 · Nims RW et al. BioPharm Int 2008",
    "CDC Historical concerns (SV40) · IOM 2002 · Genzyme Form 10-K FY2009",
  ]],
  ["NGS 원리 · 분석", [
    "Illumina. An Introduction to NGS Technology · Patterned Flow Cell · 2-Channel SBS",
    "Illumina white paper 770-2017-004 (index misassignment)",
    "Bentley DR et al. Nature 2008;456:53",
    "Allander T et al. PNAS 2001;98:11609",
    "Ng SHS et al. Viruses 2018;10:566",
    "Mee ET et al. Vaccine 2016;34:2035",
    "Salter SJ et al. BMC Biol 2014;12:87",
    "Khan AS et al. mSphere 2017;2:e00307-17",
    "Bolger AM et al. Bioinformatics 2014;30:2114 · Chen S et al. Bioinformatics 2018;34:i884",
    "Li H, Durbin R. Bioinformatics 2009;25:1754 · Langmead B, Salzberg SL. Nat Methods 2012;9:357",
    "Buchfink B et al. Nat Methods 2015;12:59 · Wood DE et al. Genome Biol 2019;20:257",
    "Bankevich A et al. J Comput Biol 2012;19:455 · Li D et al. Bioinformatics 2015;31:1674",
    "Goodacre N et al. mSphere 2018;3:e00069-18 · Chin P et al. mSphere 2025;10:e00286-25",
    "MacDonald ML et al. mSphere 2021;6:e01336-20 · Simonyan V et al. Database 2016;baw022",
    "Xu X et al. Nat Biotechnol 2011;29:735 · Lewis NE et al. Nat Biotechnol 2013;31:759",
    "Rupp O et al. Biotechnol Bioeng 2018;115:2087 · Osada N et al. DNA Res 2014;21:673",
    "Lin YC et al. Nat Commun 2014;5:4767 · Nandakumar S et al. Genome Announc 2017;5:e00829-17",
    "Chiu CY, Miller SA. Nat Rev Genet 2019;20:341",
  ]],
];
{
  // References 1: 사례 문헌 (좌) + 규제·사건 (우)
  const s = base({ section: "Supplementary", title: "References (1/2)", notes: "참고문헌. 전체 링크는 notes/03_발표콘티.md 참조." });
  [[REFS[0], 0.6], [REFS[1], 6.85]].forEach(([grp, x]) => {
    txt(s, grp[0], { x, y: 1.6, w: 5.9, h: 0.35, fontSize: 13, bold: true });
    txt(s, grp[1].map((r, j) => ({ text: r, options: { breakLine: j < grp[1].length - 1 } })), { x, y: 2.0, w: 5.9, h: 4.85, fontSize: 9, color: GRAY, fontFace: NUMF, paraSpaceAfter: 3 });
  });
}
{
  const s = base({ section: "Supplementary", title: "References (2/2)", notes: "참고문헌. 전체 링크는 notes/03_발표콘티.md 참조." });
  const grp = REFS[2];
  const half = Math.ceil(grp[1].length / 2);
  txt(s, grp[0], { x: 0.6, y: 1.6, w: 12, h: 0.35, fontSize: 13, bold: true });
  [[grp[1].slice(0, half), 0.6], [grp[1].slice(half), 6.85]].forEach(([arr, x]) => {
    txt(s, arr.map((r, j) => ({ text: r, options: { breakLine: j < arr.length - 1 } })), { x, y: 2.0, w: 5.9, h: 4.85, fontSize: 9.5, color: GRAY, fontFace: NUMF, paraSpaceAfter: 4 });
  });
}

pres.writeFile({ fileName: process.argv[2] || "NGS_외래성바이러스_초안.pptx" }).then((f) => console.log("written", f, "slides", page));
