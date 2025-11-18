const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

/**
 * Manually create the 6 missing slides using PptxGenJS API
 * These slides failed html2pptx conversion due to layout constraints
 */

async function addMissingSlides() {
  // Load the existing PPTX with 17 slides
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Claude Code";
  pptx.title = "전략적 재고운영 및 자재계획 수립";

  // Define color palette
  const colors = {
    primary: "1a5276",
    secondary: "3498db",
    white: "FFFFFF",
    muted: "f5f5f5",
    mutedText: "737373",
    warning: "fff3cd",
    warningBorder: "ffc107",
    info: "e3f2fd",
    infoBorder: "3498db",
  };

  // Helper function to add footer
  function addFooter(slide, text) {
    slide.addText(text, {
      x: 8.5,
      y: 4.8,
      w: 1.3,
      h: 0.3,
      fontSize: 10,
      color: colors.mutedText,
      align: "right",
    });
  }

  // Slide 9: Kraljic Matrix 탄생
  const slide9 = pptx.addSlide();
  slide9.background = { color: colors.white };

  slide9.addText("Kraljic Matrix의 탄생과 의미", {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  slide9.addText("1983년, 석유파동이 낳은 혁신", {
    x: 0.5,
    y: 1.1,
    w: 9,
    h: 0.3,
    fontSize: 18,
    color: colors.mutedText,
  });

  // Box 1: 탄생 배경
  slide9.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.6,
    w: 9,
    h: 1.0,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide9.addText([
    { text: "탄생 배경\n", options: { fontSize: 16, bold: true, color: colors.primary } },
    { text: "• Peter Kraljic, HBR 발표: \"Purchasing Must Become Supply Management\"\n", options: { fontSize: 12 } },
    { text: "• 1970년대 석유파동 → \"모든 자재 동일 관리\" 방식의 한계\n", options: { fontSize: 12 } },
    { text: "• 차별화된 접근의 필요성 대두", options: { fontSize: 12 } },
  ], {
    x: 0.7,
    y: 1.75,
    w: 8.6,
    h: 0.8,
  });

  // Box 2: 핵심 통찰
  slide9.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 2.8,
    w: 9,
    h: 1.2,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide9.addText("Kraljic Matrix의 핵심 통찰", {
    x: 0.7,
    y: 2.95,
    w: 8.6,
    h: 0.3,
    fontSize: 14,
    bold: true,
    color: colors.primary,
  });

  slide9.addText([
    { text: "\"Not all materials are created equal\"\n", options: { fontSize: 16, bold: true } },
    { text: "모든 자재가 동등하게 만들어지지 않았다.\n자재의 특성에 따라 차별화된 전략이 필요하다.", options: { fontSize: 13 } },
  ], {
    x: 0.7,
    y: 3.35,
    w: 8.6,
    h: 0.6,
    align: "center",
  });

  // Box 3: 중요한 오해 해소
  slide9.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 4.2,
    w: 9,
    h: 0.5,
    fill: { color: colors.warning },
    line: { color: colors.warningBorder, width: 2 },
  });

  slide9.addText([
    { text: "⚠️ 중요한 오해 해소: ", options: { fontSize: 13, bold: true } },
    { text: "JIC ≠ 무조건 재고 증가\n", options: { fontSize: 13, bold: true } },
    { text: "JIC는 \"모든 자재의 재고를 늘리자\"가 아닙니다. 자재 특성에 따라 차별화하는 것입니다.", options: { fontSize: 11 } },
  ], {
    x: 0.7,
    y: 4.3,
    w: 8.6,
    h: 0.35,
  });

  addFooter(slide9, "1회차 | Kraljic Matrix");

  // Slide 12: 병목자재
  const slide12 = pptx.addSlide();
  slide12.background = { color: colors.white };

  slide12.addText("🔴 병목자재 (Bottleneck Items)", {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  slide12.addText("높은 공급 리스크 + 낮은 구매 임팩트", {
    x: 0.5,
    y: 1.1,
    w: 9,
    h: 0.3,
    fontSize: 16,
    color: colors.mutedText,
  });

  // Left column: 특징 + 사례
  slide12.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.6,
    w: 4.4,
    h: 1.3,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide12.addText([
    { text: "특징\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 금액은 작지만 없으면 생산 중단\n", options: { fontSize: 11 } },
    { text: "• 공급업체가 1-2개로 매우 제한적\n", options: { fontSize: 11 } },
    { text: "• 대체 자재나 공급선을 찾기 어려움\n", options: { fontSize: 11 } },
    { text: "• 리드타임이 길거나 불안정", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 1.75,
    w: 4.1,
    h: 1.1,
  });

  slide12.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 3.0,
    w: 4.4,
    h: 1.0,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide12.addText([
    { text: "사례\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 차량용 MCU\n• 특수 규격 센서\n• 희소 원자재\n• 인증 필요 부품", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 3.15,
    w: 4.1,
    h: 0.8,
  });

  // Right column: 핵심 과제 & 관리 전략
  slide12.addShape(pptx.ShapeType.rect, {
    x: 5.1,
    y: 1.6,
    w: 4.4,
    h: 1.8,
    fill: { color: colors.warning },
    line: { color: colors.warningBorder, width: 2 },
  });

  slide12.addText([
    { text: "핵심 과제 & 관리 전략\n", options: { fontSize: 14, bold: true } },
    { text: "목표: 공급 안정성 확보 | 철학: \"비용보다 공급이 우선\" | KPI: 가용률 95%+\n\n", options: { fontSize: 10 } },
    { text: "• 안전재고: 4-8주 확보\n", options: { fontSize: 11 } },
    { text: "• 공급업체: 1개 → 2-3개 다변화\n", options: { fontSize: 11 } },
    { text: "• 계약: 1-3년 중장기 공급 보증\n", options: { fontSize: 11 } },
    { text: "• 발주방식: ROP (재주문점)", options: { fontSize: 11 } },
  ], {
    x: 5.25,
    y: 1.75,
    w: 4.1,
    h: 1.6,
  });

  slide12.addShape(pptx.ShapeType.rect, {
    x: 5.1,
    y: 3.5,
    w: 4.4,
    h: 0.5,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide12.addText([
    { text: "보험 관점: ", options: { fontSize: 11, bold: true } },
    { text: "안전재고 비용(연 5억) ≪ 생산 중단 손실(하루 100억)\n재고는 \"보험료\"입니다", options: { fontSize: 10 } },
  ], {
    x: 5.25,
    y: 3.6,
    w: 4.1,
    h: 0.35,
  });

  addFooter(slide12, "1회차 | 4대 자재군");

  // Slide 14: 전략자재
  const slide14 = pptx.addSlide();
  slide14.background = { color: colors.white };

  slide14.addText("🟣 전략자재 (Strategic Items)", {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  slide14.addText("높은 공급 리스크 + 높은 구매 임팩트", {
    x: 0.5,
    y: 1.1,
    w: 9,
    h: 0.3,
    fontSize: 16,
    color: colors.mutedText,
  });

  // Left column
  slide14.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.6,
    w: 4.4,
    h: 1.3,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide14.addText([
    { text: "특징\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 금액도 크고 공급도 어려움\n", options: { fontSize: 11 } },
    { text: "• 사업의 성패를 좌우\n", options: { fontSize: 11 } },
    { text: "• 대체 불가능, 전환 비용 높음\n", options: { fontSize: 11 } },
    { text: "• 장기 개발 필요", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 1.75,
    w: 4.1,
    h: 1.1,
  });

  slide14.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 3.0,
    w: 4.4,
    h: 1.0,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide14.addText([
    { text: "사례\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 핵심 반도체 (AP, SoC)\n• OLED 발광재료\n• 장납기 외자재\n• 독점 기술 부품", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 3.15,
    w: 4.1,
    h: 0.8,
  });

  // Right column
  slide14.addShape(pptx.ShapeType.rect, {
    x: 5.1,
    y: 1.6,
    w: 4.4,
    h: 1.8,
    fill: { color: colors.warning },
    line: { color: colors.warningBorder, width: 2 },
  });

  slide14.addText([
    { text: "핵심 과제 & 관리 전략\n", options: { fontSize: 14, bold: true } },
    { text: "목표: 전략적 파트너십 | 철학: \"Win-Win 상호 성장\" | KPI: 연속성 100%\n\n", options: { fontSize: 10 } },
    { text: "• 안전재고: 3-6주 중상 수준\n", options: { fontSize: 11 } },
    { text: "• 공급업체: 1-2개 전략적 파트너\n", options: { fontSize: 11 } },
    { text: "• 계약: 3-5년 장기 (안정성)\n", options: { fontSize: 11 } },
    { text: "• 발주방식: LTP + Hybrid", options: { fontSize: 11 } },
  ], {
    x: 5.25,
    y: 1.75,
    w: 4.1,
    h: 1.6,
  });

  slide14.addShape(pptx.ShapeType.rect, {
    x: 5.1,
    y: 3.5,
    w: 4.4,
    h: 0.5,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide14.addText([
    { text: "파트너십 관점: ", options: { fontSize: 11, bold: true } },
    { text: "단기 절감(연 5억) < 장기 가치(연 50억)\n관계가 곧 자산", options: { fontSize: 10 } },
  ], {
    x: 5.25,
    y: 3.6,
    w: 4.1,
    h: 0.35,
  });

  addFooter(slide14, "1회차 | 4대 자재군");

  // Slide 21: 학습 여정
  const slide21 = pptx.addSlide();
  slide21.background = { color: colors.white };

  slide21.addText("7회차 학습 여정", {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  slide21.addText("전략적 재고운영 완전 마스터 로드맵", {
    x: 0.5,
    y: 1.1,
    w: 9,
    h: 0.3,
    fontSize: 16,
    color: colors.mutedText,
  });

  // Module 1
  slide21.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.6,
    w: 9,
    h: 0.75,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide21.addText([
    { text: "Module 1: Foundation (1-2회차)\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 1회차: JIT→JIC + Kraljic Matrix (← 지금)\n", options: { fontSize: 11 } },
    { text: "• 2회차: 소싱 전략 + 공급업체 관리 프로세스", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 1.7,
    w: 8.7,
    h: 0.6,
  });

  // Module 2
  slide21.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 2.5,
    w: 9,
    h: 1.0,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide21.addText([
    { text: "Module 2: 자재군별 심화 (3-6회차)\n", options: { fontSize: 14, bold: true, color: colors.primary } },
    { text: "• 3회차: 병목자재 + ROP\n• 4회차: 레버리지자재 + MRP\n• 5회차: 전략자재 + LTP\n• 6회차: 일상자재 + 자동화", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 2.6,
    w: 8.7,
    h: 0.8,
  });

  // Module 3
  slide21.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 3.65,
    w: 9,
    h: 0.5,
    fill: { color: colors.warning },
    line: { color: colors.warningBorder, width: 2 },
  });

  slide21.addText([
    { text: "Module 3: 실전 통합 (7회차)\n", options: { fontSize: 14, bold: true } },
    { text: "• 7회차: Kraljic Matrix 실전 워크샵", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 3.75,
    w: 8.7,
    h: 0.35,
  });

  addFooter(slide21, "1회차 | 학습 여정");

  // Slide 22: 핵심 요약
  const slide22 = pptx.addSlide();
  slide22.background = { color: colors.white };

  slide22.addText("핵심 요약", {
    x: 0.5,
    y: 0.4,
    w: 9,
    h: 0.4,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  // Box 1: 패러다임 전환
  slide22.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.0,
    w: 9,
    h: 0.85,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide22.addText([
    { text: "1. 패러다임의 전환\n", options: { fontSize: 13, bold: true, color: colors.primary } },
    { text: "JIT (과거): 재고 = 낭비, 효율성 추구, 획일적 관리\n", options: { fontSize: 10 } },
    { text: "→ JIC (현재/미래): 재고 = 전략적 자산, 회복력 확보, 차별화된 전략", options: { fontSize: 10 } },
  ], {
    x: 0.65,
    y: 1.1,
    w: 8.7,
    h: 0.7,
  });

  // Box 2 & 3: Kraljic + 방법론
  slide22.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 2.0,
    w: 4.4,
    h: 1.0,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide22.addText([
    { text: "2. Kraljic Matrix\n", options: { fontSize: 13, bold: true, color: colors.primary } },
    { text: "• 2개 축: 공급 리스크 × 구매 임팩트\n", options: { fontSize: 10 } },
    { text: "• 4개 자재군: 병목/레버리지/전략/일상\n", options: { fontSize: 10 } },
    { text: "• 차별화 전략: 각 자재군별 맞춤형", options: { fontSize: 10 } },
  ], {
    x: 0.65,
    y: 2.1,
    w: 4.1,
    h: 0.85,
  });

  slide22.addShape(pptx.ShapeType.rect, {
    x: 5.1,
    y: 2.0,
    w: 4.4,
    h: 1.0,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide22.addText([
    { text: "3. 자재계획 방법론\n", options: { fontSize: 13, bold: true, color: colors.primary } },
    { text: "• 병목 → ROP (지속 모니터링)\n", options: { fontSize: 10 } },
    { text: "• 레버리지 → MRP (계획 기반)\n", options: { fontSize: 10 } },
    { text: "• 전략 → LTP + Hybrid\n", options: { fontSize: 10 } },
    { text: "• 일상 → Min-Max / VMI", options: { fontSize: 10 } },
  ], {
    x: 5.25,
    y: 2.1,
    w: 4.1,
    h: 0.85,
  });

  // Box 4: 과정의 가치
  slide22.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 3.15,
    w: 9,
    h: 0.5,
    fill: { color: colors.warning },
    line: { color: colors.warningBorder, width: 2 },
  });

  slide22.addText([
    { text: "4. 본 과정의 가치: ", options: { fontSize: 12, bold: true } },
    { text: "단순한 이론이 아닌, 즉시 적용 가능한 구체적 방안을 제공합니다", options: { fontSize: 11 } },
  ], {
    x: 0.65,
    y: 3.3,
    w: 8.7,
    h: 0.25,
  });

  addFooter(slide22, "1회차 | 요약");

  // Slide 23: 다음 회차 예고
  const slide23 = pptx.addSlide();
  slide23.background = { color: colors.white };

  slide23.addText("다음 회차 예고", {
    x: 0.5,
    y: 0.5,
    w: 9,
    h: 0.5,
    fontSize: 28,
    bold: true,
    color: colors.primary,
  });

  slide23.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 1.2,
    w: 9,
    h: 1.2,
    fill: { color: colors.info },
    line: { color: colors.infoBorder, width: 2 },
  });

  slide23.addText([
    { text: "2회차: 소싱 전략 및 공급업체 관계 관리\n\n", options: { fontSize: 16, bold: true, color: colors.primary } },
    { text: "• 자재군별 차별화된 소싱 전략\n", options: { fontSize: 12 } },
    { text: "• SRM(Supplier Relationship Management) 프레임워크\n", options: { fontSize: 12 } },
    { text: "• 계약 전략 및 협상 포인트\n", options: { fontSize: 12 } },
    { text: "• 공급업체 성과 평가 체계 (Supplier Scorecard)", options: { fontSize: 12 } },
  ], {
    x: 0.65,
    y: 1.35,
    w: 8.7,
    h: 1.0,
  });

  slide23.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: 2.55,
    w: 9,
    h: 0.8,
    fill: { color: colors.muted },
    line: { color: colors.primary, width: 3, dashType: "solid" },
  });

  slide23.addText([
    { text: "강사 TIP\n", options: { fontSize: 13, bold: true, color: colors.primary } },
    { text: "Kraljic Matrix는 단순한 분류 도구가 아닙니다. 조직 전체가 자재를 바라보는 공통 언어입니다.\n\n", options: { fontSize: 11 } },
    { text: "다음 회차부터는 각 자재군별 구체적인 전략과 방법론을 배우게 됩니다!", options: { fontSize: 11, bold: true } },
  ], {
    x: 0.65,
    y: 2.65,
    w: 8.7,
    h: 0.65,
  });

  slide23.addText("감사합니다!", {
    x: 0.5,
    y: 3.6,
    w: 9,
    h: 0.4,
    fontSize: 36,
    bold: true,
    color: colors.primary,
    align: "center",
  });

  addFooter(slide23, "1회차 | 전략적 재고운영");

  // Save PPTX
  const outputPath = "C:\\Users\\ahfif\\SuperClaude\\Project_Strategic_edu\\pptx\\Part1\\Part1_전략적재고운영Foundation_Missing6slides.pptx";
  await pptx.writeFile({ fileName: outputPath });

  console.log(`\n✅ Successfully created PPTX with 6 missing slides!`);
  console.log(`📁 Output: ${outputPath}`);
  console.log(`\n📋 Created slides: 9, 12, 14, 21, 22, 23`);
}

// Run
addMissingSlides().catch(console.error);
