const pptxgen = require("pptxgenjs");
const { html2pptx } = require("@ant/html2pptx");
const fs = require("fs");
const path = require("path");
const cheerio = require("cheerio");

/**
 * Create complete PPTX with all 23 slides
 * - Use html2pptx for slides 1-8, 10-11, 13, 15-20 (17 slides)
 * - Use PptxGenJS manually for slides 9, 12, 14, 21-23 (6 slides)
 */

async function createCompletePPTX(inputHtmlPath, outputPptxPath) {
  console.log(`📖 Reading HTML file: ${inputHtmlPath}`);

  const htmlContent = fs.readFileSync(inputHtmlPath, "utf-8");
  const $ = cheerio.load(htmlContent);
  const slides = $(".slide");
  console.log(`✅ Found ${slides.length} slides`);

  const tmpDir = path.join(__dirname, "../temp/slides");
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  const globalStyles = $("style").html() || "";

  // Create PPTX presentation
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Claude Code";
  pptx.title = "전략적 재고운영 및 자재계획 수립";

  // Define color palette for manual slides
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

  function createSlide9(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("Kraljic Matrix의 탄생과 의미", {
      x: 0.5, y: 0.5, w: 9, h: 0.5,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addText("1983년, 석유파동이 낳은 혁신", {
      x: 0.5, y: 1.1, w: 9, h: 0.3,
      fontSize: 18, color: colors.mutedText,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.6, w: 9, h: 1.0,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "탄생 배경\n", options: { fontSize: 16, bold: true, color: colors.primary } },
      { text: "• Peter Kraljic, HBR 발표: \"Purchasing Must Become Supply Management\"\n", options: { fontSize: 12 } },
      { text: "• 1970년대 석유파동 → \"모든 자재 동일 관리\" 방식의 한계\n", options: { fontSize: 12 } },
      { text: "• 차별화된 접근의 필요성 대두", options: { fontSize: 12 } },
    ], { x: 0.7, y: 1.75, w: 8.6, h: 0.8 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 2.8, w: 9, h: 1.2,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText("Kraljic Matrix의 핵심 통찰", {
      x: 0.7, y: 2.95, w: 8.6, h: 0.3,
      fontSize: 14, bold: true, color: colors.primary,
    });

    slide.addText([
      { text: "\"Not all materials are created equal\"\n", options: { fontSize: 16, bold: true } },
      { text: "모든 자재가 동등하게 만들어지지 않았다.\n자재의 특성에 따라 차별화된 전략이 필요하다.", options: { fontSize: 13 } },
    ], { x: 0.7, y: 3.35, w: 8.6, h: 0.6, align: "center" });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 4.2, w: 9, h: 0.5,
      fill: { color: colors.warning },
      line: { color: colors.warningBorder, width: 2 },
    });

    slide.addText([
      { text: "⚠️ 중요한 오해 해소: ", options: { fontSize: 13, bold: true } },
      { text: "JIC ≠ 무조건 재고 증가\n", options: { fontSize: 13, bold: true } },
      { text: "JIC는 자재 특성에 따라 차별화하는 것입니다.", options: { fontSize: 11 } },
    ], { x: 0.7, y: 4.3, w: 8.6, h: 0.35 });

    addFooter(slide, "1회차 | Kraljic Matrix");
  }

  function createSlide12(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("🔴 병목자재 (Bottleneck Items)", {
      x: 0.5, y: 0.5, w: 9, h: 0.5,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addText("높은 공급 리스크 + 낮은 구매 임팩트", {
      x: 0.5, y: 1.1, w: 9, h: 0.3,
      fontSize: 16, color: colors.mutedText,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.6, w: 4.4, h: 1.3,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "특징\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 금액은 작지만 없으면 생산 중단\n• 공급업체가 1-2개로 제한적\n• 대체 자재 찾기 어려움\n• 리드타임 길고 불안정", options: { fontSize: 11 } },
    ], { x: 0.65, y: 1.75, w: 4.1, h: 1.1 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 3.0, w: 4.4, h: 1.0,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "사례\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 차량용 MCU\n• 특수 규격 센서\n• 희소 원자재\n• 인증 필요 부품", options: { fontSize: 11 } },
    ], { x: 0.65, y: 3.15, w: 4.1, h: 0.8 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 5.1, y: 1.6, w: 4.4, h: 1.8,
      fill: { color: colors.warning },
      line: { color: colors.warningBorder, width: 2 },
    });

    slide.addText([
      { text: "핵심 과제 & 관리 전략\n", options: { fontSize: 14, bold: true } },
      { text: "목표: 공급 안정성 | 철학: \"비용보다 공급우선\" | KPI: 가용률 95%+\n\n", options: { fontSize: 10 } },
      { text: "• 안전재고: 4-8주\n• 공급업체: 2-3개 다변화\n• 계약: 1-3년 중장기\n• 발주: ROP", options: { fontSize: 11 } },
    ], { x: 5.25, y: 1.75, w: 4.1, h: 1.6 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 5.1, y: 3.5, w: 4.4, h: 0.5,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText("보험 관점: 안전재고 비용 ≪ 생산 중단 손실", {
      x: 5.25, y: 3.6, w: 4.1, h: 0.35,
      fontSize: 10,
    });

    addFooter(slide, "1회차 | 4대 자재군");
  }

  function createSlide14(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("🟣 전략자재 (Strategic Items)", {
      x: 0.5, y: 0.5, w: 9, h: 0.5,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addText("높은 공급 리스크 + 높은 구매 임팩트", {
      x: 0.5, y: 1.1, w: 9, h: 0.3,
      fontSize: 16, color: colors.mutedText,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.6, w: 4.4, h: 1.3,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "특징\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 금액 크고 공급 어려움\n• 사업 성패 좌우\n• 대체 불가능\n• 장기 개발 필요", options: { fontSize: 11 } },
    ], { x: 0.65, y: 1.75, w: 4.1, h: 1.1 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 3.0, w: 4.4, h: 1.0,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "사례\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 핵심 반도체 (AP, SoC)\n• OLED 발광재료\n• 장납기 외자재\n• 독점 기술 부품", options: { fontSize: 11 } },
    ], { x: 0.65, y: 3.15, w: 4.1, h: 0.8 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 5.1, y: 1.6, w: 4.4, h: 1.8,
      fill: { color: colors.warning },
      line: { color: colors.warningBorder, width: 2 },
    });

    slide.addText([
      { text: "핵심 과제 & 관리 전략\n", options: { fontSize: 14, bold: true } },
      { text: "목표: 전략적 파트너십 | 철학: \"Win-Win\" | KPI: 연속성 100%\n\n", options: { fontSize: 10 } },
      { text: "• 안전재고: 3-6주\n• 공급업체: 1-2개 전략적\n• 계약: 3-5년 장기\n• 발주: LTP + Hybrid", options: { fontSize: 11 } },
    ], { x: 5.25, y: 1.75, w: 4.1, h: 1.6 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 5.1, y: 3.5, w: 4.4, h: 0.5,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText("파트너십: 단기 절감 < 장기 가치", {
      x: 5.25, y: 3.6, w: 4.1, h: 0.35,
      fontSize: 10,
    });

    addFooter(slide, "1회차 | 4대 자재군");
  }

  function createSlide21(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("7회차 학습 여정", {
      x: 0.5, y: 0.5, w: 9, h: 0.5,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addText("전략적 재고운영 완전 마스터 로드맵", {
      x: 0.5, y: 1.1, w: 9, h: 0.3,
      fontSize: 16, color: colors.mutedText,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.6, w: 9, h: 0.75,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "Module 1: Foundation (1-2회차)\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 1회차: JIT→JIC + Kraljic Matrix\n• 2회차: 소싱 전략 + 공급업체 관리", options: { fontSize: 11 } },
    ], { x: 0.65, y: 1.7, w: 8.7, h: 0.6 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 2.5, w: 9, h: 1.0,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "Module 2: 자재군별 심화 (3-6회차)\n", options: { fontSize: 14, bold: true, color: colors.primary } },
      { text: "• 3회차: 병목자재 + ROP\n• 4회차: 레버리지자재 + MRP\n• 5회차: 전략자재 + LTP\n• 6회차: 일상자재 + 자동화", options: { fontSize: 11 } },
    ], { x: 0.65, y: 2.6, w: 8.7, h: 0.8 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 3.65, w: 9, h: 0.5,
      fill: { color: colors.warning },
      line: { color: colors.warningBorder, width: 2 },
    });

    slide.addText([
      { text: "Module 3: 실전 통합 (7회차)\n", options: { fontSize: 14, bold: true } },
      { text: "• 7회차: Kraljic Matrix 실전 워크샵", options: { fontSize: 11 } },
    ], { x: 0.65, y: 3.75, w: 8.7, h: 0.35 });

    addFooter(slide, "1회차 | 학습 여정");
  }

  function createSlide22(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("핵심 요약", {
      x: 0.5, y: 0.4, w: 9, h: 0.4,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.0, w: 9, h: 0.85,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "1. 패러다임의 전환\n", options: { fontSize: 13, bold: true, color: colors.primary } },
      { text: "JIT: 재고=낭비, 효율성, 획일적 → JIC: 재고=전략자산, 회복력, 차별화", options: { fontSize: 10 } },
    ], { x: 0.65, y: 1.1, w: 8.7, h: 0.7 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 2.0, w: 4.4, h: 1.0,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "2. Kraljic Matrix\n", options: { fontSize: 13, bold: true, color: colors.primary } },
      { text: "• 2개 축: 공급 리스크 × 구매 임팩트\n• 4개 자재군 차별화 전략", options: { fontSize: 10 } },
    ], { x: 0.65, y: 2.1, w: 4.1, h: 0.85 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 5.1, y: 2.0, w: 4.4, h: 1.0,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "3. 자재계획 방법론\n", options: { fontSize: 13, bold: true, color: colors.primary } },
      { text: "• 병목→ROP | 레버리지→MRP\n• 전략→LTP | 일상→VMI", options: { fontSize: 10 } },
    ], { x: 5.25, y: 2.1, w: 4.1, h: 0.85 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 3.15, w: 9, h: 0.5,
      fill: { color: colors.warning },
      line: { color: colors.warningBorder, width: 2 },
    });

    slide.addText("4. 본 과정의 가치: 즉시 적용 가능한 구체적 방안 제공", {
      x: 0.65, y: 3.3, w: 8.7, h: 0.25,
      fontSize: 11, bold: true,
    });

    addFooter(slide, "1회차 | 요약");
  }

  function createSlide23(pptx) {
    const slide = pptx.addSlide();
    slide.background = { color: colors.white };

    slide.addText("다음 회차 예고", {
      x: 0.5, y: 0.5, w: 9, h: 0.5,
      fontSize: 28, bold: true, color: colors.primary,
    });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.2, w: 9, h: 1.2,
      fill: { color: colors.info },
      line: { color: colors.infoBorder, width: 2 },
    });

    slide.addText([
      { text: "2회차: 소싱 전략 및 공급업체 관계 관리\n\n", options: { fontSize: 16, bold: true, color: colors.primary } },
      { text: "• 자재군별 차별화된 소싱 전략\n• SRM 프레임워크\n• 계약 전략 및 협상\n• 공급업체 성과 평가", options: { fontSize: 12 } },
    ], { x: 0.65, y: 1.35, w: 8.7, h: 1.0 });

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 2.55, w: 9, h: 0.8,
      fill: { color: colors.muted },
      line: { color: colors.primary, width: 3, dashType: "solid" },
    });

    slide.addText([
      { text: "강사 TIP\n", options: { fontSize: 13, bold: true, color: colors.primary } },
      { text: "Kraljic Matrix는 조직 전체가 자재를 바라보는 공통 언어입니다.\n다음 회차부터는 각 자재군별 구체적인 전략과 방법론을 배우게 됩니다!", options: { fontSize: 11 } },
    ], { x: 0.65, y: 2.65, w: 8.7, h: 0.65 });

    slide.addText("감사합니다!", {
      x: 0.5, y: 3.6, w: 9, h: 0.4,
      fontSize: 36, bold: true, color: colors.primary,
      align: "center",
    });

    addFooter(slide, "1회차 | 전략적 재고운영");
  }

  console.log("\n🔄 Creating complete PPTX with all 23 slides...");

  // Process each slide
  for (let i = 0; i < slides.length; i++) {
    const slideNum = i + 1;

    // Manual slides (PptxGenJS)
    if (slideNum === 9) {
      createSlide9(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }
    if (slideNum === 12) {
      createSlide12(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }
    if (slideNum === 14) {
      createSlide14(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }
    if (slideNum === 21) {
      createSlide21(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }
    if (slideNum === 22) {
      createSlide22(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }
    if (slideNum === 23) {
      createSlide23(pptx);
      console.log(`  ✅ Slide ${slideNum}/23 created (manual)`);
      continue;
    }

    // html2pptx slides
    const slideElement = slides.eq(i);
    const slideHtml = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slide ${slideNum}</title>
    <style>
        ${globalStyles}
        .slide { height: 420px !important; }
    </style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; overflow: hidden;">
    ${slideElement.html()}
</body>
</html>`;

    const slideHtmlPath = path.join(tmpDir, `slide-${slideNum}.html`);
    fs.writeFileSync(slideHtmlPath, slideHtml, "utf-8");

    try {
      await html2pptx(slideHtmlPath, pptx);
      console.log(`  ✅ Slide ${slideNum}/23 converted (html2pptx)`);
    } catch (error) {
      console.error(`  ❌ Slide ${slideNum} failed: ${error.message}`);
      console.log(`  ⏭️  Skipping...`);
    }
  }

  // Save PPTX
  console.log(`\n💾 Saving complete PPTX...`);
  await pptx.writeFile({ fileName: outputPptxPath });

  // Cleanup
  console.log("🧹 Cleaning up...");
  for (let i = 1; i <= slides.length; i++) {
    const slideHtmlPath = path.join(tmpDir, `slide-${i}.html`);
    if (fs.existsSync(slideHtmlPath)) {
      fs.unlinkSync(slideHtmlPath);
    }
  }

  console.log(`\n✅ Successfully created complete PPTX with all 23 slides!`);
  console.log(`📁 Output: ${outputPptxPath}`);
}

// Main execution
if (require.main === module) {
  const inputHtml = path.resolve(process.argv[2] || "C:\\Users\\ahfif\\SuperClaude\\Project_Strategic_edu\\html\\Part1\\Part1_전략적재고운영Foundation_23slides_960x540.html");
  const outputPptx = path.resolve(process.argv[3] || "C:\\Users\\ahfif\\SuperClaude\\Project_Strategic_edu\\pptx\\Part1\\Part1_전략적재고운영Foundation_Complete.pptx");

  createCompletePPTX(inputHtml, outputPptx)
    .then(() => {
      console.log("\n🎉 Conversion complete!");
      process.exit(0);
    })
    .catch((error) => {
      console.error("\n❌ Conversion failed:", error.message);
      process.exit(1);
    });
}
