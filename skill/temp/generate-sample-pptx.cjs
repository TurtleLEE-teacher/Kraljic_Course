const pptxgen = require("pptxgenjs");
const { html2pptx } = require("@ant/html2pptx");
const path = require("path");

async function createSamplePresentation() {
  console.log("Starting PPTX generation...");

  // Create presentation
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9"; // 960x540
  pptx.author = "Claude Code - pptx-mslee";
  pptx.title = "재고관리 샘플 장표";
  pptx.subject = "ABC-XYZ 매트릭스 및 E2E 프로세스 비교";

  try {
    // Slide 1: ABC-XYZ Matrix
    console.log("Converting Slide 1: ABC-XYZ 재고관리 매트릭스...");
    const slide1Path = path.join(__dirname, "slide1-abc-xyz-matrix.html");
    const { slide: slide1 } = await html2pptx(slide1Path, pptx);
    console.log("✓ Slide 1 converted successfully");

    // Slide 2: E2E Process Comparison
    console.log("Converting Slide 2: 제조업 vs 유통업 E2E 프로세스...");
    const slide2Path = path.join(__dirname, "slide2-e2e-comparison.html");
    const { slide: slide2 } = await html2pptx(slide2Path, pptx);
    console.log("✓ Slide 2 converted successfully");

    // Save presentation
    const outputPath = path.join(__dirname, "..", "output", "재고관리_샘플장표_2slides.pptx");
    await pptx.writeFile({ fileName: outputPath });
    console.log(`\n✅ Presentation created successfully!`);
    console.log(`📁 Output: ${outputPath}`);
    console.log(`📊 Total slides: 2`);

  } catch (error) {
    console.error("❌ Error creating presentation:", error.message);
    if (error.stack) {
      console.error(error.stack);
    }
    process.exit(1);
  }
}

// Run
createSamplePresentation().catch(console.error);
