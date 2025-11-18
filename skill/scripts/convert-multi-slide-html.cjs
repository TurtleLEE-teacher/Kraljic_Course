const pptxgen = require("pptxgenjs");
const { html2pptx } = require("@ant/html2pptx");
const fs = require("fs");
const path = require("path");
const cheerio = require("cheerio");

/**
 * Convert a multi-slide HTML file to PPTX
 * Extracts individual slides and converts them to PowerPoint
 *
 * Usage:
 *   NODE_PATH="$(npm root -g)" node convert-multi-slide-html.js <input.html> <output.pptx>
 */

async function convertMultiSlideHTML(inputHtmlPath, outputPptxPath) {
  console.log(`📖 Reading HTML file: ${inputHtmlPath}`);

  // Read the HTML file
  const htmlContent = fs.readFileSync(inputHtmlPath, "utf-8");

  // Parse with cheerio
  const $ = cheerio.load(htmlContent);

  // Find all slides
  const slides = $(".slide");
  console.log(`✅ Found ${slides.length} slides`);

  if (slides.length === 0) {
    throw new Error("No slides found with class 'slide'");
  }

  // Create temporary directory for individual slide HTML files
  const tmpDir = path.join(__dirname, "../temp/slides");
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  // Extract global styles from the original HTML
  const globalStyles = $("style").html() || "";

  // Create PPTX presentation
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.author = "Claude Code";
  pptx.title = "전략적 재고운영 및 자재계획 수립";

  console.log("\n🔄 Converting slides to PPTX...");

  // Process each slide
  for (let i = 0; i < slides.length; i++) {
    const slideElement = slides.eq(i);
    const slideNum = i + 1;

    // Create individual HTML file for this slide
    // html2pptx requires 0.5" bottom margin, so adjust slide height
    const slideHtml = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slide ${slideNum}</title>
    <style>
        ${globalStyles}
        /* Override slide height for html2pptx constraints */
        .slide {
            height: 492px !important;
        }
    </style>
</head>
<body style="width: 960px; height: 540px; margin: 0; padding: 0; overflow: hidden;">
    ${slideElement.html()}
</body>
</html>`;

    // Save to temp file
    const slideHtmlPath = path.join(tmpDir, `slide-${slideNum}.html`);
    fs.writeFileSync(slideHtmlPath, slideHtml, "utf-8");

    try {
      // Convert to PPTX slide
      const { slide, placeholders } = await html2pptx(slideHtmlPath, pptx);
      console.log(`  ✅ Slide ${slideNum}/${slides.length} converted`);
    } catch (error) {
      console.error(`  ❌ Slide ${slideNum} failed: ${error.message}`);
      console.log(`  ⏭️  Skipping slide ${slideNum} and continuing...`);
      // Don't throw - continue with next slide
    }
  }

  // Save PPTX
  console.log(`\n💾 Saving PPTX to: ${outputPptxPath}`);
  await pptx.writeFile({ fileName: outputPptxPath });

  // Cleanup temp files (optional)
  console.log("\n🧹 Cleaning up temporary files...");
  for (let i = 1; i <= slides.length; i++) {
    const slideHtmlPath = path.join(tmpDir, `slide-${i}.html`);
    if (fs.existsSync(slideHtmlPath)) {
      fs.unlinkSync(slideHtmlPath);
    }
  }

  console.log(`\n✅ Successfully created PPTX with ${slides.length} slides!`);
  console.log(`📁 Output: ${outputPptxPath}`);
}

// Main execution
if (require.main === module) {
  const args = process.argv.slice(2);

  if (args.length < 2) {
    console.error("Usage: NODE_PATH=\"$(npm root -g)\" node convert-multi-slide-html.js <input.html> <output.pptx>");
    process.exit(1);
  }

  const inputHtml = path.resolve(args[0]);
  const outputPptx = path.resolve(args[1]);

  if (!fs.existsSync(inputHtml)) {
    console.error(`Error: Input file not found: ${inputHtml}`);
    process.exit(1);
  }

  convertMultiSlideHTML(inputHtml, outputPptx)
    .then(() => {
      console.log("\n🎉 Conversion complete!");
      process.exit(0);
    })
    .catch((error) => {
      console.error("\n❌ Conversion failed:", error.message);
      console.error(error.stack);
      process.exit(1);
    });
}

module.exports = { convertMultiSlideHTML };
