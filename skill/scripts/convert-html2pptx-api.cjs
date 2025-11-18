#!/usr/bin/env node

/**
 * Convert HTML to PPTX using html2pptx Node.js API
 * Complete workflow: Split → Convert using API → Merge
 */

const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// Import html2pptx (CommonJS)
const html2pptxPath = path.join(__dirname, '../node_modules/@ant/html2pptx/dist/html2pptx.cjs');
const { html2pptx } = require(html2pptxPath);

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length < 1) {
  console.error('Usage: node convert-html2pptx-api.cjs <html-file> [output-pptx]');
  console.error('Example: node convert-html2pptx-api.cjs Part2.html output/Part2.pptx');
  process.exit(1);
}

const htmlFilePath = path.resolve(args[0]);
const outputFile = args[1]
  ? path.resolve(args[1])
  : path.join(process.cwd(), 'output', path.basename(htmlFilePath, '.html') + '_html2pptx.pptx');

console.log('🚀 HTML to PPTX Converter (html2pptx API)');
console.log('=========================================');
console.log(`Input: ${htmlFilePath}`);
console.log(`Output: ${outputFile}\n`);

// Check if input file exists
if (!fs.existsSync(htmlFilePath)) {
  console.error(`❌ Error: Input file not found: ${htmlFilePath}`);
  process.exit(1);
}

async function convertHTML() {
  try {
    // Create temp directories
    const tempDir = path.join(process.cwd(), 'temp');
    const slidesDir = path.join(tempDir, 'slides');

    [tempDir, slidesDir].forEach(dir => {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    });

    // Step 1: Split HTML into individual slides
    console.log('📝 STEP 1: Splitting HTML slides...');
    console.log('─'.repeat(50));

    const cheerio = require('cheerio');
    const htmlContent = fs.readFileSync(htmlFilePath, 'utf-8');
    const $ = cheerio.load(htmlContent);

    // Extract head content
    const styleContent = $('style').html() || '';
    const slides = $('.slide');

    console.log(`✅ Found ${slides.length} slides\n`);

    const htmlFiles = [];

    // Create individual HTML files
    slides.each((index, element) => {
      const slideNum = index + 1;
      const slideHtml = $(element).toString();

      const completeHtml = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slide ${slideNum}</title>
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        ${styleContent}
        body {
            margin: 0;
            padding: 0;
            width: 960px;
            height: 540px;
            overflow: hidden;
        }
        .slide {
            margin: 0 !important;
        }
    </style>
</head>
<body>
${slideHtml}
</body>
</html>`;

      const filename = `slide-${String(slideNum).padStart(3, '0')}.html`;
      const filepath = path.join(slidesDir, filename);

      fs.writeFileSync(filepath, completeHtml, 'utf-8');
      htmlFiles.push(filepath);

      if (slideNum % 5 === 0 || slideNum === slides.length) {
        console.log(`  ✓ Split ${slideNum}/${slides.length} slides`);
      }
    });

    console.log(`\n✅ Step 1 complete: ${htmlFiles.length} HTML files created\n`);

    // Step 2: Convert each HTML to PPTX slides using html2pptx API
    console.log('🔄 STEP 2: Converting HTML to PPTX using API...');
    console.log('─'.repeat(50));

    // Create PowerPoint presentation
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'Strategic Inventory Management Course';
    pptx.company = 'Education Materials';
    pptx.title = 'Part 2. 소싱 전략 및 유형별 재고관리 프로세스';

    // Convert each HTML file to a slide
    for (let i = 0; i < htmlFiles.length; i++) {
      const htmlFile = htmlFiles[i];
      const slideNum = i + 1;

      console.log(`  Converting ${slideNum}/${htmlFiles.length}: ${path.basename(htmlFile)}`);

      try {
        // Use html2pptx API to add slide
        const result = await html2pptx(htmlFile, pptx, {
          // Options if needed
        });

        if (slideNum % 5 === 0 || slideNum === htmlFiles.length) {
          console.log(`    ✓ Converted ${slideNum}/${htmlFiles.length} slides`);
        }

      } catch (error) {
        console.error(`    ⚠️ Error converting slide ${slideNum}:`, error.message);
        // Continue with next slide
      }
    }

    console.log(`\n✅ Step 2 complete: All slides converted\n`);

    // Step 3: Save PPTX
    console.log('💾 STEP 3: Saving PowerPoint file...');
    console.log('─'.repeat(50));

    // Create output directory
    const outputDir = path.dirname(outputFile);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Write PPTX file
    await pptx.writeFile({ fileName: outputFile });

    // Get file stats
    const stats = fs.statSync(outputFile);
    const fileSizeMB = (stats.size / (1024 * 1024)).toFixed(2);

    console.log('\n' + '='.repeat(50));
    console.log('✅ CONVERSION COMPLETE!');
    console.log('='.repeat(50));
    console.log(`📊 Output: ${outputFile}`);
    console.log(`📈 Total slides: ${htmlFiles.length}`);
    console.log(`💾 File size: ${fileSizeMB} MB`);

    console.log('\n💡 Next steps:');
    console.log(`   1. Open: ${outputFile}`);
    console.log(`   2. Verify layout accuracy`);
    console.log(`   3. Compare with original HTML`);

    console.log('\n🧹 Temporary files:');
    console.log(`   HTML slides: ${slidesDir}`);
    console.log(`   (Can be deleted after verification)`);

  } catch (error) {
    console.error('\n❌ Conversion failed:', error);
    console.error('\nStack trace:');
    console.error(error.stack);
    process.exit(1);
  }
}

// Run conversion
convertHTML();
