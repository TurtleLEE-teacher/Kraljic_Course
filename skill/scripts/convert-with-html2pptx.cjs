#!/usr/bin/env node

/**
 * Convert HTML to PPTX using html2pptx
 * Complete workflow: Split → Convert → Merge
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length < 1) {
  console.error('Usage: node convert-with-html2pptx.cjs <html-file> [output-pptx]');
  console.error('Example: node convert-with-html2pptx.cjs Part2.html output/Part2.pptx');
  process.exit(1);
}

const htmlFilePath = path.resolve(args[0]);
const outputFile = args[1]
  ? path.resolve(args[1])
  : path.join(process.cwd(), 'output', path.basename(htmlFilePath, '.html') + '_html2pptx.pptx');

console.log('🚀 HTML to PPTX Converter (html2pptx)');
console.log('=====================================');
console.log(`Input: ${htmlFilePath}`);
console.log(`Output: ${outputFile}\n`);

// Check if input file exists
if (!fs.existsSync(htmlFilePath)) {
  console.error(`❌ Error: Input file not found: ${htmlFilePath}`);
  process.exit(1);
}

// Create temp directories
const tempDir = path.join(process.cwd(), 'temp');
const slidesDir = path.join(tempDir, 'slides');
const pptxDir = path.join(tempDir, 'pptx');

[tempDir, slidesDir, pptxDir].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

try {
  // Step 1: Split HTML into individual slides
  console.log('📝 STEP 1: Splitting HTML slides...');
  console.log('─'.repeat(50));

  const splitScript = path.join(__dirname, 'split-html-slides.cjs');
  execSync(`node "${splitScript}" "${htmlFilePath}" "${slidesDir}"`, {
    stdio: 'inherit'
  });

  // Read file list
  const fileListPath = path.join(slidesDir, '_file-list.json');
  const fileList = JSON.parse(fs.readFileSync(fileListPath, 'utf-8'));
  const htmlFiles = fileList.files;

  console.log(`\n✅ Step 1 complete: ${htmlFiles.length} HTML files created\n`);

  // Step 2: Convert each HTML to PPTX using html2pptx
  console.log('🔄 STEP 2: Converting HTML slides to PPTX...');
  console.log('─'.repeat(50));

  const pptxFiles = [];

  for (let i = 0; i < htmlFiles.length; i++) {
    const htmlFile = htmlFiles[i];
    const slideNum = i + 1;
    const pptxFile = path.join(pptxDir, `slide-${String(slideNum).padStart(3, '0')}.pptx`);

    console.log(`  Converting ${slideNum}/${htmlFiles.length}: ${path.basename(htmlFile)}`);

    try {
      // Run html2pptx
      execSync(
        `html2pptx "${htmlFile}" "${pptxFile}" --width 960 --height 540`,
        { stdio: 'pipe' }
      );

      pptxFiles.push(pptxFile);

      // Progress indicator
      if (slideNum % 5 === 0 || slideNum === htmlFiles.length) {
        console.log(`    ✓ Converted ${slideNum}/${htmlFiles.length} slides`);
      }

    } catch (error) {
      console.error(`    ❌ Error converting slide ${slideNum}:`, error.message);
      // Continue with next slide
    }
  }

  console.log(`\n✅ Step 2 complete: ${pptxFiles.length} PPTX slides created\n`);

  // Step 3: Merge all PPTX files
  console.log('🔗 STEP 3: Merging PPTX slides...');
  console.log('─'.repeat(50));

  // Create output directory
  const outputDir = path.dirname(outputFile);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Run Python merge script
  const mergePyScript = path.join(__dirname, 'merge-pptx.py');
  execSync(
    `python "${mergePyScript}" "${outputFile}" "${pptxDir}"`,
    { stdio: 'inherit' }
  );

  console.log(`\n✅ Step 3 complete: Merged PPTX created\n`);

  // Final summary
  console.log('=' * 50);
  console.log('✅ CONVERSION COMPLETE!');
  console.log('=' * 50);
  console.log(`📊 Output: ${outputFile}`);

  // Get file stats
  const stats = fs.statSync(outputFile);
  const fileSizeMB = (stats.size / (1024 * 1024)).toFixed(2);

  console.log(`📈 Total slides: ${pptxFiles.length}`);
  console.log(`💾 File size: ${fileSizeMB} MB`);

  console.log('\n💡 Next steps:');
  console.log(`   1. Open: ${outputFile}`);
  console.log(`   2. Verify layout accuracy`);
  console.log(`   3. Compare with original HTML`);

  // Cleanup option
  console.log('\n🧹 Temporary files:');
  console.log(`   HTML slides: ${slidesDir}`);
  console.log(`   PPTX slides: ${pptxDir}`);
  console.log(`   (These can be deleted after verification)`);

} catch (error) {
  console.error('\n❌ Conversion failed:', error.message);
  console.error('\nDebug info:');
  console.error(`  Temp dir: ${tempDir}`);
  console.error(`  Check logs above for specific errors`);
  process.exit(1);
}
