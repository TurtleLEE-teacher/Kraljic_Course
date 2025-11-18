#!/usr/bin/env node

/**
 * Split HTML Slides
 * Extracts individual slides from a single HTML file containing multiple .slide divs
 */

const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length < 1) {
  console.error('Usage: node split-html-slides.cjs <html-file> [output-dir]');
  console.error('Example: node split-html-slides.cjs Part2.html temp/slides');
  process.exit(1);
}

const htmlFilePath = path.resolve(args[0]);
const outputDir = args[1]
  ? path.resolve(args[1])
  : path.join(process.cwd(), 'temp', 'slides');

console.log('🔪 HTML Slide Splitter');
console.log('=====================');
console.log(`Input: ${htmlFilePath}`);
console.log(`Output Dir: ${outputDir}`);

// Check if input file exists
if (!fs.existsSync(htmlFilePath)) {
  console.error(`❌ Error: Input file not found: ${htmlFilePath}`);
  process.exit(1);
}

// Create output directory
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`✅ Created output directory: ${outputDir}`);
}

// Read HTML file
console.log('\n📖 Reading HTML file...');
const htmlContent = fs.readFileSync(htmlFilePath, 'utf-8');

// Parse HTML
console.log('🔍 Parsing HTML structure...');
const $ = cheerio.load(htmlContent);

// Extract head content (CSS, links, etc.)
const headContent = $('head').html();
const styleContent = $('style').html() || '';

// Find all slides
const slides = $('.slide');
console.log(`✅ Found ${slides.length} slides\n`);

// Split each slide into separate HTML file
console.log('📄 Splitting slides...');
const outputFiles = [];

slides.each((index, element) => {
  const slideNum = index + 1;
  const slideHtml = $(element).toString();

  // Create complete HTML document for this slide
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

  // Generate filename with zero-padded number
  const filename = `slide-${String(slideNum).padStart(3, '0')}.html`;
  const filepath = path.join(outputDir, filename);

  // Write file
  fs.writeFileSync(filepath, completeHtml, 'utf-8');
  outputFiles.push(filepath);

  // Progress log every 5 slides
  if (slideNum % 5 === 0 || slideNum === slides.length) {
    console.log(`  ✓ Split ${slideNum}/${slides.length} slides`);
  }
});

console.log(`\n✅ Successfully split ${slides.length} slides!`);
console.log(`📁 Output: ${outputDir}`);
console.log(`📄 Files: ${outputFiles[0]} ... ${outputFiles[outputFiles.length - 1]}`);

// Return output files list as JSON for next step
const outputListFile = path.join(outputDir, '_file-list.json');
fs.writeFileSync(
  outputListFile,
  JSON.stringify({ files: outputFiles, count: outputFiles.length }, null, 2),
  'utf-8'
);
console.log(`\n📋 File list saved: ${outputListFile}`);
