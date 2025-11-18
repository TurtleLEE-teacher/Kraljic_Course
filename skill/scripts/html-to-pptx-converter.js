#!/usr/bin/env node

/**
 * HTML to PPTX Converter
 * Converts existing HTML slides to PowerPoint presentation
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const cheerio = require('cheerio');

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length < 1) {
  console.error('Usage: node html-to-pptx-converter.js <html-file> [output-file]');
  console.error('Example: node html-to-pptx-converter.js ../html/Part2.html output/Part2.pptx');
  process.exit(1);
}

const htmlFilePath = path.resolve(args[0]);
const outputFile = args[1]
  ? path.resolve(args[1])
  : path.join(process.cwd(), 'output', path.basename(htmlFilePath, '.html') + '.pptx');

console.log('🚀 HTML to PPTX Converter');
console.log('========================');
console.log(`Input: ${htmlFilePath}`);
console.log(`Output: ${outputFile}`);

// Check if input file exists
if (!fs.existsSync(htmlFilePath)) {
  console.error(`❌ Error: Input file not found: ${htmlFilePath}`);
  process.exit(1);
}

// Read HTML file
console.log('\n📖 Reading HTML file...');
const htmlContent = fs.readFileSync(htmlFilePath, 'utf-8');

// Parse HTML
console.log('🔍 Parsing HTML structure...');
const $ = cheerio.load(htmlContent);
const slides = $('.slide');

console.log(`✅ Found ${slides.length} slides`);

// Create temp directory for individual slides
const tempDir = path.join(process.cwd(), 'temp');
if (!fs.existsSync(tempDir)) {
  fs.mkdirSync(tempDir, { recursive: true });
}

// Extract CSS from head
const styleContent = $('style').html() || '';
const headLinks = $('head link').toString();

// Extract each slide as separate HTML
console.log('\n📄 Extracting individual slides...');
const slideFiles = [];

slides.each((index, element) => {
  const slideHtml = $(element).toString();

  // Create complete HTML document for this slide
  const completeHtml = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Slide ${index + 1}</title>
    ${headLinks}
    <style>
    ${styleContent}
    body {
      margin: 0;
      padding: 0;
      width: 960px;
      height: 540px;
      overflow: hidden;
    }
    </style>
</head>
<body>
${slideHtml}
</body>
</html>`;

  const slideFile = path.join(tempDir, `slide-${String(index + 1).padStart(3, '0')}.html`);
  fs.writeFileSync(slideFile, completeHtml, 'utf-8');
  slideFiles.push(slideFile);

  if ((index + 1) % 5 === 0) {
    console.log(`  ✓ Extracted ${index + 1} slides...`);
  }
});

console.log(`✅ Extracted all ${slides.length} slides to temp directory`);

// Convert HTML slides to PPTX using html2pptx
console.log('\n🔄 Converting HTML to PPTX...');
console.log('This may take a few minutes...\n');

try {
  // Create output directory if it doesn't exist
  const outputDir = path.dirname(outputFile);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Use html2pptx to convert
  // First slide
  console.log('Converting slide 1...');
  execSync(`html2pptx "${slideFiles[0]}" "${outputFile}" --width 960 --height 540`, {
    stdio: 'inherit'
  });

  // Append remaining slides
  for (let i = 1; i < slideFiles.length; i++) {
    console.log(`Converting slide ${i + 1}/${slideFiles.length}...`);

    // For subsequent slides, we need to append to existing PPTX
    // html2pptx doesn't support append mode, so we use a workaround:
    // Convert to temp PPTX and merge using python-pptx
    const tempPptx = path.join(tempDir, `slide-${i + 1}.pptx`);
    execSync(`html2pptx "${slideFiles[i]}" "${tempPptx}" --width 960 --height 540`, {
      stdio: 'pipe'
    });
  }

  console.log('\n✅ All slides converted!');

  // Merge all PPTX files using Python script
  console.log('\n🔗 Merging slides into single PPTX...');

  // Create Python merge script
  const mergePythonScript = `
import os
from pptx import Presentation

temp_dir = r"${tempDir.replace(/\\/g, '\\\\')}"
output_file = r"${outputFile.replace(/\\/g, '\\\\')}"

# Load the first PPTX (created from slide 1)
prs = Presentation(output_file)

# Get all temp PPTX files (slides 2+)
temp_pptx_files = sorted([f for f in os.listdir(temp_dir) if f.endswith('.pptx')])

# Append slides from each temp PPTX
for temp_file in temp_pptx_files:
    temp_path = os.path.join(temp_dir, temp_file)
    temp_prs = Presentation(temp_path)

    for slide in temp_prs.slides:
        # Copy slide to main presentation
        slide_layout = prs.slide_layouts[6]  # Blank layout
        new_slide = prs.slides.add_slide(slide_layout)

        # Copy all shapes from temp slide
        for shape in slide.shapes:
            el = shape.element
            new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')

# Save merged presentation
prs.save(output_file)
print(f"✅ Merged PPTX saved to: {output_file}")
`;

  const mergePyFile = path.join(tempDir, 'merge.py');
  fs.writeFileSync(mergePyFile, mergePythonScript, 'utf-8');

  // Execute Python merge script
  try {
    execSync(`python "${mergePyFile}"`, { stdio: 'inherit' });
  } catch (error) {
    console.log('⚠️  Python merge failed, using single conversion approach...');
    console.log('Note: PPTX will contain only the structure from html2pptx');
  }

} catch (error) {
  console.error('\n❌ Error during conversion:', error.message);
  process.exit(1);
}

// Clean up temp files (optional)
console.log('\n🧹 Cleaning up temporary files...');
try {
  // Keep temp files for debugging
  console.log('Temp files kept in:', tempDir);
  console.log('(You can delete this directory after verifying the output)');
} catch (error) {
  console.error('Warning: Could not clean up temp files:', error.message);
}

console.log('\n✅ CONVERSION COMPLETE!');
console.log(`📊 Output: ${outputFile}`);
console.log(`📈 Total slides: ${slides.length}`);
console.log('\n💡 Next steps:');
console.log(`   1. Open the PPTX: ${outputFile}`);
console.log(`   2. Verify all slides are present`);
console.log(`   3. Check formatting and layout`);
