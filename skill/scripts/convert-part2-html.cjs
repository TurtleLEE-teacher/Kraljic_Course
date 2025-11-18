#!/usr/bin/env node

/**
 * Part2 HTML to PPTX Converter
 * Converts Part2 HTML slides to PowerPoint using PptxGenJS
 */

const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');
const PptxGenJS = require('pptxgenjs');

// Parse command line arguments
const args = process.argv.slice(2);
if (args.length < 1) {
  console.error('Usage: node convert-part2-html.js <html-file> [output-file]');
  console.error('Example: node convert-part2-html.js ../../Project_Strategic_edu/html/Part2/Part2_소싱전략_25slides_960x540.html');
  process.exit(1);
}

const htmlFilePath = path.resolve(args[0]);
const outputFile = args[1]
  ? path.resolve(args[1])
  : path.join(process.cwd(), 'output', 'Part2_소싱전략.pptx');

console.log('🚀 Part2 HTML to PPTX Converter');
console.log('================================');
console.log(`Input: ${htmlFilePath}`);
console.log(`Output: ${outputFile}`);

// Check if input file exists
if (!fs.existsSync(htmlFilePath)) {
  console.error(`❌ Error: Input file not found: ${htmlFilePath}`);
  process.exit(1);
}

// Create output directory if it doesn't exist
const outputDir = path.dirname(outputFile);
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// Read HTML file
console.log('\n📖 Reading HTML file...');
const htmlContent = fs.readFileSync(htmlFilePath, 'utf-8');

// Parse HTML
console.log('🔍 Parsing HTML structure...');
const $ = cheerio.load(htmlContent);
const slides = $('.slide');

console.log(`✅ Found ${slides.length} slides`);

// Create PowerPoint
console.log('\n🎨 Creating PowerPoint presentation...');
const pptx = new PptxGenJS();

// Set presentation properties
pptx.layout = 'LAYOUT_16x9';
pptx.author = 'Strategic Inventory Management Course';
pptx.company = 'Education Materials';
pptx.title = 'Part 2. 소싱 전략 및 유형별 재고관리 프로세스';

// Define color scheme (from HTML)
const colors = {
  primaryDark: '00546f',
  primaryLight: '008db9',
  textDark: '23282a',
  textGray: '6c757d',
  bgGray: 'f8fafc',
  white: 'FFFFFF',
  redBg: 'fee',
  redText: 'dc3545',
  greenBg: 'eff9f0',
  greenText: '28a745',
  purpleBg: 'f3e5f5',
  purpleText: '6f42c1',
  orangeText: 'e67e22',
  grayBg: 'f5f5f5'
};

// Helper function to extract text content
function extractText(element) {
  return $(element).text().trim();
}

// Helper function to add title with accent bar
function addSectionTitle(slide, text, yPos = 0.5) {
  // Title text
  slide.addText(text, {
    x: 0.5,
    y: yPos,
    w: 9,
    h: 0.4,
    fontSize: 18,
    bold: true,
    color: colors.primaryLight,
    fontFace: 'Noto Sans KR'
  });

  // Accent bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.5,
    y: yPos + 0.45,
    w: 0.6,
    h: 0.04,
    fill: { type: 'solid', color: colors.primaryLight }
  });
}

// Helper function to add feature card
function addFeatureCard(slide, x, y, w, h, iconText, title, content) {
  // Card background
  slide.addShape(pptx.ShapeType.rect, {
    x: x,
    y: y,
    w: w,
    h: h,
    fill: { type: 'solid', color: colors.white },
    line: { type: 'solid', color: 'E5E7EB', width: 1 }
  });

  // Icon circle
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.1,
    y: y + 0.1,
    w: 0.35,
    h: 0.35,
    fill: { type: 'solid', color: colors.primaryLight }
  });

  // Icon text
  slide.addText(iconText, {
    x: x + 0.1,
    y: y + 0.1,
    w: 0.35,
    h: 0.35,
    fontSize: 14,
    color: colors.white,
    align: 'center',
    valign: 'middle',
    fontFace: 'Noto Sans KR'
  });

  // Title
  slide.addText(title, {
    x: x + 0.1,
    y: y + 0.5,
    w: w - 0.2,
    h: 0.3,
    fontSize: 12,
    bold: true,
    color: colors.textDark,
    fontFace: 'Noto Sans KR'
  });

  // Content
  slide.addText(content, {
    x: x + 0.1,
    y: y + 0.85,
    w: w - 0.2,
    h: h - 1,
    fontSize: 10,
    color: colors.textGray,
    fontFace: 'Noto Sans KR'
  });
}

// Process each slide
console.log('\n📄 Converting slides...');
let slideCount = 0;

slides.each((index, element) => {
  const slideEl = $(element);
  slideCount++;

  console.log(`  Converting slide ${slideCount}/${slides.length}...`);

  const slide = pptx.addSlide();

  // Set background
  slide.background = { color: colors.white };

  // Check if it's the cover slide
  const isCover = slideEl.find('h1.title-text').length > 0;

  if (isCover) {
    // Cover slide
    const mainTitle = extractText(slideEl.find('h1.title-text'));
    const subtitle = slideEl.find('p').eq(1).text().trim();
    const objectives = [];

    slideEl.find('.info-box li').each((i, li) => {
      objectives.push($(li).text().replace(/[✓●]/g, '').trim());
    });

    // Add decorative circles
    slide.addShape(pptx.ShapeType.ellipse, {
      x: -0.5,
      y: -0.5,
      w: 2,
      h: 2,
      fill: { type: 'solid', color: colors.primaryDark, transparency: 85 }
    });

    slide.addShape(pptx.ShapeType.ellipse, {
      x: 8.5,
      y: 4,
      w: 1.5,
      h: 1.5,
      fill: { type: 'solid', color: colors.primaryLight, transparency: 85 }
    });

    // Course badge
    slide.addText('전략적 재고운영 및 자재계획 수립 교육 과정', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 0.3,
      fontSize: 12,
      color: colors.primaryLight,
      bold: true,
      fontFace: 'Noto Sans KR'
    });

    // Accent bar
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5,
      y: 0.85,
      w: 0.6,
      h: 0.04,
      fill: { type: 'solid', color: colors.primaryLight }
    });

    // Main title
    slide.addText(mainTitle, {
      x: 0.5,
      y: 1.5,
      w: 6,
      h: 1.5,
      fontSize: 36,
      bold: true,
      color: colors.primaryDark,
      fontFace: 'Noto Sans KR'
    });

    // Subtitle
    slide.addText(subtitle, {
      x: 0.5,
      y: 3.2,
      w: 6,
      h: 0.4,
      fontSize: 18,
      color: colors.textGray,
      fontFace: 'Noto Sans KR'
    });

    slide.addText('Part 2', {
      x: 0.5,
      y: 3.7,
      w: 6,
      h: 0.3,
      fontSize: 14,
      color: colors.textGray,
      fontFace: 'Noto Sans KR'
    });

    // Learning objectives box
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5,
      y: 4.3,
      w: 6,
      h: 1.1,
      fill: { type: 'solid', color: 'E3F2FD' },
      line: { type: 'solid', color: colors.primaryLight, width: 2 }
    });

    slide.addText('🎯 학습 목표', {
      x: 0.7,
      y: 4.4,
      w: 5.6,
      h: 0.25,
      fontSize: 12,
      bold: true,
      color: colors.primaryDark,
      fontFace: 'Noto Sans KR'
    });

    const objectivesText = objectives.map((obj, i) => `✓ ${obj}`).join('\n');
    slide.addText(objectivesText, {
      x: 0.7,
      y: 4.7,
      w: 5.6,
      h: 0.65,
      fontSize: 9,
      color: colors.textGray,
      fontFace: 'Noto Sans KR'
    });

    // Footer
    slide.addText(`1 / ${slides.length}`, {
      x: 9,
      y: 5.3,
      w: 0.5,
      h: 0.2,
      fontSize: 9,
      color: colors.textGray,
      align: 'right',
      fontFace: 'Noto Sans KR'
    });

  } else {
    // Content slide
    const titleEl = slideEl.find('h3.title-text, h2.title-text');
    const sectionTitleEl = slideEl.find('h2').first();
    const slideTitle = titleEl.length > 0 ? extractText(titleEl) : '';
    const sectionTitle = sectionTitleEl.length > 0 ? extractText(sectionTitleEl) : '';

    // Add section title with accent bar
    if (sectionTitle && slideTitle !== sectionTitle) {
      addSectionTitle(slide, sectionTitle, 0.3);
    }

    // Add slide title
    if (slideTitle) {
      slide.addText(slideTitle, {
        x: 0.5,
        y: sectionTitle ? 0.85 : 0.5,
        w: 9,
        h: 0.5,
        fontSize: 20,
        bold: true,
        color: colors.primaryDark,
        fontFace: 'Noto Sans KR'
      });
    }

    // Extract and add content (simplified for now)
    const contentStartY = sectionTitle ? 1.5 : 1.2;

    // Handle different content types
    const featureCards = slideEl.find('.feature-card');
    const infoBoxes = slideEl.find('.info-box, .success-box, .warning-box');
    const tables = slideEl.find('table');

    if (featureCards.length > 0) {
      // Grid layout for feature cards
      const cardsPerRow = 2;
      const cardW = 4.2;
      const cardH = 1.4;
      const gapX = 0.4;
      const gapY = 0.3;

      featureCards.each((i, card) => {
        const row = Math.floor(i / cardsPerRow);
        const col = i % cardsPerRow;
        const x = 0.5 + col * (cardW + gapX);
        const y = contentStartY + row * (cardH + gapY);

        const cardTitle = extractText($(card).find('h4'));
        const cardContent = extractText($(card).find('p'));

        addFeatureCard(slide, x, y, cardW, cardH, '📌', cardTitle, cardContent);
      });
    }

    if (infoBoxes.length > 0) {
      infoBoxes.each((i, box) => {
        const boxEl = $(box);
        const boxTitle = extractText(boxEl.find('h4, strong').first());
        const boxContent = extractText(boxEl);

        // Determine box color based on class
        let boxColor = 'E3F2FD'; // info-box default
        let borderColor = colors.primaryLight;

        if (boxEl.hasClass('success-box')) {
          boxColor = colors.greenBg;
          borderColor = colors.greenText;
        } else if (boxEl.hasClass('warning-box')) {
          boxColor = 'FFF3E0';
          borderColor = colors.orangeText;
        }

        const yPos = contentStartY + (i * 1.2);

        slide.addShape(pptx.ShapeType.rect, {
          x: 0.5,
          y: yPos,
          w: 9,
          h: 1,
          fill: { type: 'solid', color: boxColor },
          line: { type: 'solid', color: borderColor, width: 2 }
        });

        slide.addText(boxTitle, {
          x: 0.7,
          y: yPos + 0.1,
          w: 8.6,
          h: 0.25,
          fontSize: 11,
          bold: true,
          color: colors.textDark,
          fontFace: 'Noto Sans KR'
        });

        slide.addText(boxContent.replace(boxTitle, '').trim(), {
          x: 0.7,
          y: yPos + 0.4,
          w: 8.6,
          h: 0.5,
          fontSize: 9,
          color: colors.textGray,
          fontFace: 'Noto Sans KR'
        });
      });
    }

    if (tables.length > 0) {
      // Add table (simplified)
      const table = tables.first();
      const headers = [];
      const rows = [];

      table.find('thead th').each((i, th) => {
        headers.push({ text: extractText($(th)), options: { bold: true, color: colors.white, fill: colors.primaryDark } });
      });

      table.find('tbody tr').each((i, tr) => {
        const row = [];
        $(tr).find('td').each((j, td) => {
          row.push({ text: extractText($(td)), options: { fontSize: 9 } });
        });
        rows.push(row);
      });

      if (headers.length > 0) {
        slide.addTable([headers, ...rows], {
          x: 0.5,
          y: contentStartY,
          w: 9,
          fontSize: 10,
          fontFace: 'Noto Sans KR',
          border: { pt: 1, color: 'CCCCCC' }
        });
      }
    }

    // Footer
    const slideNumber = slideEl.find('.text-xs p').last().text().trim();
    slide.addText(slideNumber || `${slideCount} / ${slides.length}`, {
      x: 9,
      y: 5.3,
      w: 0.5,
      h: 0.2,
      fontSize: 9,
      color: colors.textGray,
      align: 'right',
      fontFace: 'Noto Sans KR'
    });

    slide.addText('소싱 전략 및 유형별 재고관리 프로세스', {
      x: 0.5,
      y: 5.3,
      w: 6,
      h: 0.2,
      fontSize: 9,
      color: colors.primaryLight,
      fontFace: 'Noto Sans KR'
    });
  }
});

// Save PPTX
console.log('\n💾 Saving PowerPoint file...');
pptx.writeFile({ fileName: outputFile })
  .then(() => {
    console.log('\n✅ CONVERSION COMPLETE!');
    console.log(`📊 Output: ${outputFile}`);
    console.log(`📈 Total slides: ${slideCount}`);
    console.log('\n💡 Next steps:');
    console.log(`   1. Open the PPTX file`);
    console.log(`   2. Review all slides`);
    console.log(`   3. Adjust formatting as needed`);
  })
  .catch((err) => {
    console.error('\n❌ Error saving PPTX:', err);
    process.exit(1);
  });
