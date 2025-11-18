# Design Guidelines for Business Reporting (S4HANA Style)

## Overview

This document provides detailed guidelines for creating professional business presentations and reports using the S4HANA monochrome design system. This style emphasizes clarity, high information density, and professional appearance for executive briefings and project reports.

## Core Principles

### 1. Monochrome Color System

**Rule**: Use black/gray scale only for text and backgrounds. No bright colors except in data visualization (charts).

**Rationale**:
- Professional, formal appearance
- Maximum readability and contrast
- Focus on content, not decoration
- Universal accessibility

**Implementation**:
```css
:root {
  --color-text: #000000;           /* Primary text */
  --color-text-secondary: #333333; /* Secondary text */
  --color-text-muted: #666666;     /* Muted text, captions */
  --color-border: #CCCCCC;         /* Borders, separators */
  --color-bg: #FFFFFF;             /* Background */
  --color-bg-light: #F5F5F5;       /* Light background, boxes */
}
```

**Color usage**:
- ✅ Good: Black titles + Gray body text + Light gray boxes + White background
- ❌ Bad: Blue headers + Green highlights + Orange boxes (except charts)

### 2. Typography System (Arial Only)

**Font**: Arial - web-safe, universally available, professional

**Font Size System**:
```
20pt Bold     - Slide titles
14pt Italic   - Governing messages
14pt Bold     - Section headings (h2, h3)
12pt Regular  - Body text, bullet points
11pt Regular  - Captions, table content
```

**Rationale**:
- Arial ensures compatibility across all platforms
- Clear hierarchy through size and weight
- Maximum readability
- No custom fonts = no rendering issues

### 3. Governing Message Pattern (REQUIRED)

**Definition**: A one-sentence summary placed directly under the slide title that captures the entire slide's key point.

**Requirements**:
- **Mandatory** on all content slides
- Font: 14pt italic
- Color: `--color-text-muted` (#666666)
- Length: 1-2 sentences maximum
- Position: Immediately below title, within title-section

**Implementation**:
```html
<div class="title-section fit">
  <h1>Slide Title</h1>
  <p class="governing-message">One-sentence summary that captures the slide's key insight.</p>
</div>
```

**Examples**:
- ✅ "자재 특성에 따라 ROP와 MRP 방식을 차별적으로 적용하여 재고 최적화와 공급 안정성을 동시에 달성합니다."
- ✅ "Modern approach delivers 40% faster results through automation and data-driven decisions."
- ❌ "This slide is about ROP and MRP." (Too vague, no insight)
- ❌ "We will discuss various inventory management methods including ROP, MRP, EOQ, and JIT, each with different applications..." (Too long)

### 4. Layout Requirements

**Slide Dimensions**:
```css
body {
  width: 780px;   /* 13:9 aspect ratio */
  height: 540px;
  padding: 25px 40px 45px 40px;  /* top, sides, bottom */
}
```

**Aspect Ratio**: 13:9 (business reporting standard)
- Not 16:9 (widescreen) or 4:3 (traditional)
- Optimized for professional business documents
- Matches S4HANA project presentation format

**Padding Guidelines**:
- Top: 25px (minimal, slides are dense)
- Sides: 40px (consistent left/right margins)
- Bottom: 45px (minimum 0.5" = ~48px margin required)

**Critical**: Minimum 0.5" (45px) bottom margin to prevent text overflow

### 5. Information Density

**Target**: 80%+ content density

**Principle**: S4HANA business reports maximize information while maintaining readability

**Characteristics**:
- Dense content is acceptable if well-structured
- Minimal decorative elements
- Focus on data and insights
- Use full slide space efficiently

**Bad practices (too much whitespace)**:
- ❌ Large empty areas
- ❌ Oversized titles (> 24pt)
- ❌ Excessive padding (> 60px)
- ❌ Single bullet point centered on slide

**Good practices (efficient use of space)**:
- ✅ 2-column layouts for comparisons
- ✅ Tables with multiple rows
- ✅ Bullet lists with 4-6 items
- ✅ Combined text + data visualizations

### 6. Structured Content (구조화)

**Principle**: Content can be dense, but must be well-structured

**Good practices**:
- ✅ Use headings to organize sections
- ✅ Use bullet points for lists
- ✅ Use numbering for sequential steps
- ✅ Group related information
- ✅ White space for visual breaks

**Example of well-structured dense content (S4HANA style)**:

```html
<div class="title-section fit">
  <h1>병목자재 vs 레버리지자재 - 특성 및 관리 전략</h1>
  <p class="governing-message">공급 리스크와 사업 영향도의 차이에 따라 병목자재는 공급 안정성에, 레버리지자재는 원가 절감에 집중합니다.</p>
</div>

<div class="fill-height row gap items-fill-width">
  <div>
    <h2>병목자재 (Bottleneck)</h2>

    <h3>특성</h3>
    <ul>
      <li>공급 리스크: 높음 (소수/독점 공급업체)</li>
      <li>사업 영향도: 낮음 (소액 구매)</li>
      <li>대체 가능성: 낮음 (기술적 제약)</li>
    </ul>

    <h3>관리 전략</h3>
    <ul>
      <li>대체품 확보: 유사 기능 자재 발굴</li>
      <li>안전재고 유지: 리드타임 2배 이상</li>
      <li>장기 계약: 공급 연속성 확보</li>
    </ul>
  </div>

  <div>
    <h2>레버리지자재 (Leverage)</h2>

    <h3>특성</h3>
    <ul>
      <li>공급 리스크: 낮음 (다수 공급업체)</li>
      <li>사업 영향도: 높음 (대량 구매)</li>
      <li>대체 가능성: 높음 (표준화 품목)</li>
    </ul>

    <h3>관리 전략</h3>
    <ul>
      <li>경쟁입찰: 연 2회 이상 RFQ 실시</li>
      <li>가격협상: Volume Discount 활용</li>
      <li>시장 분석: 원자재 가격 동향 모니터링</li>
    </ul>
  </div>
</div>
```

### 7. 2-Column Comparison Layouts

**Use case**: Perfect for "vs" slides (ROP vs MRP, Before vs After, Traditional vs Modern)

**Requirements**:
- Equal width columns (50:50 split)
- Use `.row .gap .items-fill-width` classes
- Consistent heading levels in both columns
- Similar content length for visual balance

**Example**: See above "병목자재 vs 레버리지자재" structure

## Complete Example (S4HANA Style)

**See**: `temp/s4hana-style/slide4-improved.html` for full working example

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>재고관리 방법론 비교</title>
  <style>
    :root {
      --color-text: #000000;
      --color-text-secondary: #333333;
      --color-text-muted: #666666;
      --color-border: #CCCCCC;
      --color-bg: #FFFFFF;
      --color-bg-light: #F5F5F5;
    }
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    body {
      width: 780px;
      height: 540px;
      background: var(--color-bg);
      font-family: Arial, Helvetica, sans-serif;
      color: var(--color-text-secondary);
      padding: 25px 40px 45px 40px;
      overflow: hidden;
    }
    .title-section {
      border-bottom: 2px solid var(--color-border);
      padding-bottom: 8px;
      margin-bottom: 15px;
    }
    h1 {
      font-size: 20pt;
      font-weight: bold;
      color: var(--color-text);
      margin-bottom: 6px;
    }
    .governing-message {
      font-size: 14pt;
      color: var(--color-text-muted);
      font-style: italic;
      line-height: 1.3;
    }
    h2 {
      font-size: 14pt;
      font-weight: bold;
      color: var(--color-text);
    }
    p {
      font-size: 12pt;
      line-height: 1.4;
    }
    ul {
      margin-left: 18px;
    }
    li {
      font-size: 12pt;
      line-height: 1.4;
      margin-bottom: 3px;
    }
  </style>
</head>
<body class="col">
  <div class="title-section fit">
    <h1>재고관리 방법론 비교: ROP vs MRP</h1>
    <p class="governing-message">자재 특성에 따라 ROP와 MRP 방식을 차별적으로 적용하여 재고 최적화와 공급 안정성을 동시에 달성합니다.</p>
  </div>

  <div class="fill-height row gap items-fill-width">
    <!-- ROP -->
    <div>
      <h2>ROP (Reorder Point)</h2>
      <h3>적용 대상</h3>
      <p>병목자재 - 공급 리스크 높음</p>

      <h3>핵심 원리</h3>
      <p>재고 수준이 발주점에 도달하면 자동으로 발주</p>

      <h3>주요 특징</h3>
      <ul>
        <li>독립 수요 기반</li>
        <li>재고 수준 모니터링</li>
        <li>자동화된 의사결정</li>
      </ul>
    </div>

    <!-- MRP -->
    <div>
      <h2>MRP (Material Requirements Planning)</h2>
      <h3>적용 대상</h3>
      <p>레버리지자재 - 공급 리스크 낮음</p>

      <h3>핵심 원리</h3>
      <p>생산계획 기반으로 필요 자재의 수량과 시점 계산</p>

      <h3>주요 특징</h3>
      <ul>
        <li>종속 수요 기반</li>
        <li>BOM 구조 활용</li>
        <li>계획적 구매 시점 결정</li>
      </ul>
    </div>
  </div>
</body>
</html>
```

## Quality Checklist

Before finalizing any S4HANA business presentation:

**Color (Monochrome)**:
- [ ] Only black/gray scale for text and backgrounds
- [ ] No bright colors except in charts
- [ ] Consistent use of color variables
- [ ] Maximum contrast for readability

**Typography**:
- [ ] Arial font only (no custom fonts)
- [ ] Governing message present on all content slides (14pt italic)
- [ ] Font hierarchy maintained (20pt > 14pt > 12pt > 11pt)
- [ ] No font smaller than 11pt

**Layout**:
- [ ] Slide dimensions: 780px × 540px (13:9 ratio)
- [ ] Padding: 25px top, 40px sides, 45px bottom
- [ ] Minimum 0.5" (45px) bottom margin
- [ ] Content density 80%+ (minimal whitespace)

**Governing Messages**:
- [ ] Present on all content slides
- [ ] One sentence maximum
- [ ] Captures key insight (not just topic)
- [ ] 14pt italic, gray color (#666666)

**2-Column Layouts (for comparisons)**:
- [ ] Equal width columns (`.items-fill-width`)
- [ ] Consistent heading levels in both columns
- [ ] Similar content length for balance
- [ ] Clear visual separation

**Content Organization**:
- [ ] Headings organize sections
- [ ] Bullet points for lists
- [ ] Numbering for sequences
- [ ] Structured hierarchy (h1 > h2 > h3)
- [ ] Dense content is well-organized

**Tables**:
- [ ] NOT using HTML `<table>` tags
- [ ] Using PptxGenJS `slide.addTable()` instead
- [ ] Proper column widths and row heights defined
- [ ] Border and styling applied

## Tips

1. **Start with governing message**: Capture the "so what" in one sentence
2. **Monochrome only**: Resist the urge to add color (except charts)
3. **Maximize density**: Use full slide space efficiently
4. **2-column for comparisons**: Side-by-side beats vertical stacking
5. **Check bottom margin**: Ensure minimum 45px to prevent text overflow
6. **Arial everywhere**: No exceptions for fonts
