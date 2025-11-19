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

**Font Size System** (based on professional S4HANA analysis):
```
48pt Bold     - Cover slide main title
20pt Bold     - Content slide titles
16pt Bold     - Governing messages (updated from 14pt Italic)
14pt Bold     - Section headings (h2, h3)
12-13pt       - Regular bullet points, emphasis
10-11pt       - Body text, descriptions (PRIMARY - use 10pt most)
8-9pt         - Small annotations, footnotes
6-7pt         - Tiny notes (rare)
```

**CRITICAL:** Analysis of professional presentations shows **10pt is the dominant body text size (65% of all text)**. This enables high content density (85%+) while maintaining readability.

**Font Size Distribution (from professional samples):**
- 10pt: 65.2% (body text, descriptions) - **MOST COMMON**
- 12pt: 23.4% (bullets, emphasis)
- 14pt: 6.5% (section headers)
- 18pt: 3.0% (large headers)
- Other: 1.9% (small notes)

**Rationale**:
- Arial/맑은 고딕 ensures compatibility across all platforms
- Clear hierarchy through size and weight
- **10pt font is THE professional standard** for high-density slides
- Maximum readability at small sizes
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

### 8. Shape Count and Visual Density Requirements

**Principle**: Professional S4HANA slides use **20-50+ shapes per content slide** to achieve high visual impact and information density.

**Target shape counts**:
- **Simple slides**: 10-20 shapes (text boxes, basic rectangles)
- **Standard slides**: 20-40 shapes (boxes, arrows, icons)
- **Complex slides**: 40-75+ shapes (door charts, process flows, matrices)

**Shape types to use**:
- **Rectangles/AUTO_SHAPES**: 15-30% - Wrap all text content
- **GROUPS**: 70-80% - Organize related elements into logical groups
- **Arrows/Connectors**: 5-10% - Show relationships, flow, sequence
- **TEXT_BOX/PLACEHOLDER**: 5-10% - Labels, annotations

**GROUP element strategy** (CRITICAL):
- **Use groups extensively** (70-80% of shapes should be in groups)
- Groups enable complex visual compositions while maintaining organization
- Groups allow reusable patterns and consistent styling
- Examples: Process step group, comparison box group, timeline phase group

**Door Chart (문차트) Pattern**:
A specialized high-density visualization technique for showing spectrums, priorities, or risk matrices.

**Characteristics**:
- **75+ shapes** in a single slide
- Heavy use of groups (70-80% of elements)
- Directional indicators: "매우 높음 →" and "← 매우 낮음"
- Creates visual "opening" or "door" effect
- Combines multiple layers of information

**Use cases**:
- Priority matrices (urgent vs important)
- Risk assessment grids (likelihood vs impact)
- Strategic positioning (Kraljic Matrix!)
- Spectrum visualizations (traditional → modern)

**Implementation**:
```html
<!-- Door Chart structure -->
<div class="door-chart">
  <!-- Left door: Low end -->
  <div class="door-panel left">
    <div class="indicator">← 매우 낮음</div>
    <div class="content-group">
      <!-- 20-30 shapes: boxes, text, icons -->
    </div>
  </div>

  <!-- Center: Transition zone -->
  <div class="door-center">
    <div class="axis-label">공급 리스크</div>
    <!-- Gradient or spectrum visualization -->
  </div>

  <!-- Right door: High end -->
  <div class="door-panel right">
    <div class="indicator">매우 높음 →</div>
    <div class="content-group">
      <!-- 20-30 shapes: boxes, text, icons -->
    </div>
  </div>
</div>
```

**Benefits**:
- High visual impact (immediately draws attention)
- Clear spectrum/range communication
- Supports dense information while maintaining structure
- Perfect for Kraljic Matrix 2x2 diagrams

### 9. Persuasive Storyline Development

**Principle**: "청자의 마음에 Penetrate 하는 것이 핵심" (The core is to penetrate the listener's heart/mind)

**Three Strategic Approaches** (from professional training):

#### 9.1 Structural Approach (구조적 접근)
Break down problems into logical structures and frameworks.

**When to use**:
- Introducing new frameworks (Session 1: Kraljic Matrix)
- Building foundational knowledge
- Explaining methodologies

**Technique**:
- Start with the framework (2x2 matrix, flowchart)
- Break down into components
- Build arguments from foundation to application
- Use hierarchical structure (1.1, 1.2, 1.3...)

**Example slide flow**:
1. Framework overview
2. Component 1 details
3. Component 2 details
4. How components interact
5. Application examples

#### 9.2 Dynamics Approach (역동적 접근)
Focus on change, transformation, and impact.

**When to use**:
- Showing improvements (JIT → JIC transformation)
- Before/after comparisons
- Demonstrating value/ROI

**Technique**:
- Start with "before" state (problems, pain points)
- Introduce the change catalyst
- Show transformation process
- Emphasize results and impact
- Use visual contrast (dark → light, small → large)

**Example slide flow**:
1. Current state challenges
2. What changed (new approach)
3. How transformation happened
4. After state (improvements)
5. Quantified impact (metrics)

#### 9.3 Market Change Based Approach (시장 변화 기반)
Start from external market trends and connect to internal strategy.

**When to use**:
- Strategic planning (Sessions 8-9: Workshops)
- Business case development
- Justifying new initiatives

**Technique**:
- Begin with market data/trends
- Connect external changes to internal implications
- Show how strategy addresses market reality
- Data-driven storytelling
- Link to competitive advantage

**Example slide flow**:
1. Market trend data (2020 pandemic → supply chain disruption)
2. Industry impact (JIT failures)
3. Implications for our business
4. Strategic response (JIC adoption)
5. Competitive positioning

**Combining Approaches**:
- **Session 1-3**: Structural (build framework knowledge)
- **Session 4-7**: Dynamics (show transformation per material type)
- **Session 8-9**: Market Change (apply to real-world scenarios)

**Key elements of persuasive slides**:
1. **설득력 있는 전략** (Convincing strategy)
   - Clear governing messages (one-sentence insights)
   - Evidence-based arguments (data, case studies)
   - Logical flow (cause → effect)

2. **구조화된 전개** (Structured development)
   - Consistent section numbering (X.Y format)
   - Clear chapter divisions (1장, 2장, 3장...)
   - Progressive complexity (foundation → application)

3. **시각적 임팩트** (Visual impact)
   - Use 20-50+ shapes per slide
   - Combine text + visuals (Toy Page: 60-70% visual, 30-40% text)
   - Door charts for high-impact moments

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
- [ ] Arial/맑은 고딕 font only (no custom fonts)
- [ ] Governing message present on all content slides (16pt bold)
- [ ] Font hierarchy maintained (48pt > 20pt > 16pt > 14pt > 12pt > 10pt)
- [ ] **10pt is PRIMARY body text size** (should be 60-70% of all text)
- [ ] 12pt for bullets and emphasis (20-25% of text)
- [ ] No font smaller than 8pt (footnotes only)

**Layout**:
- [ ] Slide dimensions: 780px × 540px (13:9 ratio)
- [ ] Padding: 25px top, 40px sides, 45px bottom
- [ ] Minimum 0.5" (45px) bottom margin
- [ ] Content density 80%+ (minimal whitespace)

**Governing Messages**:
- [ ] Present on all content slides
- [ ] One sentence maximum
- [ ] Captures key insight (not just topic)
- [ ] 16pt bold, gray color (#666666)

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

**Shape Counts and Visual Density** (NEW):
- [ ] **20-50+ shapes per content slide** (target range)
- [ ] Simple slides: 10-20 shapes minimum
- [ ] Standard slides: 20-40 shapes
- [ ] Complex slides (door charts, matrices): 40-75+ shapes
- [ ] **70-80% of shapes should be in GROUPS** (for organization)
- [ ] Use arrows/connectors to show flow and relationships
- [ ] All text wrapped in rectangles/boxes (no floating text)

**Storyline and Persuasion** (NEW):
- [ ] Slide flow follows one of three approaches:
  - Structural: Framework → Components → Application
  - Dynamics: Before → Change → After → Impact
  - Market Change: Trend → Implication → Strategy → Positioning
- [ ] Section numbering consistent (X.Y format: 1.1, 1.2, 2.1...)
- [ ] Chapter divisions clear (1장, 2장, 3장...)
- [ ] Governing messages "penetrate listener's mind" (insight, not just topic)
- [ ] Evidence-based arguments (data, case studies, metrics)
- [ ] Logical flow maintained throughout presentation

## Tips

1. **Start with governing message**: Capture the "so what" in one sentence that penetrates the listener's mind
2. **Monochrome only**: Resist the urge to add color (except charts and Kraljic Matrix)
3. **Use 10pt body text**: This is THE professional standard (65% of all text) - enables high density
4. **Maximize density**: Use full slide space efficiently - aim for 85%+ filled
5. **20-50 shapes per slide**: Don't just use text - add boxes, arrows, diagrams
6. **Group extensively**: 70-80% of shapes should be in groups for organization
7. **2-column for comparisons**: Side-by-side beats vertical stacking (Toy Page: 60-70% visual, 30-40% text)
8. **Follow storyline approach**: Choose Structural, Dynamics, or Market Change flow
9. **Check bottom margin**: Ensure minimum 45px to prevent text overflow
10. **Arial/맑은 고딕 everywhere**: No exceptions for fonts
