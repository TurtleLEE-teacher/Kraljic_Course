# Design Patterns Findings from Reference Materials

## Analysis Summary (2025-11-19)

This document consolidates key design patterns extracted from reference PPTX and PDF files to enhance our course presentation guidelines.

---

## 1. Reference Materials Analyzed

### PPTX Files
1. **[RedSlide]문차트.pptx** (127KB, 3 slides)
   - Dimensions: 10.83" × 7.50" (1.44:1) ✓ Confirms S4HANA standard
   - Focus: "문차트" (Door Chart) visualization technique
   - Slide 2 has 75 shapes (16 AUTO_SHAPES, 56 GROUPS)
   - High density visual presentation

2. **추가자료1_Time Saver_2024_v1.0.pptx** (2.3MB, 73 slides)
   - Dimensions: 10.00" × 7.50" (1.33:1 - 4:3 ratio)
   - Font sizes: 61.2pt, 19.4pt, 12.0pt (most common)
   - Primarily "Stickers and Templates" collection

3. **추가자료2_Inspiration _2024.pptx** (1.3MB, 40 slides)
   - Dimensions: 10.83" × 7.50" (1.44:1) ✓ Confirms S4HANA standard
   - **Font size distribution (CRITICAL finding):**
     - **10pt: 131 uses** (MOST COMMON - 65%)
     - 12pt: 47 uses (23%)
     - 14pt: 13 uses (6%)
     - 18pt: 6 uses (3%)
   - Content: Professional business consulting slides
   - Topics: SSC (Shared Service Centers), financial analysis, capacity allocation

### PDF Files
4. **Day2_ppt_lec_아이컨_v1.0_unlocked.pdf** (4.0MB, 89 pages)
   - Source: 아이컨(인사이트베이) professional PPT training
   - Topic: "Principle & Storyline 그리고 만들어보기"
   - Key methods:
     - Structural Approach
     - Dynamics Approach
     - Market Change based Approach

5. **Day3_ppt_lec_아이컨_v1.0_1_unlocked.pdf** (3.1MB, 48 pages)
   - Source: 아이컨(인사이트베이) professional PPT training
   - Focus: Business strategy presentation
   - Teaches: "보고서 전개에 따른 표현방식" (Expression methods based on report flow)

6. **[RedSlide]문차트를 활용한 보고서.pdf** (420KB, 4 pages)
   - Author: RedSlide (홍장표, hjp2025@gmail.com)
   - Font: 본고딕 (Noto Sans CJK KR Bold)
   - Focus: Door Chart technique for reports

---

## 2. Key Design Patterns Discovered

### 2.1 Font Size Strategy (CRITICAL UPDATE)

**Current guideline says:** 8-11pt for body text
**New evidence confirms:** 10pt is THE dominant size (65% of all text)

**Revised Font Size Hierarchy:**
```
48pt Bold     - Cover slide main title
20pt Bold     - Content slide titles
16pt Bold     - Governing messages (NOT 14pt Italic!)
14pt          - Section headers, large bullets
12-13pt       - Regular bullets, emphasis
10-11pt       - Body text, descriptions (PRIMARY - use 10pt most)
8-9pt         - Small annotations, footnotes
6-7pt         - Tiny notes (rare)
```

**Key insight:** The 10pt font size enables **high content density** while maintaining readability. This is THE secret to achieving 85%+ filled slides.

### 2.2 "문차트" (Door Chart) Pattern

**Discovery:** A specific visualization technique from RedSlide

**Characteristics:**
- Creates a "door-like" opening effect
- Uses directional indicators: "매우 높음 →" and "← 매우 낮음"
- Highly structured with extensive grouping (56 groups in one slide!)
- Combines multiple visual elements into complex compositions

**Use cases:**
- Showing spectrum or range (high to low)
- Priority matrices
- Risk assessments
- Strategic positioning

**Implementation notes:**
- Requires heavy use of AUTO_SHAPES and GROUP elements
- Should use 40-75+ shapes per slide for this pattern
- Aligns with S4HANA high-density principle

### 2.3 Structural Approach Methodology

**From Day2 PDF - Three approaches to problem-solving:**

1. **Structural Approach** (구조적 접근)
   - Break down problems into logical structures
   - Use frameworks and matrices
   - Build arguments from foundation

2. **Dynamics Approach** (역동적 접근)
   - Focus on change and transformation
   - Show before/after comparisons
   - Emphasize impact and results

3. **Market Change Based Approach** (시장 변화 기반)
   - Start from market trends
   - Connect external changes to internal strategy
   - Data-driven storytelling

**Application to Kraljic Course:**
- Session 1: Structural (Framework introduction)
- Session 2-7: Dynamics (Transformation per material type)
- Session 8-9: Market Change (Real-world application)

### 2.4 Storyline Development (from 아이컨 training)

**Key principle:** "청자의 마음에 Penetrate 하는 것이 핵심"
(The core is to penetrate the listener's heart/mind)

**Elements of persuasive presentation:**
1. **설득력 있는 전략** (Convincing strategy)
   - Clear governing messages
   - Evidence-based arguments
   - Logical flow

2. **구조화된 전개** (Structured development)
   - Consistent section numbering (X.Y format)
   - Clear chapter divisions
   - Progressive complexity

3. **시각적 임팩트** (Visual impact)
   - Use 10-50+ shapes per slide
   - Combine text + visuals
   - Toy Page layout (60-70% visual, 30-40% text)

---

## 3. Content Density Analysis

### High-Density Slide Composition (from RedSlide example)

**Slide 2 breakdown:**
- Total shapes: 75
  - AUTO_SHAPES: 16 (21%)
  - GROUPS: 56 (75%)
  - TEXT_BOX: 2 (3%)
  - PLACEHOLDER: 1 (1%)

**Interpretation:**
- Heavy use of grouped elements (75% are groups)
- This allows complex visual compositions while maintaining organization
- Groups enable reusable patterns and consistent styling

**Recommendation for Kraljic Course:**
- Aim for 20-50 shapes per content slide
- Use groups to organize related elements
- Combine: boxes + arrows + text + icons

---

## 4. Professional Training Insights

### From 아이컨(인사이트베이) Materials

**Key teaching points:**
1. **실전 중심** (Practice-focused)
   - All examples adapted from real projects
   - Uses realistic company names and numbers (anonymized)
   - Emphasizes practical application

2. **각색과 예시화** (Adaptation and exemplification)
   - Takes logic and story from real cases
   - Modifies for educational effectiveness
   - Balances realism with learning objectives

3. **표현방식의 체계화** (Systemization of expression methods)
   - Different layouts for different message types
   - Consistent patterns for recurring concepts
   - Visual grammar for business storytelling

**Application to our course:**
- Use realistic supplier names (already done: 미래금속, 글로벌스틸, etc.)
- Anonymize but keep realistic (already done: scores, grades)
- Adapt Kraljic framework to look like real consulting project

---

## 5. Color and Typography Validation

### Dimensions Confirmed
- **S4HANA standard:** 10.83" × 7.50" (1.44:1) ✓
- **Alternative:** 10.00" × 7.50" (1.33:1 - 4:3) for some content

### Font Usage Patterns
- **Primary font sizes:** 10pt (body), 12pt (bullets), 14pt (headers)
- **This confirms:** Small fonts (8-11pt) are PROFESSIONAL STANDARD
- **Avoid:** Using 16-18pt for body text (too large, wastes space)

### Color System
- Monochrome confirmed as professional standard
- Color only for:
  - Charts and data visualizations
  - Specific diagrams (like Kraljic Matrix)
  - Cover slides (minimal)

---

## 6. Recommendations for Implementation

### Immediate Updates to Guidelines

1. **Update Font Size Table:**
   - Emphasize 10pt as PRIMARY body text size
   - Show actual usage statistics (10pt = 65% of text)
   - Demote 16-18pt to "large bullets only"

2. **Add "문차트" Pattern:**
   - New section: "Door Chart Visualization"
   - Include shape count requirements (40-75+)
   - Emphasize GROUP usage for complex compositions

3. **Add Storyline Structure:**
   - New section: "Persuasive Storyline Development"
   - Include three approaches: Structural, Dynamics, Market Change
   - Show how to "penetrate listener's mind"

4. **Add Professional Standards:**
   - Real-world adaptation principles
   - Anonymization techniques
   - Balance between realism and education

5. **Enhance Toy Page Documentation:**
   - Add specific percentage splits: 60-70% visual, 30-40% text
   - Show shape count targets per section
   - Include grouping strategies

### Priority Actions

**HIGH PRIORITY:**
1. Update font size guidance (10pt is king!)
2. Add shape count targets (20-50 per slide)
3. Document grouping strategies

**MEDIUM PRIORITY:**
1. Add 문차트 pattern details
2. Expand storyline development section
3. Add case study adaptation guidelines

**LOW PRIORITY:**
1. Additional layout variations
2. Advanced visual techniques
3. Animation and transition guidelines (if needed)

---

## 7. Validation Against Existing Guidelines

### What's Already Correct ✓
- Monochrome color system
- Governing message requirement
- Toy Page layout concept
- High content density principle (85%+)
- Slide dimensions (10.83" × 7.50")

### What Needs Enhancement ⚡
- **Font sizes:** Add more emphasis on 10pt dominance
- **Shape counts:** Add specific targets (20-50 shapes)
- **Grouping:** Document GROUP element strategy
- **Storyline:** Add persuasion and narrative techniques
- **실전 중심:** Add professional adaptation guidelines

### What's Missing 🆕
- **문차트 pattern:** New visualization technique
- **Three approaches:** Structural, Dynamics, Market Change
- **Usage statistics:** Actual font size distribution data
- **Professional training insights:** From 아이컨 materials

---

## 8. Next Steps

1. **Update design-guidelines.md:**
   - Add new font size statistics
   - Add 문차트 pattern section
   - Enhance Toy Page documentation
   - Add storyline development section

2. **Update CLAUDE.md:**
   - Reference new patterns
   - Add font size emphasis
   - Include shape count requirements
   - Add professional adaptation guidelines

3. **Create examples:**
   - Generate sample slide using 10pt body text
   - Create 문차트 visualization example
   - Build Toy Page template with shape counts

4. **Test and validate:**
   - Generate sample PPTX with new guidelines
   - Verify 85%+ content density achieved
   - Confirm 20-50 shapes per slide
   - Validate 10pt readability

---

## Appendix: Data Tables

### Font Size Distribution (from Inspiration_2024.pptx)

| Font Size | Count | Percentage | Usage |
|-----------|-------|------------|-------|
| 10pt | 131 | 65.2% | Body text, descriptions (PRIMARY) |
| 12pt | 47 | 23.4% | Bullets, emphasis |
| 14pt | 13 | 6.5% | Section headers |
| 18pt | 6 | 3.0% | Large headers |
| 9pt | 1 | 0.5% | Small notes |
| **Total** | **201** | **100%** | |

### Slide Dimension Standards

| Source | Width | Height | Ratio | Status |
|--------|-------|--------|-------|--------|
| S4HANA Reference | 10.83" | 7.50" | 1.44:1 | ✓ Primary |
| RedSlide 문차트 | 10.83" | 7.50" | 1.44:1 | ✓ Confirmed |
| Inspiration 2024 | 10.83" | 7.50" | 1.44:1 | ✓ Confirmed |
| Time Saver 2024 | 10.00" | 7.50" | 1.33:1 | Alternative |

### Shape Count Analysis (RedSlide Slide 2)

| Shape Type | Count | Percentage |
|------------|-------|------------|
| GROUP | 56 | 74.7% |
| AUTO_SHAPE | 16 | 21.3% |
| TEXT_BOX | 2 | 2.7% |
| PLACEHOLDER | 1 | 1.3% |
| **Total** | **75** | **100%** |

---

**Document Version:** 1.0
**Date:** 2025-11-19
**Analyst:** Claude (AI Assistant)
**Purpose:** Extract design patterns from reference materials to enhance Kraljic Course presentation quality
