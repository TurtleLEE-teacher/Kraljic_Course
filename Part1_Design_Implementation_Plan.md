# Part 1 Design Implementation Plan
**Session:** 1회차 - 전략적 재고운영 Foundation
**Generated:** 2025-11-20
**Purpose:** Detailed design plan to achieve S4HANA professional quality standards

---

## S4HANA Analysis Results (ACTUAL DATA)

### Slide Dimensions
- ✅ Width: 10.83"
- ✅ Height: 7.50"
- ✅ Aspect ratio: 1.44:1

### Shape Count Analysis (First 15 slides)
| Slide | Shapes | AUTO_SHAPES | GROUPS | LINES | TEXT_BOX | Density |
|-------|--------|-------------|--------|-------|----------|---------|
| **4** | **56** | **26** | **6** | **16** | **7** | **~100%** |
| **10** | **46** | **29** | **10** | **2** | **5** | **~92%** |
| **11** | **65** | **46** | **1** | **11** | **7** | **~100%** |
| **12** | **102** | **87** | **0** | **12** | **3** | **~100%** |
| **13** | **120** | **82** | **0** | **33** | **3** | **~100%** |
| **14** | **102** | **81** | **0** | **16** | **3** | **~100%** |

**Key Insights:**
- Complex slides use 50-120 shapes
- AUTO_SHAPES dominate (50-90% of shapes)
- LINES are critical for arrows/connectors (10-30 per slide)
- GROUPS organize complex compositions (0-10 per slide)

### Font Size Distribution (2,734 text runs analyzed)
| Size | Count | Percentage | Usage |
|------|-------|------------|-------|
| **9pt** | **568** | **20.8%** | **Body text (PRIMARY)** |
| **10pt** | **520** | **19.0%** | **Body text (SECONDARY)** |
| **12pt** | **471** | **17.2%** | **Bullets** |
| **8pt** | **233** | **8.5%** | **Small text** |
| **14pt** | **229** | **8.4%** | **Headers** |
| 20pt | 181 | 6.6% | Slide titles |
| 16pt | 177 | 6.5% | Governing messages |

**Key Insights:**
- 8-10pt combined: **48.3%** (nearly HALF of all text!)
- 12pt: 17.2% (bullets, emphasis)
- Small fonts (6-10pt): **52.4%** (MAJORITY)
- This enables 85-100% content density

---

## Part 1 Current Status (PROBLEMS IDENTIFIED)

### Shape Count Issues
```
Current Part 1:
- Average: 6.6 shapes/slide ❌ (TARGET: 40-100)
- Slide 1: 3 shapes ❌
- Slide 2: 10 shapes ⚠️ (borderline)
- Slides 3-15: 4-5 shapes each ❌
- Slide 11: 21 shapes ⚠️ (only decent one)

Problem: 20/24 slides have < 10 shapes
```

### Font Size Issues
```
Current Part 1:
- 10pt usage: 0% ❌ (TARGET: 19-21%)
- 9pt usage: 0% ❌ (TARGET: 20-21%)
- Most text: 12pt, 16pt, 20pt ❌
- No small fonts used

Problem: Using too large fonts → wastes space → low density
```

### Visual Elements Missing
```
Missing from Part 1:
- ❌ No LINES (arrows, connectors)
- ❌ No flowcharts
- ❌ No timelines
- ❌ No comparison boxes with arrows
- ❌ No door chart for Kraljic Matrix
```

---

## Part 1 Redesign Plan (25 Slides Total)

### Slide Density Targets

| Slide # | Type | Title | Target Shapes | Strategy |
|---------|------|-------|---------------|----------|
| 1 | Cover | 전략적 재고운영 Foundation | 5-10 | Simple cover design |
| 2 | TOC | 목차 | 15-20 | 7 chapter boxes + borders |
| 3 | Divider | 1장: JIT → JIC 패러다임 전환 | 5-10 | Large chapter number + decorative elements |
| 4 | Content | 1.1 JIT의 영광과 몰락 | 40-50 | **Timeline with arrows** (15 boxes + 10 arrows + text) |
| 5 | Content | 1.2 팬데믹이 드러낸 약점 | 30-40 | Problem diagram (10 boxes + 5 arrows + icons) |
| 6 | Table | 1.3 JIT vs JIC 비교 | 25-35 | Table grid (6 rows × 3 cols) + borders |
| 7 | Content | 1.4 JIC 채택 기업들 | 20-30 | Company boxes (8 companies × 2 shapes each) |
| 8 | Divider | 2장: Kraljic Matrix 프레임워크 | 5-10 | Chapter divider |
| 9 | Content | 2.1 Kraljic Matrix 탄생 | 30-40 | Historical timeline (12 boxes + arrows) |
| 10 | Content | 2.2 2×2 매트릭스의 두 축 | 35-45 | Axis explanation (2 large arrows + 15 boxes) |
| 11 | **Door Chart** | **2.3 Kraljic Matrix** | **75-100** | **DOOR CHART: 2×2 with spectrum indicators** |
| 12 | Content | 2.4 병목자재 (Bottleneck) | 35-45 | Characteristics grid (12 boxes + icons + arrows) |
| 13 | Content | 2.5 레버리지자재 (Leverage) | 35-45 | Same pattern as slide 12 |
| 14 | Content | 2.6 전략자재 (Strategic) | 35-45 | Same pattern as slide 12 |
| 15 | Content | 2.7 일상자재 (Routine) | 35-45 | Same pattern as slide 12 |
| 16 | Divider | 3장: 차별화 전략 | 5-10 | Chapter divider |
| 17 | Content | 3.1 차별화의 필요성 | 30-40 | Before/After comparison (10 + 10 boxes + arrow) |
| 18 | Matrix | 3.2 자재군별 전략 매트릭스 | 40-50 | 5-column comparison table (6 rows × 5 cols) |
| 19 | Divider | 4장: 계획 방법론 | 5-10 | Chapter divider |
| 20 | Content | 4.1 5대 방법론 개요 | 30-40 | 4 methodology boxes (4 × 8 shapes) |
| 21 | Content | 4.2 하이브리드 접근법 | 35-45 | Process flow (10 steps + arrows + decision diamonds) |
| 22 | Content | 5장: 통합 KPI 프레임워크 | 40-50 | KPI grid (4 categories × 8 shapes each) |
| 23 | Content | 6장: 산업별 적용 | 30-40 | Industry comparison (5 industries × 6 shapes) |
| 24 | Content | 7장: 9회차 학습 여정 | 45-55 | **Timeline: 9 sessions with arrows** |
| 25 | Summary | Summary & Next Steps | 25-35 | 4 summary boxes (4 × 6 shapes) |

**Total target:** 725-950 shapes across 25 slides
**Average:** 29-38 shapes/slide ✓ (meets 20-50 target)

---

## Shape Usage Plan (PER SLIDE)

### Slide 11: Kraljic Matrix Door Chart (75-100 shapes)

**Implementation:**
```python
# Door chart structure:
# 1. Background grid (1 large rectangle)
# 2. Vertical axis (1 arrow + 2 labels = 3 shapes)
# 3. Horizontal axis (1 arrow + 2 labels = 3 shapes)
# 4. Four quadrants (4 colored rectangles)
# 5. Quadrant labels (4 text boxes)
# 6. Quadrant details (4 × 10 = 40 detail boxes)
# 7. Spectrum indicators (4 arrows: "매우 높음 →", "← 매우 낮음")
# 8. Borders and separators (10 lines)
# 9. Icons for each quadrant (4 icons)
# 10. Example materials (4 × 3 = 12 text boxes)

# Total: 1 + 3 + 3 + 4 + 4 + 40 + 4 + 10 + 4 + 12 = 85 shapes ✓
```

### Slide 4: JIT Timeline (40-50 shapes)

**Implementation:**
```python
# Timeline structure:
# 1. Horizontal timeline arrow (1 shape)
# 2. Year markers (5 points: 1970, 1980, 2000, 2010, 2020 = 5 circles)
# 3. Year labels (5 text boxes)
# 4. Event boxes (5 periods × 3 shapes = 15 boxes)
# 5. Event descriptions (5 × 2 = 10 text boxes)
# 6. Connecting arrows (4 arrows between events)
# 7. Impact indicators (5 triangles: ↑ or ↓)
# 8. Background decorative elements (5 light gray boxes)

# Total: 1 + 5 + 5 + 15 + 10 + 4 + 5 + 5 = 50 shapes ✓
```

### Slide 6: JIT vs JIC Table (25-35 shapes)

**Implementation:**
```python
# Table structure:
# 1. Header row (1 row × 3 cells = 3 rectangles)
# 2. Header text (3 text boxes)
# 3. Row 1: 구분 (3 rectangles + 3 text boxes = 6)
# 4. Row 2: 목표 (3 + 3 = 6)
# 5. Row 3: 재고 전략 (3 + 3 = 6)
# 6. Row 4: 공급망 관리 (3 + 3 = 6)
# 7. Row 5: 리스크 대응 (3 + 3 = 6)
# 8. Border lines (12 lines: vertical × 2, horizontal × 6)

# Total: 3 + 3 + (6 × 5) + 12 = 48 shapes... (too many!)
# Simplify: Use table borders efficiently
# Actual: Header (3) + 5 rows × 3 cells (15) + text (18) = 36... still too many

# Better approach: Use python-pptx addTable (counts as 1 shape in PPTX)
# Then add decorative arrows/icons around it (5-10 shapes)
# Total: 1 table + 8 decorative elements = 9 shapes...

# WAIT - S4HANA analysis shows tables still create many shapes internally
# Let's use rectangles + text boxes as S4HANA does

# Final: 6 rows × 3 columns = 18 cell rectangles + 18 text boxes = 36 shapes ✓
```

---

## Font Size Distribution Strategy

**Target distribution (based on S4HANA data):**
- 8-10pt: **48%** (body text, descriptions)
- 12pt: **17%** (bullets, emphasis)
- 14pt: **8%** (section headers)
- 16pt: **7%** (governing messages)
- 20pt: **7%** (slide titles)
- Other: **13%** (cover, notes)

**Application per slide type:**

### Content slides (Slides 4-7, 9-10, 12-15, 17, 20-23)
```
- Slide title: 20pt bold (1 text box)
- Governing message: 16pt bold (1 text box)
- Section headers (h3): 14pt bold (2-3 text boxes)
- Bullet points: 12pt regular (4-6 items)
- Body text/descriptions: 9-10pt regular (10-15 text boxes)
- Small notes/captions: 8pt regular (2-3 text boxes)

Font distribution per slide:
- 8-10pt: ~60% (primary!)
- 12pt: ~20%
- 14pt: ~10%
- 16-20pt: ~10%
```

### Table/Matrix slides (Slides 6, 18, 22)
```
- Slide title: 20pt bold
- Governing message: 16pt bold
- Table headers: 12pt bold
- Table cells: 9-10pt regular (primary!)
- Footnotes: 8pt regular

Font distribution:
- 8-10pt: ~70% (even more!)
- 12pt: ~15%
- 16-20pt: ~15%
```

### Door Chart (Slide 11)
```
- Slide title: 20pt bold
- Governing message: 16pt bold
- Quadrant labels: 14pt bold (4 labels)
- Axis labels: 12pt bold (2 labels)
- Detail boxes: 9-10pt regular (40+ text boxes!)
- Spectrum indicators: 10pt regular (4 indicators)

Font distribution:
- 9-10pt: ~80% (maximum density!)
- 12-14pt: ~15%
- 16-20pt: ~5%
```

---

## Toy Page Layout Implementation

**Slides using Toy Page pattern:** 4, 5, 9, 10, 17, 20, 21, 24

**Structure:**
```html
<div class="row gap">
  <div style="flex: 0 0 65%; padding-right: 20px;">
    <!-- Left: Visual elements (diagrams, flowcharts, timelines) -->
    <!-- 30-50 shapes here -->
  </div>
  <div style="flex: 0 0 30%; padding-left: 10px;">
    <!-- Right: Text explanations -->
    <h3>시사점</h3>
    <p>... (9-10pt text)</p>
    <h3>방안</h3>
    <p>... (9-10pt text)</p>
  </div>
</div>
```

**Example: Slide 4 (JIT Timeline)**
- **Left 65%:** Timeline diagram (50 shapes)
- **Right 30%:** "시사점", "전환 배경", "영향" (3 sections, 9-10pt text)

---

## Governing Messages (ALL 25 Slides)

### 1. Cover
(No governing message)

### 2. TOC
"본 과정은 Kraljic Matrix 기반으로 자재군별 차별화 전략과 계획 방법론을 체계적으로 학습합니다."

### 3. Chapter 1 Divider
(No governing message)

### 4. 1.1 JIT의 영광과 몰락
"JIT 방식은 40년간 제조업의 표준이었으나 2020년 팬데믹으로 치명적 약점이 드러났습니다."

### 5. 1.2 팬데믹이 드러낸 약점
"글로벌 공급망 마비로 JIT의 3대 위험(재고 부족, 공급 중단, 생산 마비)이 현실화되었습니다."

### 6. 1.3 JIT vs JIC 비교
"JIT는 원가 절감에, JIC는 공급 안정성에 초점을 맞춰 서로 다른 리스크 환경에 대응합니다."

### 7. 1.4 JIC 채택 기업들
"팬데믹 이후 글로벌 제조사들은 JIC로 전환하여 안전재고와 다변화 전략을 채택했습니다."

### 8. Chapter 2 Divider
(No governing message)

### 9. 2.1 Kraljic Matrix 탄생
"1983년 Peter Kraljic이 개발한 2×2 매트릭스는 자재 특성에 따른 차별화 전략의 기초가 되었습니다."

### 10. 2.2 2×2 매트릭스의 두 축
"공급 리스크(X축)와 구매 임팩트(Y축) 두 기준으로 자재를 4개 군으로 분류합니다."

### 11. 2.3 Kraljic Matrix (Door Chart)
"Kraljic Matrix는 공급 리스크와 구매 금액을 기준으로 자재를 4개 군으로 분류하여 차별화 전략을 수립합니다."

### 12. 2.4 병목자재 (Bottleneck)
"병목자재는 공급 리스크가 높지만 구매 금액이 낮아 공급 안정성 확보가 최우선 과제입니다."

### 13. 2.5 레버리지자재 (Leverage)
"레버리지자재는 공급 리스크가 낮고 구매 금액이 높아 경쟁입찰과 가격협상으로 원가 절감을 추구합니다."

### 14. 2.6 전략자재 (Strategic)
"전략자재는 공급 리스크와 구매 금액이 모두 높아 장기 파트너십과 리스크 관리가 핵심입니다."

### 15. 2.7 일상자재 (Routine)
"일상자재는 공급 리스크와 구매 금액이 모두 낮아 프로세스 효율화와 자동화로 관리합니다."

### 16. Chapter 3 Divider
(No governing message)

### 17. 3.1 차별화의 필요성
"자재군 특성을 무시한 획일적 관리는 비효율과 리스크를 초래하며 차별화 전략이 필수입니다."

### 18. 3.2 자재군별 전략 매트릭스
"4개 자재군별로 소싱 전략, 재고 정책, 공급업체 관리 방식을 차별화하여 최적의 성과를 달성합니다."

### 19. Chapter 4 Divider
(No governing message)

### 20. 4.1 5대 방법론 개요
"ROP, MRP, LTP, Min-Max, VMI 등 5대 방법론을 자재 특성에 맞춰 선택하여 재고 효율을 극대화합니다."

### 21. 4.2 하이브리드 접근법
"전략자재는 예측 기반 LTP와 수요 기반 MRP를 결합한 하이브리드 방식으로 유연성을 확보합니다."

### 22. 5장: 통합 KPI 프레임워크
"원가, 서비스 수준, 재고 회전율, 공급 안정성 4대 KPI로 자재군별 성과를 측정하고 개선합니다."

### 23. 6장: 산업별 적용
"자동차, 전자, 화학, 식품, 건설 등 산업별 Kraljic Matrix 적용 사례와 베스트 프랙티스를 학습합니다."

### 24. 7장: 9회차 학습 여정
"9회차 과정을 통해 Kraljic 이론부터 실전 워크샵까지 단계적으로 학습하여 실무 적용 역량을 확보합니다."

### 25. Summary & Next Steps
"Kraljic Matrix 프레임워크와 차별화 전략을 학습했으며 다음 세션에서 소싱 전략과 공급업체 관리를 다룹니다."

---

## Section Structure and Numbering

**TOC Slide (Slide 2):**
```
1장 JIT → JIC 패러다임 전환
2장 Kraljic Matrix 프레임워크
3장 차별화 전략
4장 계획 방법론
5장 통합 KPI 프레임워크
6장 산업별 적용 사례
7장 9회차 학습 여정
```

**Slide Title Format:**
```
X.Y Topic Name

Examples:
- 1.1 JIT의 영광과 몰락
- 2.3 Kraljic Matrix (Door Chart)
- 4.2 하이브리드 접근법
```

**Chapter Divider Format:**
```
N장
Main Chapter Title

Examples:
- 1장
  JIT → JIC 패러다임 전환
- 2장
  Kraljic Matrix 프레임워크
```

---

## Storyline Approach: Structural (구조적 접근)

**Session 1 uses Structural approach** because we're building the Kraljic framework from foundation to application.

**Flow:**
1. **Context:** JIT → JIC paradigm shift (Slides 4-7)
2. **Framework Introduction:** Kraljic Matrix origin and structure (Slides 9-11)
3. **Component Details:** Four material quadrants (Slides 12-15)
4. **Application:** Differentiation strategies (Slides 17-18)
5. **Methodologies:** Planning approaches (Slides 20-21)
6. **Integration:** KPI, industry, learning journey (Slides 22-24)
7. **Conclusion:** Summary and next steps (Slide 25)

**Persuasive elements:**
- **Governing messages** penetrate listener's mind (insight, not just topic)
- **Evidence-based:** Use real company examples, pandemic data, industry statistics
- **Progressive complexity:** Foundation → Framework → Details → Application
- **Visual impact:** Door charts, timelines, comparison matrices

---

## Quality Gates (MUST PASS BEFORE PROCEEDING)

### Gate 1: S4HANA Analysis ✓
- [x] Analyzed ≥10 slides structure
- [x] Identified shape counts: 50-120 shapes for complex slides
- [x] Identified font usage: 9-10pt = 40% (primary)
- [x] Confirmed dimensions: 10.83" × 7.50"

### Gate 2: Design Plan ✓
- [x] Documented plan exists (this file!)
- [x] Shape targets: 725-950 total (29-38 avg)
- [x] Font distribution: 48% in 8-10pt range
- [x] Toy Page layouts: 8 slides planned
- [x] Governing messages: All 25 drafted
- [x] Storyline approach: Structural confirmed

### Gate 3: Code Review (NEXT STEP)
- [ ] All 6 checklist items from CLAUDE.md verified
- [ ] python-pptx code implements shape variety
- [ ] Rectangles, arrows, connectors, lines planned
- [ ] WHITE text on dark backgrounds enforced
- [ ] Groups planned (where applicable)
- [ ] Font sizes: 48% in 8-10pt range

### Gate 4: Post-Generation Verification
- [ ] Verification script passes all checks
- [ ] Slide count: 25 ✓
- [ ] Dimensions: 10.83" × 7.50" ✓
- [ ] Shape counts: ≥20 per content slide
- [ ] Font distribution: 40-50% in 8-10pt range

---

## Implementation Checklist

### Before Coding:
- [x] Read all documentation (SKILL.md, html2pptx.md, css.md, references)
- [x] Run S4HANA analysis
- [x] Document design plan
- [ ] Review python-pptx code structure
- [ ] Plan shape generation functions (arrows, boxes, connectors)

### During Coding:
- [ ] Use 9-10pt as PRIMARY font size (48% of text)
- [ ] Add 40-100 shapes per complex slide
- [ ] Use LINES for arrows and connectors
- [ ] Use AUTO_SHAPES for rectangles and boxes
- [ ] WHITE text on dark backgrounds
- [ ] 16pt BOLD for governing messages (not 14pt italic!)
- [ ] Create door chart for Kraljic Matrix (75-100 shapes)

### After Generation:
- [ ] Run verification script
- [ ] Manual spot-check 5 slides
- [ ] Confirm shape counts meet targets
- [ ] Confirm font distribution meets targets
- [ ] Confirm consistency with future Part 2-9

---

## Next Step: Code Implementation

**Ready to proceed?** YES ✓

**Estimated implementation time:** 3-4 hours (for high quality)

**Approach:**
1. Create helper functions for shape generation
   - `add_rectangle(slide, x, y, w, h, fill, text, font_size)`
   - `add_arrow(slide, x1, y1, x2, y2, color)`
   - `add_timeline(slide, y, events)` - generates 40-50 shapes
   - `add_door_chart_matrix(slide)` - generates 75-100 shapes
   - `add_comparison_boxes(slide, left_content, right_content)`

2. Implement slides systematically
   - Cover (slide 1)
   - TOC (slide 2) - 15-20 shapes
   - Each content slide with proper shape counts
   - Door chart (slide 11) - MOST CRITICAL

3. Test and verify
   - Run verification after every 5 slides
   - Adjust if shape counts too low
   - Confirm font sizes correct

**Start implementation now.**
