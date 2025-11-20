# CLAUDE.md - AI Assistant Guide for Kraljic_Course Repository

## Repository Overview

This repository contains a comprehensive Korean-language educational course on **Strategic Inventory Management and Material Planning** using the Kraljic Matrix framework. The course addresses the paradigm shift from Just-In-Time (JIT) to Just-In-Case (JIC) inventory management and provides practical frameworks for material categorization and planning.

### Project Purpose
- Educational content for supply chain management professionals
- Practical training on the Kraljic Matrix methodology
- Strategic inventory management and material planning techniques
- Supplier relationship management and performance evaluation

### Language
- **Primary Language**: Korean (한국어)
- All course content, documentation, and data files are in Korean
- File names and directory names use Korean characters

---

## Repository Structure

```
Kraljic_Course/
├── README.md                                    # Repository overview and course guide
├── CLAUDE.md                                    # This file - AI assistant guide
├── .gitignore                                   # Git ignore patterns
├── Kraljic_Course_Contents.zip                 # Original archive
├── ExportBlock-*.zip                            # Extracted course content archive
├── 전략적 재고운영 및 자재계획수립.csv         # Course curriculum index
└── 전략적 재고운영 및 자재계획수립/             # Main course directory
    ├── [1회차] 전략적 재고운영 Foundation...md  # Session 1: Foundation & Kraljic Matrix
    ├── [2회차] 자재군별 소싱 전략...md           # Session 2: Sourcing strategies
    │   └── 공급업체 성과 평가/                  # Supplier scorecard data
    │       ├── *.csv                             # Scorecard CSV files
    │       └── [공급업체명]/*.md                # Individual supplier profiles (10 suppliers)
    ├── [3회차] ABC-XYZ 재고 분류...md           # Session 3: ABC-XYZ analysis
    ├── [4회차] 병목자재 전략 & ROP.md           # Session 4: Bottleneck materials & ROP
    ├── [5회차] 레버리지자재 전략 & MRP.md       # Session 5: Leverage materials & MRP
    ├── [6회차] 전략자재 전략 & 하이브리드...md  # Session 6: Strategic materials
    ├── [7회차] 일상자재 효율화 & 자동화.md      # Session 7: Routine materials
    ├── [8회차] Kraljic Matrix 실전 워크샵.md   # Session 8: Practical workshop
    └── [9회차] 통합 워크샵...md                 # Session 9: Integrated workshop
```

---

## Course Structure

### 9-Session Curriculum

| Session | Topic | Category | Importance | Difficulty | Duration |
|---------|-------|----------|------------|------------|----------|
| 1회차 | Kraljic Matrix Foundation & Methodology | Overview | High | Intermediate | 45 min |
| 2회차 | Sourcing Strategy & Supplier Management | Overview | High | Intermediate | 45 min |
| 3회차 | ABC-XYZ Inventory Classification | Overview | High | Intermediate | 45 min |
| 4회차 | Bottleneck Materials & ROP | Bottleneck | High | Intermediate | 45 min |
| 5회차 | Leverage Materials & MRP | Leverage | High | Intermediate | 45 min |
| 6회차 | Strategic Materials & Hybrid Planning | Strategic | High | Advanced | 45 min |
| 7회차 | Routine Materials Efficiency & Automation | Routine | Low | Beginner | 45 min |
| 8회차 | Kraljic Matrix Practical Workshop | Workshop | High | Intermediate | 45 min |
| 9회차 | Integrated Workshop: Real-world Application | Workshop | High | Advanced | 45 min |

### Core Concepts Covered

#### 1. Kraljic Matrix Framework
The Kraljic Matrix categorizes materials into 4 quadrants based on:
- **X-axis**: Supply Risk (공급 리스크)
- **Y-axis**: Purchase Amount/Strategic Impact (구매 금액)

**Four Material Categories**:
1. **전략자재 (Strategic Materials)**: High risk, high impact
2. **레버리지자재 (Leverage Materials)**: Low risk, high impact
3. **병목자재 (Bottleneck Materials)**: High risk, low impact
4. **일상자재 (Routine Materials)**: Low risk, low impact

#### 2. Planning Methodologies
- **ROP (Re-Order Point)**: For bottleneck materials
- **MRP (Material Requirements Planning)**: For leverage materials
- **Hybrid Planning**: For strategic materials
- **Automation**: For routine materials

#### 3. ABC-XYZ Analysis
- **ABC**: Classification by value/amount
- **XYZ**: Classification by demand variability
- **Matrix**: 9 combinations for operational segmentation

---

## PPTX Generation Guidelines (CRITICAL - READ BEFORE ANY PPTX WORK)

### ⚠️ Mandatory Prerequisites

**BEFORE generating any PPTX files, AI assistants MUST:**

1. **Read the complete Skill documentation** (no line limits):
   - `/home/user/Kraljic_Course/skill/SKILL.md` (complete file, ~800 lines)
   - `/home/user/Kraljic_Course/skill/html2pptx.md` (complete file)
   - `/home/user/Kraljic_Course/skill/css.md` (complete file)
   - **DO NOT** use offset or limit parameters when reading these files

1-b. **Read the complete Reference guidelines** (CRITICAL - no line limits):
   - `/home/user/Kraljic_Course/skill/references/design-guidelines.md` (complete file, 567 lines)
     - Contains: Font size hierarchy, Shape count targets, Door chart pattern, Storyline approaches
     - Section 8: Shape Count & Visual Density Requirements (20-50+ shapes per slide)
     - Section 9: Persuasive Storyline Development (Structural, Dynamics, Market Change)
   - `/home/user/Kraljic_Course/skill/references/design-patterns-findings.md` (complete file, 350 lines)
     - Contains: Professional analysis data (10pt font = 65.2%, 75 shapes in door charts)
     - Actual statistics from 추가자료2_Inspiration_2024.pptx and RedSlide materials
   - `/home/user/Kraljic_Course/skill/references/DESIGN_ENHANCEMENT_REPORT.md` (complete file, 560 lines)
     - Contains: Executive summary, Session-specific recommendations, Implementation checklist
   - **DO NOT** use offset or limit parameters - Read these files COMPLETELY
   - **These contain the professional standards** that prevented Part 1 quality issues

2. **Analyze the reference PPTX file thoroughly**:
   - Path: `/home/user/Kraljic_Course/PPTX_SAMPLE/S4HANA_PI단계_단계 종료보고_20230510_v.1.4.pptx`
   - This is the **gold standard** for style, tone, and formatting
   - Extract slides and analyze: dimensions, colors, fonts, layout patterns
   - Use `python-pptx` library to inspect properties

3. **Understand and apply the pptx-mslee skill workflow**:
   - Workflow: **JSON → HTML → PPTX** (not direct python-pptx coding)
   - Use Handlebars templates in `skill/templates/education-course/layouts/`
   - Generate HTML first, then convert to PPTX using `@ant/html2pptx`

### 🎨 S4HANA Design System (MANDATORY)

The reference file uses **S4HANA monochrome design principles**:

#### Color System (STRICT Monochrome)
**CRITICAL: Use ONLY monochrome colors (black/white/gray) for ALL slides**

- **Standard palette** (use for 99% of slides):
  - Black (#000000): Primary text, key emphasis
  - Dark Gray (#333333): Secondary text, headers
  - Medium Gray (#666666): Normal text
  - Light Gray (#CCCCCC): Backgrounds, borders
  - Very Light Gray (#E6E6E6): Subtle backgrounds
  - White (#FFFFFF): White backgrounds, reverse text
  - Dark Blue (#1A5276): Accent color (MINIMAL use only)

- **Kraljic colors**: Use ONLY in Kraljic Matrix 2×2 diagram slide
  - Strategic: Purple (#8E44AD) - ONE slide only
  - Bottleneck: Orange (#E67E22) - ONE slide only
  - Leverage: Green (#27AE60) - ONE slide only
  - Routine: Gray (#95A5A6) - ONE slide only
  - **DO NOT** use these colors in any other slides!

- **Forbidden EVERYWHERE ELSE**: Rainbow colors, multiple bright colors, gradients
- **Rule**: If not Matrix diagram → Use ONLY black/white/gray

#### Typography
- **Title font**: Arial (English), 맑은 고딕 (Korean)
- **Body font**: 맑은 고딕 (Korean), Arial (English)
- **Font sizes** (actual usage from S4HANA and professional samples):
  - **48pt**: Cover slide main title (Bold)
  - **20pt**: Content slide titles (Bold)
  - **16pt**: Governing messages (Bold)
  - **14pt**: Section headers, large bullet points
  - **12-13pt**: Regular bullet points (20-25% of text)
  - **10-11pt**: Body text, descriptions (**PRIMARY - 60-70% of all text**)
  - **8-9pt**: Small annotations, footnotes
  - **6-7pt**: Tiny notes (rare)
- **Font weights**: Bold for titles/headers, Regular for body
- **CRITICAL insight**: Professional analysis shows **10pt is THE dominant body text size (65.2% of all text)**. This enables high content density (85%+) while maintaining readability. Don't use 16-18pt for body text - that's too large and wastes space.

- **Text color rules** (CRITICAL for readability):
  - **Dark backgrounds** (Dark Gray, Med Gray, Black) → **White text (#FFFFFF)**
  - **Light backgrounds** (Light Gray, Very Light Gray, White) → **Black/Dark Gray text (#000000, #333333)**
  - **Rule**: Always maintain high contrast between text and background
  - Examples:
    - Dark Gray box (#333333) → White text (#FFFFFF)
    - Light Gray box (#CCCCCC) → Black text (#000000)
    - Medium Gray box (#666666) → White text (#FFFFFF)

#### Slide Dimensions
- **Width**: 10.83 inches
- **Height**: 7.5 inches
- **Aspect ratio**: ~1.44:1 (not 16:9!)

#### Layout Principles
- **White background**: Default for all content slides (cover slide can use color)

- **Grid System (MANDATORY)**: All elements MUST align to grid
  - **2-column layout**: x = [0.8", 5.5"] (width: 4.5" each)
  - **3-column layout**: x = [1.0", 4.2", 7.4"] (width: 3.0" each)
  - **4-column layout**: x = [0.8", 3.2", 5.6", 8.0"] (width: 2.2" each)
  - **Row spacing**: 0.8-1.0" between rows
  - **NO random positioning**: Every box must snap to grid

- **Content density**: CRITICAL - Slides must use **85%+ of slide area**
  - S4HANA average: 83.4% (median: 75.5%)
  - Many slides exceed 100% density due to overlapping elements
  - **Minimize whitespace** - Use small fonts (8-11pt) to fit more content
  - Example: Slide 4 has 26 AUTO_SHAPES + 7 text boxes = 84.6% density

- **Visual elements**: Use extensive diagrams, flowcharts, and shapes
  - **Shapes per slide**: 10-50+ AUTO_SHAPES (rectangles, arrows, connectors)
  - **Shape variety (CRITICAL)**:
    - Rectangles: Wrap ALL text content (no floating text!)
    - Arrows: Show time sequence (Before → After), process flow (Step 1 → Step 2)
    - Triangles: Indicate increase/decrease, priorities
    - Rounded rectangles: Emphasize key points
    - Connectors: Show relationships between concepts
  - **Flowcharts**: Timeline diagrams, process flows with arrows
  - **Tables**: Data grids, comparison matrices
  - **Groups**: Organize related shapes into logical groups
  - Example: Slide 4 has timeline with phases, arrows, and 20+ detail boxes

- **Structuring with shapes**:
  - Every text block → wrapped in rectangle box
  - Alternate background colors: Light Gray ↔ Very Light Gray ↔ White
  - Use borders (0.75-1pt) to separate sections
  - Comparisons: Side-by-side boxes with arrow between
  - Sequences: Boxes in row with arrows connecting

- **Toy Page Layout (PRIMARY PATTERN - CRITICAL)**:
  - **MOST content slides should use this layout**
  - **Left side (60-70% of slide width)**: Visual elements
    - Diagrams, flowcharts, process flows
    - Timelines with arrows
    - Comparison matrices
    - Structured shapes and boxes
    - Charts, graphs, illustrations
    - Position: x = 0.8", width = ~6.5-7.5"
  - **Right side (30-40% of slide width)**: Text explanations
    - 시사점 (Implications)
    - 방안 (Solutions/Approaches)
    - 상세설명 (Detailed explanations)
    - Key takeaways, insights
    - Position: x = ~7.5-8.0", width = ~2.5-3.0"
  - **Examples of Toy Page slides**:
    - Timeline (left) → Key insights (right)
    - Process flow diagram (left) → Implementation steps (right)
    - Comparison matrix (left) → Strategic recommendations (right)
  - **Benefits**: High visual impact + Clear narrative structure

- **Table of Contents & Section Structure (MANDATORY)**:
  - **TOC slide at beginning**: Create clear chapter structure
    - Format: "1장 Title", "2장 Title", "3장 Title"
    - Show complete course outline with chapter numbers
    - Use clean, grid-aligned layout
  - **Section numbering in slide titles**:
    - Format: "X.Y Topic Name" where X = chapter, Y = slide in chapter
    - Example: "2.3 JIT의 7가지 원칙" (3rd slide in Chapter 2)
    - Example: "4.1 Kraljic Matrix 개요" (1st slide in Chapter 4)
  - **Clear navigation**: User should always know current location
    - Which chapter they're in
    - Which topic within that chapter
    - How it fits in the overall structure
  - **Chapter dividers**: Use section break slides between chapters
    - Format: Large "N장" with chapter title
    - Minimal design, high visual impact

- **Governing messages**: REQUIRED for all content slides
  - One-sentence summary under the title
  - Position: (0.30", 1.01"), Size: 10.32" × 0.63"
  - 16pt Bold 맑은 고딕 (NOT 14pt Italic)
  - Captures the "so what" of the slide

### 📋 Governing Message Pattern

**Every content slide MUST include a governing message**:

```html
<div class="title-section fit">
  <h1>Slide Title</h1>
  <p class="governing-message">One-sentence summary that captures the entire slide's key point.</p>
</div>
```

**Examples of good governing messages**:
- ✅ "JIT 방식은 2020년 팬데믹으로 치명적 약점이 드러났고, 기업들은 JIC로 전환하고 있습니다."
- ✅ "Kraljic Matrix는 공급 리스크와 구매 임팩트 두 축으로 자재를 4개 군으로 분류합니다."
- ❌ "이 슬라이드는 JIT와 JIC를 비교합니다." (Too vague)
- ❌ (No governing message) (Missing!)

### 🔧 Technical Workflow

**Correct approach** (using pptx-mslee skill):
1. Create JSON data file in `skill/data/{session-name}.json`
2. Create/use Handlebars templates in `skill/templates/education-course/layouts/`
3. Run: `node scripts/generate-course.js data/{session-name}.json`
4. Output: `skill/output/{session-name}.pptx`

**Fallback approach** (if html2pptx fails):
- Use `python-pptx` library with **strict adherence to S4HANA design system**
- Replicate the reference file's style exactly
- Include governing messages in code
- Apply monochrome color scheme

### 📁 Reference Files

- **Style reference**: `PPTX_SAMPLE/S4HANA_PI단계_단계 종료보고_20230510_v.1.4.pptx`
- **Skill documentation**: `skill/SKILL.md`, `skill/html2pptx.md`, `skill/css.md`
- **Templates**: `skill/templates/education-course/layouts/*.hbs`
- **Partials**: `skill/templates/education-course/partials/*.hbs`
- **Styles**: `skill/templates/education-course/styles/*.css`

### ❌ Common Mistakes to Avoid

1. **Using colorful designs**: S4HANA is monochrome!
2. **Skipping governing messages**: They are REQUIRED
3. **Not reading SKILL.md completely**: Read the entire file, no limits
4. **Not reading skill/references/ guidelines**: The 3 reference files (design-guidelines.md, design-patterns-findings.md, DESIGN_ENHANCEMENT_REPORT.md) contain critical professional standards - MUST read all 1,477 lines!
5. **Using wrong dimensions**: Must be 10.83" × 7.5"
6. **Ignoring the reference file**: It's the gold standard
7. **Direct python-pptx coding without skill workflow**: Try skill approach first
8. **Using gradients on cover slide**: Use solid colors (gradient causes rendering issues)
9. **Using too large fonts**: Don't use 16-18pt for body text - use 10pt! (65% of all text should be 10pt)
10. **Too much whitespace**: Slides must be 85%+ filled - add more content, diagrams, shapes
11. **Missing visual elements**: No flowcharts, arrows, or shapes - Professional slides use 20-50+ shapes per slide
12. **Not using GROUPS**: 70-80% of shapes should be in groups for organization - don't just scatter individual shapes
13. **No door charts for matrices**: Kraljic Matrix and spectrum visualizations need the door chart pattern (75+ shapes)
14. **Missing storyline approach**: Slides lack coherent flow - choose Structural, Dynamics, or Market Change approach
15. **Misunderstanding Part numbers**: "Part 1" = Session 1 only, NOT Sessions 1-3
16. **Poor text contrast**: Using dark text on dark backgrounds or light text on light backgrounds - Always use white text on dark backgrounds!
17. **Not using Toy Page layout**: Most content slides should use 60-70% visual (left) + 30-40% text (right) structure
18. **Missing section structure**: No TOC slide, no section numbers in titles (e.g., "2.3"), unclear navigation
19. **Weak governing messages**: Messages just describe topic instead of providing insight that "penetrates the listener's mind"
20. **Ignoring checklist items**: Reading checklist but not actually verifying each item before generation
21. **Superficial reference analysis**: Extracting only colors/fonts from S4HANA without analyzing actual slide structure, shape counts, layout patterns

---

## 🚨 CRITICAL: Preventing Quality Failures (Part 1-9 Consistency)

### Why This Section Exists

Part 1 초기 생성에서 발생한 문제:
- 체크리스트를 읽기만 하고 실제로 검증하지 않음
- python-pptx fallback 사용 시 모든 디자인 요구사항을 무시함
- S4HANA 참고 파일을 색상/폰트만 추출하고 구조 분석 안 함
- 결과: 텍스트박스만 있는 저품질 슬라이드 (content density 30-40%, shapes < 5개/슬라이드)

**Part 1-9까지 일관성이 중요**: 한 Part만 품질이 다르면 전체 과정의 신뢰도 하락

### Mandatory Pre-Generation Steps (절대 생략 불가)

#### Step 1: S4HANA Reference Deep Analysis (30분 소요)

**단순히 색상/폰트만 추출하는 것이 아니라, 실제 슬라이드 구조를 분석해야 함**

```python
# 필수 실행 스크립트
python3 -c "
from pptx import Presentation
prs = Presentation('PPTX_SAMPLE/S4HANA_PI단계_단계 종료보고_20230510_v.1.4.pptx')

print('=== S4HANA Slide Structure Analysis ===')
for i, slide in enumerate(prs.slides[:10], 1):
    shapes = len(slide.shapes)
    auto_shapes = sum(1 for s in slide.shapes if str(s.shape_type) == 'AUTO_SHAPE (1)')
    text_boxes = sum(1 for s in slide.shapes if hasattr(s, 'text') and s.text.strip())
    groups = sum(1 for s in slide.shapes if str(s.shape_type) == 'GROUP (6)')

    print(f'\nSlide {i}:')
    print(f'  Total shapes: {shapes}')
    print(f'  AUTO_SHAPES: {auto_shapes}')
    print(f'  Text boxes: {text_boxes}')
    print(f'  Groups: {groups}')
    print(f'  Density estimate: {(shapes * 2)}%')  # Rough estimate
"
```

**분석 결과 예시** (실제 S4HANA):
```
Slide 4: 56 shapes (26 AUTO_SHAPES, 7 text boxes, density ~84%)
Slide 12: 102 shapes (87 AUTO_SHAPES, density ~100%+)
```

**⚠️ 이 분석 없이 생성 시작하면 안 됨!**

#### Step 2: Design Implementation Plan (필수 문서화)

생성 시작 전에 다음을 명시적으로 계획하고 문서화:

```markdown
## Part N Design Plan

### Slide Density Targets
- Target: 85%+ per slide
- Strategy: [구체적으로 어떻게 달성할 것인가]
  - Example: "Timeline slides: 20-30 shapes (arrows + boxes + connectors)"
  - Example: "Comparison slides: 15-20 shapes (rectangles + arrows)"

### Shape Usage Plan
- Total shapes per slide: [minimum 20개]
- Shape types to use:
  - Rectangles: [용도]
  - Arrows: [용도]
  - Triangles: [용도]
  - Connectors: [용도]
  - Groups: [70-80% of shapes grouped]

### Toy Page Layout Implementation
- Slides using Toy Page: [슬라이드 번호 리스트]
- Left side (60-70%): [구체적 비주얼 요소]
- Right side (30-40%): [구체적 텍스트 내용]

### Governing Messages
- [각 슬라이드별로 governing message 초안 작성]
- Verification: "Does it penetrate the listener's mind?"
```

**⚠️ 이 문서 없이 코딩 시작하면 안 됨!**

#### Step 3: Template/Code Review (코드 작성 후)

**python-pptx fallback 사용 시에도 다음을 반드시 구현해야 함**:

```python
# ✅ REQUIRED Checklist for python-pptx code

# 1. Slide dimensions
prs.slide_width = Inches(10.83)  # NOT 10.0!
prs.slide_height = Inches(7.5)

# 2. Governing messages (16pt Bold, NOT 14pt Italic)
gov_box = slide.shapes.add_textbox(...)
gov_frame.paragraphs[0].font.size = Pt(16)  # NOT 14!
gov_frame.paragraphs[0].font.bold = True    # NOT italic!

# 3. Shape variety (minimum 20 per slide)
# - Must include: rectangles, arrows, connectors, groups
# - Example:
arrow = slide.shapes.add_connector(
    MSO_CONNECTOR.STRAIGHT,
    Inches(2.0), Inches(3.0),  # Start
    Inches(4.0), Inches(3.0)   # End
)
arrow.line.color.rgb = COLOR_DARK_GRAY
arrow.line.width = Pt(2)

# 4. Text on dark backgrounds = WHITE color
# CRITICAL: Check every text element
text_frame.paragraphs[0].font.color.rgb = COLOR_WHITE  # if background is dark

# 5. Groups (70-80% of shapes)
# Group related shapes together
shapes_to_group = [shape1, shape2, shape3]
# Note: python-pptx doesn't support grouping easily - document this limitation

# 6. Font size distribution
# 65% of text: 10pt (PRIMARY)
# 20-25% of text: 12pt (bullets)
# Rest: 8pt (captions), 14pt (headings)
```

### Mandatory Post-Generation Verification (생성 즉시 실행)

```python
# 필수 검증 스크립트 (생성된 PPTX 파일에 대해 실행)
python3 -c "
from pptx import Presentation
import sys

prs = Presentation('Part1_Session1_StrategicInventory.pptx')
failures = []

# Check 1: Slide dimensions
if prs.slide_width != 914400 * 10.83:
    failures.append(f'❌ Width: {prs.slide_width/914400:.2f}\" (should be 10.83\")')
if prs.slide_height != 914400 * 7.5:
    failures.append(f'❌ Height: {prs.slide_height/914400:.2f}\" (should be 7.5\")')

# Check 2: Slide count
if len(prs.slides) < 20:
    failures.append(f'❌ Only {len(prs.slides)} slides (expected 20+)')

# Check 3: Shapes per slide
low_density_slides = []
for i, slide in enumerate(prs.slides[1:], 2):  # Skip cover
    if len(slide.shapes) < 10:
        low_density_slides.append(f'Slide {i}: {len(slide.shapes)} shapes')

if low_density_slides:
    failures.append(f'❌ Low shape count:\n  ' + '\n  '.join(low_density_slides[:5]))

# Check 4: Font sizes (sample check)
font_sizes = {}
for slide in prs.slides[:5]:
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame'):
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        size = int(run.font.size.pt)
                        font_sizes[size] = font_sizes.get(size, 0) + 1

total_text = sum(font_sizes.values())
pt10_ratio = font_sizes.get(10, 0) / total_text if total_text > 0 else 0
if pt10_ratio < 0.4:  # Should be 65% but allow some tolerance
    failures.append(f'❌ 10pt text ratio: {pt10_ratio*100:.1f}% (should be 60%+)')

print('\\n=== PPTX Quality Verification ===')
if failures:
    print('\\n'.join(failures))
    print(f'\\n🚫 FAILED {len(failures)} checks - DO NOT PROCEED')
    sys.exit(1)
else:
    print('✅ All checks passed')
    print(f'   Slides: {len(prs.slides)}')
    print(f'   Dimensions: {prs.slide_width/914400:.2f}\" × {prs.slide_height/914400:.2f}\"')
"
```

**⚠️ 이 검증 통과 못하면 수정 후 재검증!**

### Quality Gates (각 단계별 통과 기준)

| Stage | Gate | Pass Criteria | Fail Action |
|-------|------|---------------|-------------|
| **Pre-Gen** | S4HANA Analysis | Analyzed ≥10 slides structure | STOP - Run analysis script |
| **Pre-Gen** | Design Plan | Documented plan exists | STOP - Write plan first |
| **Pre-Gen** | Code Review | All 6 checklist items ✅ | STOP - Fix code |
| **Post-Gen** | Verification Script | All checks pass | STOP - Fix and regenerate |
| **Post-Gen** | Manual Review | Spot-check 5 slides | STOP - Identify issues |

**⚠️ 어느 gate라도 실패하면 다음 단계로 진행 금지!**

### Common Failure Patterns (실제 발생한 문제들)

#### Pattern 1: "빠르게 완성" 마인드
- **증상**: 체크리스트 읽고 바로 코딩 시작
- **결과**: 텍스트박스만 있는 저품질 슬라이드
- **해결**: Pre-Generation Steps 강제 실행

#### Pattern 2: "일단 돌아가게" 구현
- **증상**: python-pptx fallback에서 최소한만 구현
- **결과**: Shapes < 5개/슬라이드, governing messages 누락
- **해결**: Code Review Checklist 강제 검증

#### Pattern 3: "피상적 참고"
- **증상**: S4HANA에서 색상만 추출
- **결과**: 구조, 레이아웃, 밀도 무시
- **해결**: Deep Analysis Script 강제 실행

#### Pattern 4: "검증 생략"
- **증상**: 생성 후 바로 커밋
- **결과**: 품질 문제 발견 못함
- **해결**: Verification Script 강제 실행

### Part 1-9 Consistency Enforcement

**모든 Part는 동일한 품질 기준을 충족해야 함**:

```bash
# Part 1-9 공통 검증 스크립트
for part in Part{1..9}_*.pptx; do
    echo "Verifying $part..."
    python3 verify_pptx_quality.py "$part"
    if [ $? -ne 0 ]; then
        echo "❌ $part failed quality check"
        exit 1
    fi
done

echo "✅ All Parts passed quality checks"
```

**Consistency Checklist** (Part 간 일관성):
- [ ] 동일한 슬라이드 크기 (10.83" × 7.5")
- [ ] 동일한 색상 팔레트 (monochrome + Kraljic)
- [ ] 동일한 폰트 크기 분포 (10pt 65%, 12pt 20-25%)
- [ ] 동일한 governing message 스타일 (16pt Bold)
- [ ] 동일한 shape 밀도 (20-50+ per slide)
- [ ] 동일한 레이아웃 패턴 (Toy Page, 2-col, etc.)

---

### ✅ Checklist Before Generating PPTX (Updated with Mandatory Gates)

#### Phase 1: Documentation Review (READ ONLY - MANDATORY)
- [ ] Read complete SKILL.md (no offset/limit)
- [ ] Read complete html2pptx.md
- [ ] Read complete css.md
- [ ] **MANDATORY**: Read complete skill/references/design-guidelines.md (567 lines)
  - Section 8: Shape Count & Visual Density Requirements
  - Section 9: Persuasive Storyline Development (Structural, Dynamics, Market Change)
  - Quality Checklist (complete)
- [ ] **MANDATORY**: Read complete skill/references/design-patterns-findings.md (350 lines)
  - Professional data: 10pt font = 65.2% usage
  - Door chart pattern: 75 shapes, 70-80% in groups
- [ ] **MANDATORY**: Read complete skill/references/DESIGN_ENHANCEMENT_REPORT.md (560 lines)
  - Session-specific recommendations
  - Implementation checklist
- [ ] Read "CRITICAL: Preventing Quality Failures" section above

#### Phase 2: Pre-Generation Analysis (MUST DO - 30 min)
- [ ] **MANDATORY**: Run S4HANA Deep Analysis script
  - Analyze ≥10 slides structure (shapes, AUTO_SHAPES, text boxes, groups)
  - Document findings: average shapes per slide, density estimates
  - Identify layout patterns used in reference
- [ ] **MANDATORY**: Create Design Implementation Plan document
  - Slide density targets (85%+ strategy)
  - Shape usage plan (minimum 20 per slide, types & purposes)
  - Toy Page layout implementation list
  - Governing messages draft for ALL slides
- [ ] Understood monochrome color system (black/white/gray only, Kraljic exception)
- [ ] Understood font size hierarchy (10pt PRIMARY 65%, 12pt bullets 20-25%)
- [ ] Understood text color rules (WHITE on dark, BLACK on light - CRITICAL)

#### Phase 3: Design Planning (MUST DOCUMENT)
- [ ] Planned content density strategy to achieve 85%+ (written in plan)
- [ ] Designed flowcharts, diagrams, shapes (20-50+ per slide minimum)
- [ ] Planned shape variety: rectangles, arrows, triangles, connectors
- [ ] Planned GROUP organization (70-80% of shapes grouped)
- [ ] Designed door charts for Kraljic Matrix (75+ shapes)
- [ ] Chosen storyline approach (Structural, Dynamics, or Market Change)
- [ ] Drafted governing messages (16pt Bold, insightful, "penetrate listener's mind")
- [ ] Designed Toy Page layouts (list slides: 60-70% visual left, 30-40% text right)

#### Phase 4: Implementation Preparation
- [ ] Created TOC slide with chapter structure (1장, 2장...)
- [ ] Applied section numbering to ALL slide titles (X.Y format)
- [ ] Prepared JSON data structure OR python-pptx code
- [ ] If using templates: Verified Handlebars templates exist
- [ ] If using python-pptx: Reviewed code against 6-item checklist
- [ ] Confirmed slide dimensions: 10.83" × 7.5"
- [ ] Confirmed Part/Session mapping (Part N = Session N only)

#### Phase 5: Quality Gates (STOP if fail)
- [ ] **GATE 1**: S4HANA Analysis complete? (YES/NO) - STOP if NO
- [ ] **GATE 2**: Design Plan documented? (YES/NO) - STOP if NO
- [ ] **GATE 3**: Code reviewed against checklist? (YES/NO) - STOP if NO

#### Phase 6: Post-Generation Verification (MUST RUN)
- [ ] **MANDATORY**: Run verification script immediately after generation
- [ ] Verification passed all checks? (YES/NO) - STOP & FIX if NO
- [ ] Manual spot-check 5 slides for visual quality
- [ ] Confirmed consistency with previous Parts (if Part 2+)

---

## File Conventions

### Naming Patterns

1. **Session Files**: `[N회차] {Topic Title} {Hash}.md`
   - N: Session number (1-9)
   - Hash: Unique identifier (32 characters)
   - Example: `[1회차] 전략적 재고운영 Foundation Kraljic Matrix와 자재계획 방법론 28287a1932c4811b9e53cae79af30fa8.md`

2. **Data Files**:
   - CSV format with Korean headers
   - Two versions: `{name}.csv` and `{name}_all.csv`
   - UTF-8 encoding with BOM (﻿)

3. **Directory Structure**:
   - Korean characters for all directory names
   - Nested structure for hierarchical content
   - Supplier data organized by scorecard type

### File Content Structure

#### Markdown Files
```markdown
# [Session] Title

단계: {Category}
중요도: {Importance Level}
난이도: {Difficulty Level}
Min: {Duration}
No: {Session Number}

---

<aside>
🎯
**학습 목표** (Learning Objectives)
- Bullet points...
</aside>

## Sections...
```

#### CSV Files
- Headers in Korean
- Comma-separated values
- Date format: `YYYY년 MM월 DD일 오후/오전 HH:MM`
- Percentage values with % symbol
- Decimal separator: period (.)

---

## Data Schema

### Course Curriculum CSV
```csv
No, 교육 주제, 단계, 중요도, 난이도, 교육 자료, Min
```

**Fields**:
- `No`: Session number
- `교육 주제`: Course topic
- `단계`: Stage/category
- `중요도`: Importance (높음/낮음)
- `난이도`: Difficulty (초급/중급/고급)
- `교육 자료`: Training materials
- `Min`: Duration in minutes

### Supplier Scorecard CSV
```csv
공급업체명, 가격 안정성, 가격경쟁력 점수, 개선제안 건수, 검사통과율, ...
```

**Key Fields**:
- `공급업체명`: Supplier name
- `자재군`: Material category (전략/레버리지/병목/일상)
- `등급`: Grade (A/B/C/D)
- `총점`: Total score
- `납기준수율 OTD`: On-time delivery rate
- `품질 점수`: Quality score
- `협력성과 점수`: Collaboration performance score

**10 Suppliers in Dataset**:
1. 미래금속 (B - Strategic materials)
2. 동양플라스틱 (B - Routine materials)
3. 아시아MRO (C - Routine materials)
4. 중앙산업 (D - Bottleneck materials)
5. 글로벌스틸 (B - Leverage materials)
6. 대한전자부품 (A - Leverage materials)
7. 삼성화학 (B - Bottleneck materials)
8. 신한부품 (C - Leverage materials)
9. (주)한국정밀 (A - Strategic materials)
10. 태평양소재 (C - Bottleneck materials)

---

## AI Assistant Guidelines

### When Working with This Repository

#### 1. Language Handling
- **DO**: Preserve Korean language content exactly as written
- **DO**: Use Korean terminology when discussing course concepts
- **DO NOT**: Translate Korean content to English unless explicitly requested
- **DO**: Be aware of Korean date/time formats when parsing data

#### 2. File Modifications
- **DO**: Maintain UTF-8 encoding with BOM for CSV files
- **DO**: Preserve the hash suffixes in filenames when renaming
- **DO**: Keep the `[N회차]` prefix format for session files
- **DO NOT**: Change the directory structure without explicit request
- **DO NOT**: Remove or modify the `<aside>` blocks in markdown files

#### 3. Content Updates
- **DO**: Follow the established markdown structure for new content
- **DO**: Include learning objectives (학습 목표) in `<aside>` blocks
- **DO**: Maintain session metadata (단계, 중요도, 난이도, Min, No)
- **DO**: Use appropriate emoji indicators (🎯, 📋, 💡, etc.) consistently
- **DO NOT**: Add content that contradicts the Kraljic Matrix framework

#### 4. Data Operations
- **DO**: Validate supplier grades match performance scores (A: 90+, B: 80-89, C: 70-79, D: <70)
- **DO**: Ensure material category assignments align with Kraljic Matrix quadrants
- **DO**: Preserve all columns when updating CSV files
- **DO NOT**: Change date formats in CSV files
- **DO NOT**: Remove the BOM from CSV files

#### 5. Code/Script Development
If creating analysis scripts or tools:
- **DO**: Support Korean text (UTF-8 encoding)
- **DO**: Handle CSV files with BOM properly
- **DO**: Parse Korean date formats correctly
- **DO**: Provide bilingual comments (Korean + English) for clarity
- **DO NOT**: Assume ASCII-only input

---

## Common Tasks & Best Practices

### Adding New Course Content
1. Follow the `[N회차]` naming convention
2. Include all metadata fields at the top
3. Structure content with learning objectives
4. Add appropriate emoji indicators
5. Link related sessions using internal links

### Updating Supplier Data
1. Maintain CSV format with all columns
2. Validate grade assignments (A/B/C/D)
3. Ensure material category is one of: 전략/레버리지/병목/일상
4. Update `최종수정일` (last modified date) field
5. Keep both `{name}.csv` and `{name}_all.csv` in sync

### Analyzing Course Structure
- Reference the curriculum CSV for session ordering
- Use the Kraljic Matrix quadrants as the primary framework
- Consider the progression: Foundation → Deep Dives → Workshops
- Session 1-3: Overview concepts
- Session 4-7: Material-specific strategies
- Session 8-9: Practical application

### Working with Supplier Scorecards
- Grade A suppliers (90-100): Strategic partnerships
- Grade B suppliers (80-89): Good performance, room for improvement
- Grade C suppliers (70-79): Improvement plans needed
- Grade D suppliers (<70): Consider replacement
- Material category affects supplier strategy expectations

---

## Development Workflows

### Content Review Workflow
1. Read session file to understand topic and objectives
2. Verify alignment with Kraljic Matrix framework
3. Check internal links between related sessions
4. Validate metadata completeness
5. Ensure learning objectives match content depth

### Data Analysis Workflow
1. Load CSV with UTF-8 BOM encoding
2. Parse Korean headers correctly
3. Validate data types (percentages, scores, dates)
4. Cross-reference supplier grades with material categories
5. Generate insights aligned with course concepts

### Repository Maintenance
1. Keep extracted content in `전략적 재고운영 및 자재계획수립/` directory
2. Maintain archive files (`.zip`) for backup
3. Update README.md if major changes occur
4. Document any structural changes in commit messages
5. Preserve the git history for course evolution tracking

---

## Key Concepts Reference

### Kraljic Matrix Quadrants

**전략자재 (Strategic Materials)**
- High supply risk, high purchase impact
- Characteristics: Critical, few suppliers, complex
- Strategy: Long-term partnerships, collaborative planning
- Planning: Hybrid planning methods
- Examples in dataset: 미래금속, (주)한국정밀

**레버리지자재 (Leverage Materials)**
- Low supply risk, high purchase impact
- Characteristics: Many suppliers, standardized, high volume
- Strategy: Competitive bidding, volume leverage
- Planning: MRP (Material Requirements Planning)
- Examples in dataset: 글로벌스틸, 대한전자부품, 신한부품

**병목자재 (Bottleneck Materials)**
- High supply risk, low purchase impact
- Characteristics: Limited suppliers, specialized
- Strategy: Ensure supply continuity, buffer stock
- Planning: ROP (Re-Order Point)
- Examples in dataset: 중앙산업, 삼성화학, 태평양소재

**일상자재 (Routine Materials)**
- Low supply risk, low purchase impact
- Characteristics: Commodity items, many suppliers
- Strategy: Process efficiency, automation
- Planning: Automated ordering systems
- Examples in dataset: 동양플라스틱, 아시아MRO

### Inventory Planning Methods

**ROP (Re-Order Point)**
- For bottleneck materials
- Based on lead time and demand rate
- Safety stock for supply uncertainty

**MRP (Material Requirements Planning)**
- For leverage materials
- Demand-driven from production schedule
- Minimize holding costs through precise timing

**Hybrid Planning**
- For strategic materials
- Combines forecast-based and demand-based
- Balances relationship commitments and flexibility

**Automated Systems**
- For routine materials
- Minimize human intervention
- Focus on efficiency and cost reduction

---

## Troubleshooting

### Common Issues

**Issue**: CSV files display incorrectly
- **Cause**: BOM not recognized or wrong encoding
- **Solution**: Open with UTF-8 BOM encoding explicitly

**Issue**: Markdown formatting broken
- **Cause**: Notion-specific syntax (`<aside>` blocks)
- **Solution**: Use markdown processors that support HTML blocks

**Issue**: Internal links not working
- **Cause**: URL-encoded Korean characters in links
- **Solution**: URL-decode links when processing programmatically

**Issue**: Supplier grade doesn't match score
- **Cause**: Data entry error or outdated calculation
- **Solution**: Recalculate total score, verify grade assignment

---

## Version Control Guidelines

### Commit Messages
- Use Korean for content changes: "2회차 내용 업데이트"
- Use English for structural changes: "Add new session template"
- Reference session numbers: "[4회차] Add ROP calculation examples"

### Branch Strategy
- Current branch: `claude/claude-md-mi3s2y2jmbmk6esm-01EnHEQoFiPzet32PAdnoyKB`
- Always develop on designated feature branches
- Never push to main/master without explicit permission

### What to Commit
- ✅ Course content updates (markdown files)
- ✅ Data updates (CSV files)
- ✅ New analysis scripts or tools
- ✅ Documentation improvements
- ❌ Temporary files or build artifacts
- ❌ Extracted archives (keep only source zips)

---

## Future Extensions

### Potential Enhancements
1. **Interactive Workshops**: Add code examples for Kraljic classification
2. **Data Analysis Tools**: Python/R scripts for supplier scorecard analysis
3. **Visualization**: Generate Kraljic Matrix plots from supplier data
4. **Translation**: English version for international audiences
5. **Case Studies**: Add real-world company examples
6. **Assessment Tools**: Quizzes and exercises for each session
7. **API Integration**: Connect to actual ERP/SCM systems

### Maintaining Course Relevance
- Update supplier examples with current market conditions
- Refresh case studies annually
- Incorporate new supply chain trends (e.g., sustainability, digitalization)
- Add content on emerging topics (AI in SCM, blockchain, circular economy)

---

## Resources & References

### Course Topics Covered
- Kraljic Matrix methodology (Session 1)
- Supplier relationship management (Session 2)
- ABC-XYZ inventory classification (Session 3)
- ROP planning for bottleneck materials (Session 4)
- MRP for leverage materials (Session 5)
- Hybrid planning for strategic materials (Session 6)
- Automation for routine materials (Session 7)
- Practical workshops (Sessions 8-9)

### Related Frameworks
- JIT (Just-In-Time) vs JIC (Just-In-Case)
- ABC Analysis
- XYZ Analysis
- Supplier Scorecard methodology
- Material Requirements Planning (MRP)
- Re-Order Point (ROP) systems

---

## Contact & Contribution

### Repository Information
- **Repository**: TurtleLEE-teacher/Kraljic_Course
- **Primary Language**: Korean
- **Content Type**: Educational course materials
- **Format**: Markdown + CSV data

### For AI Assistants
- Treat Korean text with care and precision
- Respect the educational nature of the content
- Maintain consistency with Kraljic Matrix framework
- Preserve the structured learning progression
- When in doubt, ask for clarification rather than assuming

---

**Last Updated**: 2025-11-19
**CLAUDE.md Version**: 2.0
**Repository State**: Course content updated (Nov 19), Design guidelines enhanced

## Update History

### Version 2.0 (2025-11-19)
- **Content Update**: All session files updated with Notion_251119 export
  - Session 5 (레버리지자재 & MRP): +483 lines - Major expansion with industry examples
  - Session 7 (일상자재 효율화): +232 lines - Detailed automation strategies
  - Session 6 (전략자재 & 하이브리드): +184 lines - Enhanced hybrid planning
  - Session 1 (Foundation): +99 lines - Strengthened JIT/JIC paradigm explanation
  - Session 4 (병목자재 & ROP): +42 lines - Improved ROP methodology
  - **Total**: +1,051 lines of enhanced content
- **Design Guidelines**: Enhanced with professional training insights
  - Font size analysis: 10pt confirmed as THE professional standard (65.2%)
  - Door chart pattern documented (75+ shapes for matrices)
  - Three storyline approaches: Structural, Dynamics, Market Change
  - Shape count targets: 20-50+ per slide (70-80% in groups)

### Version 1.0 (2025-11-17)
- Initial structure documentation
- Repository overview and course guide
- PPTX generation guidelines
- File conventions and data schema

