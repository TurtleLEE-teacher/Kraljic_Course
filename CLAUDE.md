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
- **Font sizes** (actual usage from S4HANA reference):
  - **48pt**: Cover slide main title (Bold)
  - **20pt**: Content slide titles (Bold)
  - **16pt**: Governing messages (Bold)
  - **14pt**: Section headers, large bullet points
  - **12-13pt**: Regular bullet points
  - **9-11pt**: Body text, detailed descriptions (**most common**)
  - **8pt**: Small annotations, footnotes
  - **6-7pt**: Tiny notes (rare)
- **Font weights**: Bold for titles/headers, Regular for body
- **Key insight**: S4HANA uses **small fonts (8-11pt) extensively** to fit more content

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
4. **Using wrong dimensions**: Must be 10.83" × 7.5"
5. **Ignoring the reference file**: It's the gold standard
6. **Direct python-pptx coding without skill workflow**: Try skill approach first
7. **Using gradients on cover slide**: Use solid colors (gradient causes rendering issues)
8. **Using too large fonts**: Don't use 16-18pt for body text - use 8-11pt like S4HANA!
9. **Too much whitespace**: Slides must be 85%+ filled - add more content, diagrams, shapes
10. **Missing visual elements**: No flowcharts, arrows, or shapes - S4HANA uses 10-50+ shapes per slide
11. **Misunderstanding Part numbers**: "Part 1" = Session 1 only, NOT Sessions 1-3

### ✅ Checklist Before Generating PPTX

- [ ] Read complete SKILL.md (no offset/limit)
- [ ] Read complete html2pptx.md
- [ ] Read complete css.md
- [ ] Analyzed S4HANA reference PPTX file with detailed script
- [ ] Understood monochrome color system (3-Color Rule)
- [ ] Understood font size ranges (8-11pt for body, NOT 16-18pt)
- [ ] Planned content density to achieve 85%+ filled area
- [ ] Designed flowcharts, diagrams, shapes (10-50+ per slide)
- [ ] Planned governing messages for all content slides (16pt Bold)
- [ ] Prepared JSON data structure
- [ ] Verified Handlebars templates exist or created them
- [ ] Confirmed slide dimensions: 10.83" × 7.5"
- [ ] Confirmed Part/Session mapping (Part N = Session N, not Sessions N-M)
- [ ] Tested with small sample before full generation

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

**Last Updated**: 2025-11-17
**CLAUDE.md Version**: 1.0
**Repository State**: Initial structure documented
