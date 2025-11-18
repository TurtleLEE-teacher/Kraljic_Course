---
name: pptx-mslee
description: "Professional business reporting PPTX generator using HTML→PPTX workflow. JSON data → HTML templates → high-quality PowerPoint with S4HANA monochrome design, governing messages, and precise layout control. Optimized for business presentations and reports."
license: Proprietary. LICENSE.txt has complete terms
---

# pptx-mslee - 비즈니스 리포팅 고품질 PPTX 생성 도구

**Professional Business Reporting PowerPoint Generator**

## ⚠️ CRITICAL: Skill Priority Rules

**This skill's guidelines ALWAYS take precedence over external samples or references.**

When creating HTML/CSS for PPTX conversion:
1. **MANDATORY**: Read this SKILL.md, html2pptx.md, and css.md COMPLETELY before starting
2. **Priority order**:
   - ① This skill's guidelines (HIGHEST PRIORITY)
   - ② html2pptx.md and css.md specifications
   - ③ External HTML samples (reference only, DO NOT violate skill rules)
3. **If conflict**: Follow this skill's rules, NOT external samples

**Failure to follow these rules will result in conversion failures or broken presentations.**

---

## 🎯 Overview

This tool creates **high-quality PowerPoint presentations** for business reporting and presentations using an **HTML→PPTX workflow** that ensures:

- ✅ **Design Consistency**: Monochrome color system and professional typography
- ✅ **No Text Overflow**: Precise layout control through HTML templates
- ✅ **Full Automation**: JSON data → HTML → PPTX in one command
- ✅ **Business Standards**: Governing messages, structured layouts, high information density

**Primary Use Case**: Creating business reports, project presentations, and executive briefings with 10-50 slides maintaining S4HANA professional design standards.

---

## 🌟 PRIMARY WORKFLOW: Business Reports (JSON → HTML → PPTX)

**This is the REQUIRED workflow for creating business presentations** to ensure high quality.

### Why This Workflow?

| Feature | Direct PptxGenJS | **HTML→PPTX (This Tool)** |
|---------|------------------|---------------------------|
| Layout Precision | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ |
| Text Overflow Prevention | ❌ Manual check | ✅ Automatic |
| Design Consistency | ⭐⭐⭐ | ⭐⭐⭐⭐⭐ (CSS enforced) |
| 3-Color Rule Compliance | ⭐⭐ | ⭐⭐⭐⭐⭐ (Enforced) |
| Speed | ⭐⭐⭐⭐⭐ (1s) | ⭐⭐⭐ (3-5s) |

**Conclusion**: We sacrifice 2-4 seconds for **perfect quality** ✅

### Quick Start (5 Minutes)

```bash
# 1. Install dependencies (one-time)
cd ~/.claude/skills/pptx-mslee
npm install

# 2. Verify html2pptx is installed
npm list -g @ant/html2pptx || npm install -g html2pptx.tgz

# 3. Create JSON data
cat > data/my-report.json << 'EOF'
{
  "title": "My Business Report",
  "date": "2025-01",
  "slides": [
    {
      "id": 1,
      "layout": "cover",
      "data": {
        "title": "Report Title",
        "subtitle": "Subtitle",
        "author": "Author Name",
        "date": "2025"
      }
    }
  ]
}
EOF

# 4. Generate PPTX (auto: JSON → HTML → PPTX)
node scripts/generate-report.js data/my-report.json

# 5. Open result
open output/my-report.pptx
```

**Result**: High-quality PPTX with perfect design consistency!

### JSON Data Structure

**Report-level metadata:**
```json
{
  "title": "Report Title",
  "author": "Author Name",
  "date": "2025-01",
  "totalSlides": 5,
  "description": "Report description",
  "slides": [...]            // Array of slide objects
}
```

**Slide object structure:**
```json
{
  "id": 1,
  "layout": "cover|content-2col|list-bullets|...",
  "data": {
    // Layout-specific data (see below)
  }
}
```

**Available Layouts:**

#### 1. Cover (Title Slide)
```json
{
  "layout": "cover",
  "data": {
    "title": "Main Title",
    "subtitle": "Subtitle",
    "course": "Course Name",
    "date": "2025",
    "instructor": "Instructor Name"
  }
}
```

**Features**:
- Clean monochrome design
- Center-aligned
- 48pt bold title, 28pt subtitle

#### 2. Content-2Col (Two-Column Content)
```json
{
  "layout": "content-2col",
  "data": {
    "title": "Slide Title",
    "governingMessage": "One-sentence summary of the entire slide",
    "leftTitle": "Left Column Title",
    "leftContent": "Left content\n(use \\n for line breaks)",
    "rightTitle": "Right Column Title",
    "rightContent": "Right content",
    "footer": "Report Name",
    "slideNumber": 3
  }
}
```

**Features**:
- 50:50 column split
- Perfect for comparisons (Before/After, ROP vs MRP)
- **Governing message required** (14pt italic summary)

#### 3. List-Bullets (Bullet Points)
```json
{
  "layout": "list-bullets",
  "data": {
    "title": "Slide Title",
    "governingMessage": "One-sentence summary",
    "introduction": "Optional intro text",
    "items": [
      "First bullet point",
      "Second bullet point",
      "Third bullet point (max 6 recommended)"
    ],
    "footer": "Report Name",
    "slideNumber": 4
  }
}
```

**Features**:
- Maximum 6 items (6x6 rule)
- Clean bullet design
- Optional introduction paragraph
- **Governing message required**

### Design Principles (S4HANA Business Reporting Style)

#### Monochrome Color System
**Strictly enforced** through CSS variables for professional, clean appearance:
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

**Color usage principles**:
- Black/gray scale only for text and backgrounds
- No bright colors except for data visualization (charts)
- Maximum contrast for readability
- Consistent throughout entire presentation

#### Typography (Arial Only)
```css
--font-title: 20pt Bold;         /* Slide titles */
--font-governing: 14pt Italic;   /* Governing messages */
--font-heading: 14pt Bold;       /* Section headings */
--font-body: 12pt Regular;       /* Body text */
--font-caption: 11pt Regular;    /* Captions, notes */
```

**Font rules**:
- **Only Arial** - web-safe, universally available
- No custom fonts or font downloads
- Use weight and size for hierarchy

#### Governing Message Pattern (REQUIRED)
Every content slide must include a governing message - a one-sentence summary that captures the entire slide's key point:

```html
<div class="title-section fit">
  <h1>Slide Title</h1>
  <p class="governing-message">One-sentence summary of the entire slide content.</p>
</div>
```

**Governing message guidelines**:
- Maximum 1-2 sentences
- 14pt italic font
- Positioned directly under the title
- Summarizes the "so what" of the slide
- Reader should understand the slide's purpose from this alone

#### Slide Dimensions
```css
body {
  width: 780px;   /* 13:9 aspect ratio */
  height: 540px;
  padding: 25px 40px 45px 40px;  /* top, sides, bottom */
}
```

**Layout requirements**:
- 13:9 aspect ratio (business reporting standard)
- Minimum 0.5" (45px) bottom margin
- Content density: 80%+ target
- No wasted whitespace

#### 2-Column Comparison Layout
For vs/comparison slides (e.g., "ROP vs MRP", "Before vs After"):

```html
<div class="fill-height row gap items-fill-width">
  <div>
    <h2>Left Side</h2>
    <p>Content...</p>
  </div>
  <div>
    <h2>Right Side</h2>
    <p>Content...</p>
  </div>
</div>
```

**Features**:
- Equal width columns (50:50 split)
- Vertical alignment maintained
- Clear visual separation
- Side-by-side comparison emphasis

### Detailed Workflow Steps


#### Step 0: Prerequisites - Read Documentation

**MANDATORY - READ SKILL FILES FIRST**:
- Read this SKILL.md "HTML/CSS Creation Guidelines" section (below)
- Read [`html2pptx.md`](html2pptx.md) completely from start to finish
- Read [`css.md`](css.md) completely from start to finish
- **NEVER set any range limits when reading these files**
- These files contain critical rules that MUST be followed
- External HTML samples are for reference only - skill rules take precedence

#### Step 1: Create JSON Data File

```bash
# Example: Business Project Report
cat > data/project-report.json << 'EOF'
{
  "title": "Project Status Report",
  "author": "Project Manager",
  "date": "2025-01",
  "slides": [
    {
      "id": 1,
      "layout": "cover",
      "data": {
        "title": "Q1 Project Status Report",
        "subtitle": "Project Performance Summary",
        "author": "Project Management Office",
        "date": "2025-01"
      }
    },
    {
      "id": 2,
      "layout": "content-2col",
      "data": {
        "title": "Traditional vs Modern Approach",
        "governingMessage": "Modern approach delivers 40% faster results through automation and data-driven decisions.",
        "leftTitle": "Traditional Approach",
        "leftContent": "• Manual data collection\n• Siloed departments\n• Reactive management",
        "rightTitle": "Modern Approach",
        "rightContent": "• Automated dashboards\n• Cross-functional teams\n• Predictive analytics",
        "footer": "Project Status Report",
        "slideNumber": 2
      }
    }
  ]
}
EOF
```

#### Step 2: Generate PPTX (Automatic HTML → PPTX)

```bash
node scripts/generate-report.js data/project-report.json
```

**What happens internally** (fully automatic):
1. **JSON parsing**: Load report data
2. **HTML generation**: For each slide:
   - Load Handlebars template (`layouts/{layoutName}.hbs`)
   - Inject CSS variables (`styles/theme-s4hana.css`)
   - Apply monochrome color scheme
   - Render HTML to `temp/slide-{id}.html`
3. **html2pptx conversion**: Convert each HTML file to PPTX slide
4. **PPTX saving**: Combine all slides and save
5. **Cleanup**: Remove temporary HTML files (unless `--keep-html` flag)

**Options:**
```bash
# Debug mode (show detailed logs)
node scripts/generate-report.js data/my-report.json --debug

# Keep HTML files for inspection
node scripts/generate-report.js data/my-report.json --keep-html

# Generate quality report
node scripts/generate-report.js data/my-report.json --report

# Custom output filename
node scripts/generate-report.js data/my-report.json --output custom-name.pptx
```

#### Step 3: Verify Quality (Automatic)

The tool automatically:
- ✅ Checks monochrome design compliance
- ✅ Validates text overflow (via HTML layout)
- ✅ Ensures font consistency (Arial only)
- ✅ Verifies governing message presence
- ✅ Validates minimum bottom margin (0.5")

**Manual verification** (optional):
```bash
# Generate thumbnail grid
python scripts/thumbnail.py output/project-report.pptx

# Review thumbnail for:
# - Text cutoff (should be none)
# - Layout alignment (should be perfect)
# - Monochrome color consistency
# - Governing messages present on all content slides
```

#### Step 4: Batch Processing (Multiple Reports)

```bash
# Generate multiple reports at once
node scripts/generate-report.js data/*.json --batch

# Result:
# output/project-report.pptx (15 slides)
# output/status-update.pptx (8 slides)
# output/executive-briefing.pptx (12 slides)
# ... etc.
```

### File Structure

```
pptx-mslee/
├── data/                          # JSON data files
│   ├── project-report.json
│   └── status-update.json
│
├── templates/business-report/     # HTML templates (Handlebars)
│   ├── layouts/
│   │   ├── cover.hbs              # Cover slide template
│   │   ├── content-2col.hbs       # Two-column template
│   │   └── list-bullets.hbs       # Bullet list template
│   ├── styles/
│   │   ├── variables.css          # Design variables
│   │   └── theme-s4hana.css       # S4HANA monochrome theme
│   └── partials/
│       ├── header.hbs             # Reusable header
│       └── footer.hbs             # Reusable footer
│
├── scripts/
│   ├── generate-report.js         # Main generation script
│   ├── report-pptx-builder.js     # Builder class (HTML-based)
│   └── html-generator.js          # Handlebars → HTML
│
├── temp/                          # Temporary HTML files
│   ├── s4hana-style/              # S4HANA style samples
│   └── slide-*.html               # (auto-deleted after conversion)
│
└── output/                        # Generated PPTX files
    └── *.pptx
```

### Troubleshooting

**Error: "Cannot find module '@ant/html2pptx'"**
```bash
# Solution: Install html2pptx
npm install -g ~/.claude/skills/pptx-mslee/html2pptx.tgz
npm install -g playwright
```

**Error: "Template not found: {layoutName}"**
```bash
# Solution: Check available layouts
ls templates/education-course/layouts/

# Ensure JSON references correct layout name
# "layout": "content-2col" → layouts/content-2col.hbs must exist
```

**Warning: Text overflow detected**
```bash
# Solution: Reduce text length or adjust layout
# Edit JSON data to shorten content
# Or use different layout with more space
```

**Slow generation (>10 seconds per slide)**
```bash
# Normal: 3-5 seconds per slide with html2pptx
# If slower, check:
node --version  # Should be >= 18.0.0
npm list -g playwright  # Should be installed
```

---

## HTML/CSS Creation Guidelines for PPTX Conversion

### Supported HTML/CSS Features

#### ✅ Fully Supported (Reliable conversion)

**Layout:**
- `.row`, `.col` classes from css.md (PREFERRED over flexbox)
- `.fill-width`, `.fill-height`, `.fit` for flex sizing
- Basic positioning with padding/margin
- Two-column layouts with defined widths

**Typography:**
- `<p>`, `<h1>` through `<h6>` for ALL text
- Web-safe fonts ONLY: Arial, Helvetica, Times New Roman, Georgia, Courier New, Verdana, Tahoma, Trebuchet MS, Impact
- Font sizes, weights, colors
- Text alignment (left, center, right)

**Colors & Backgrounds:**
- Solid colors via CSS variables or hex codes
- CSS `linear-gradient()` and `radial-gradient()` (vertical/horizontal only)
- `background-image: url(...)` for images

**Borders & Shapes:**
- Solid borders with `border`, `border-radius`
- Basic shapes via CSS

**Lists:**
- `<ul>`, `<ol>` with automatic bullet/number rendering

#### ⚠️ Limited Support (Works with constraints)

**Gradients:**
- Only vertical/horizontal linear gradients
- Radial gradients (simple center-based only)
- Complex multi-stop gradients may not render accurately

**SVG:**
- Inline SVG converted to images
- Complex SVG features may be simplified

**Shadows:**
- Basic `box-shadow` (may be simplified in conversion)

#### ❌ NOT Supported (Will be ignored or cause errors)

**Dynamic/Interactive:**
- `:hover`, `:focus`, `:active` pseudo-classes
- CSS animations (`@keyframes`, `animation`, `transition`)
- JavaScript interactions

**Responsive:**
- Media queries (`@media`)
- Responsive utility classes (`md:`, `lg:`, `sm:`)
- Container queries

**Advanced CSS:**
- CSS transforms (`rotate`, `scale`, `translate`, `skew`)
- `position: fixed` or `position: sticky`
- `::before`, `::after` pseudo-elements
- Complex CSS selectors (`:nth-child()`, `:has()`)
- CSS Grid (use `.row`/`.col` instead)
- Direct `display: flex` (use `.row`/`.col` classes)

**Typography:**
- Custom web fonts (Google Fonts, etc.)
- Variable fonts
- `text-shadow` (limited support)

### Tailwind CSS Usage Guidelines

#### When to Use Tailwind

Tailwind can be used, but requires **preprocessing** to convert to the skill's native CSS framework. Consider:

- **Preferred**: Use native CSS framework (`.row`, `.col`, etc.) for maximum reliability
- **Acceptable**: Simple Tailwind utilities that map cleanly
- **Avoid**: Complex Tailwind features (responsive, hover, group)

#### ✅ Supported Tailwind Classes (Can be preprocessed)

**Layout:**
- `flex`, `flex-row`, `flex-col` → `.row`, `.col`
- `flex-1` → `.fill-width` or `.fill-height`
- `w-*`, `h-*` → Width/height utilities
- `gap-*` → `.gap`, `.gap-sm`, `.gap-lg`

**Spacing:**
- `p-*`, `px-*`, `py-*` → `.p-*`, `.px-*`, `.py-*`
- `m-*`, `mx-*`, `my-*` → `.m-*`, `.mx-*`, `.my-*`

**Typography:**
- `text-xs` through `text-8xl` → `.text-xs` through `.text-8xl`
- `font-bold`, `font-semibold` → `font-weight` properties
- `text-center`, `text-left`, `text-right` → `.text-center`, etc.

**Colors:**
- `bg-*`, `text-*` → CSS color variables
- `border-*` → Border color properties

**Visual:**
- `rounded-*` → `.rounded`, `.pill`
- `opacity-*` → `.opacity-*`

#### ❌ Unsupported Tailwind Classes (Will be ignored)

- Responsive: `md:`, `lg:`, `xl:`, `2xl:`, `sm:`
- Interactive: `hover:`, `focus:`, `active:`, `group-hover:`
- Dark mode: `dark:`
- Animations: `animate-*`, `transition-*`
- Transforms: `rotate-*`, `scale-*`, `translate-*`, `skew-*`
- Advanced positioning: `absolute`, `fixed`, `sticky`

#### Tailwind Conversion Process

If you need to use Tailwind HTML:

1. Write HTML with supported Tailwind classes only
2. Run preprocessing to convert to native framework classes
3. Verify converted HTML uses only supported features
4. Proceed with normal html2pptx workflow

**Example Conversion:**
```html
<!-- Input: Tailwind -->
<div class="flex flex-col gap-4 p-8 bg-blue-500">
  <h1 class="text-4xl font-bold text-white">Title</h1>
  <p class="text-lg text-white opacity-80">Description</p>
</div>

<!-- Output: Native Framework -->
<div class="col gap p-8 bg-primary">
  <h1 class="text-4xl" style="font-weight: 600; color: var(--color-primary-foreground);">Title</h1>
  <p class="text-lg opacity-80" style="color: var(--color-primary-foreground);">Description</p>
</div>
```

### Recommended Patterns

#### ✅ PREFERRED: Native CSS Framework
```html
<body class="col">
  <header class="fit">
    <h2>Slide Title</h2>
  </header>
  <main class="fill-height row gap-lg">
    <section class="fill-width">
      <p>Content here</p>
    </section>
    <aside class="bg-muted p-4 rounded" style="min-width: 200px;">
      <p>Sidebar</p>
    </aside>
  </main>
</body>
```

#### ✅ ACCEPTABLE: Simple Inline Styles
```html
<div style="display: flex; flex-direction: column; padding: 20px; background: #f5f5f5;">
  <h2 style="color: #333; font-size: 24px;">Title</h2>
  <p style="color: #666; font-size: 16px;">Description</p>
</div>
```

(But `.col .p-8 .bg-muted` would be better)

#### ❌ FORBIDDEN Patterns
```html
<!-- ❌ Text directly in div -->
<div>This text won't appear in PowerPoint</div>

<!-- ❌ Direct flexbox without row/col classes -->
<div style="display: flex;">Content</div>

<!-- ❌ Custom fonts -->
<p style="font-family: 'Roboto', sans-serif;">Text</p>

<!-- ❌ Hover states -->
<button class="hover:bg-blue-500">Button</button>

<!-- ❌ Pseudo-elements -->
<div class="before:content-['→']">Text</div>

<!-- ❌ Responsive classes -->
<div class="md:flex lg:grid">Content</div>
```

### Critical Validation Checklist

Before converting HTML to PPTX, verify:

- [ ] All text is inside `<p>`, `<h1>`-`<h6>`, `<ul>`, or `<ol>` tags
- [ ] Only web-safe fonts are used
- [ ] Layout uses `.row`/`.col` classes (NOT direct `display: flex`)
- [ ] No hover/focus/active states
- [ ] No animations or transitions
- [ ] No responsive utilities
- [ ] No CSS transforms
- [ ] No custom fonts
- [ ] No `::before`/`::after` pseudo-elements
- [ ] Gradients are simple (vertical/horizontal only)

**If any checklist item fails, fix it before proceeding with conversion.**

---

## 📚 Advanced Workflows (For Specific Use Cases)

### A. Editing Existing PPTX Files (OOXML)

**Use when**: You need to modify an existing PowerPoint file's XML structure directly.

**Workflow**:
1. **Read OOXML guide**: `cat ooxml.md` (complete file, ~500 lines)
2. **Unpack**: `python ooxml/scripts/unpack.py input.pptx unpacked/`
3. **Edit XML**: Modify `ppt/slides/slide{N}.xml` files
4. **Validate**: `python ooxml/scripts/validate.py unpacked/ --original input.pptx`
5. **Pack**: `python ooxml/scripts/pack.py unpacked/ output.pptx`

**Key XML files**:
- `ppt/presentation.xml` - Slide references
- `ppt/slides/slide{N}.xml` - Slide contents
- `ppt/theme/theme1.xml` - Colors and fonts

### B. Reusing Existing Templates (Python Scripts)

**Use when**: You have an existing PowerPoint template and want to duplicate/rearrange slides.

**Workflow**:
1. **Extract template text**: `python -m markitdown template.pptx > template-content.md`
2. **Create thumbnail grid**: `python scripts/thumbnail.py template.pptx`
3. **Analyze inventory**: Create `template-inventory.md` with slide descriptions
4. **Rearrange slides**: `python scripts/rearrange.py template.pptx working.pptx 0,5,5,12,20`
5. **Extract text shapes**: `python scripts/inventory.py working.pptx text-inventory.json`
6. **Replace text**: Create `replacement-text.json` and run:
   ```bash
   python scripts/replace.py working.pptx replacement-text.json output.pptx
   ```

**Key scripts**:
- `inventory.py` - Extract all text shapes with positions
- `rearrange.py` - Duplicate/reorder/delete slides
- `replace.py` - Replace text content while preserving formatting

### C. Reading and Analyzing PPTX

**Text extraction** (markdown format):
```bash
python -m markitdown presentation.pptx
```

**Visual analysis** (thumbnail grids):
```bash
# Create 5-column grid (max 30 slides per grid)
python scripts/thumbnail.py presentation.pptx

# Custom: 4 columns, 20 slides per grid
python scripts/thumbnail.py presentation.pptx --cols 4

# Output: thumbnails.jpg (or thumbnails-1.jpg, thumbnails-2.jpg for large decks)
```

**Raw XML access** (for advanced analysis):
```bash
python ooxml/scripts/unpack.py presentation.pptx unpacked/

# Read specific files:
# - ppt/presentation.xml (structure)
# - ppt/slides/slide1.xml (first slide content)
# - ppt/notesSlides/notesSlide1.xml (speaker notes)
# - ppt/comments/modernComment_*.xml (comments)
```

---

## 🔧 Dependencies

**Required (install once):**

### Node.js Dependencies
```bash
cd ~/.claude/skills/pptx-mslee
npm install

# Packages:
# - pptxgenjs: PowerPoint generation library
# - handlebars: Template engine (for HTML generation)
# - @ant/html2pptx: HTML → PPTX conversion (install globally)
# - playwright: HTML rendering (required by html2pptx)
# - cheerio: HTML parsing
# - chalk: Terminal colors
```

### Global Dependencies
```bash
# html2pptx (critical for HTML→PPTX workflow)
npm install -g ~/.claude/skills/pptx-mslee/html2pptx.tgz

# Playwright (required by html2pptx)
npm install -g playwright
npx playwright install  # Download browser binaries
```

### Python Dependencies
```bash
# For OOXML editing and template reuse
pip install "markitdown[pptx]"  # Text extraction
pip install defusedxml          # Secure XML parsing
pip install python-pptx         # PPTX manipulation
pip install Pillow              # Thumbnail generation

# For PDF conversion (thumbnail script)
sudo apt-get install libreoffice poppler-utils  # Linux
# or brew install libreoffice poppler  # macOS
```

**Verification**:
```bash
# Check Node.js version (>= 18.0.0)
node --version

# Check html2pptx installation
npm list -g @ant/html2pptx

# Check Python packages
pip list | grep -E "markitdown|defusedxml|python-pptx|Pillow"
```

---

## 📖 Additional Documentation

- **Quick Start Guide**: `docs/QUICK-START.md` - 5-minute tutorial
- **Template Guide**: `docs/TEMPLATE-GUIDE.md` - How to create custom layouts
- **HTML→PPTX Details**: `html2pptx.md` - Complete html2pptx syntax and examples
- **OOXML Editing**: `ooxml.md` - XML structure and editing guide
- **Design Guidelines**: `references/design-guidelines.md` - Education materials standards

---

## 🎓 Best Practices

### 1. JSON Data Organization
```
data/
├── project-report.json        # Project status (15 slides)
├── executive-briefing.json    # Executive summary (8 slides)
└── quarterly-review.json      # Quarterly review (12 slides)
```

### 2. Content Guidelines
- **Titles**: Max 3 seconds to read
- **Governing Messages**: REQUIRED on all content slides (1-2 sentences)
- **Bullets**: Max 6 items per slide (6x6 rule)
- **Text**: Use `\n` for line breaks in JSON strings
- **Colors**: Monochrome only (black/gray scale)
- **Fonts**: Arial only - no custom fonts

### 3. Tables in PPTX
**IMPORTANT**: HTML `<table>` tags are **NOT supported** by html2pptx. Use PptxGenJS instead:

```javascript
// After html2pptx conversion
const { slide, placeholders } = await html2pptx("slide.html", pptx);

// Add table using PptxGenJS
const tableData = [
  [{ text: "Header 1", options: { bold: true } }, { text: "Header 2", options: { bold: true } }],
  ["Data 1", "Data 2"],
  ["Data 3", "Data 4"]
];

slide.addTable(tableData, {
  x: 0.5, y: 1.5, w: 9, h: 3,
  border: { pt: 1, color: "CCCCCC" },
  fontSize: 12,
  color: "333333"
});
```

See `html2pptx.md` for complete table guide.

### 4. Version Control
```bash
# Track JSON data files
git add data/*.json

# Track template changes
git add templates/

# Ignore output files
echo "output/" >> .gitignore
echo "temp/" >> .gitignore
```

### 5. Quality Checks
```bash
# After generation, always verify:
# 1. Thumbnail visual check
python scripts/thumbnail.py output/my-report.pptx

# 2. Quality report
node scripts/generate-report.js data/my-report.json --report
cat output/my-report-report.json

# 3. Manual review in PowerPoint
open output/my-report.pptx
```

---

**Version**: 3.0 (S4HANA Business Reporting Style)
**Last Updated**: 2025-11-18
**Primary Maintainer**: pptx-mslee Team
**Style**: S4HANA Monochrome Design
