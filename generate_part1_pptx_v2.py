#!/usr/bin/env python3
"""
Part 1 PPTX Generator V2 - High Quality Implementation
Session 1: Kraljic Matrix와 자재계획 방법론

Following S4HANA Professional Standards:
- Dimensions: 10.83" × 7.50" (1.44:1)
- Shape counts: 40-120 per complex slide
- Font distribution: 9-10pt = 48% (primary body text)
- Monochrome color system (black/white/gray)
- Governing messages: 16pt Bold
- Door chart: 75-100 shapes for Kraljic Matrix
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

# S4HANA Color Palette (Monochrome)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_DARK_GRAY = RGBColor(51, 51, 51)
COLOR_MED_GRAY = RGBColor(102, 102, 102)
COLOR_LIGHT_GRAY = RGBColor(204, 204, 204)
COLOR_VERY_LIGHT_GRAY = RGBColor(230, 230, 230)
COLOR_WHITE = RGBColor(255, 255, 255)

# Kraljic Matrix colors (use ONLY in Matrix slide)
COLOR_STRATEGIC = RGBColor(142, 68, 173)
COLOR_BOTTLENECK = RGBColor(230, 126, 34)
COLOR_LEVERAGE = RGBColor(39, 174, 60)
COLOR_ROUTINE = RGBColor(149, 165, 166)

def create_presentation():
    """Create presentation with S4HANA dimensions"""
    prs = Presentation()
    prs.slide_width = Inches(10.83)
    prs.slide_height = Inches(7.50)
    return prs

# ============================================================================
# HELPER FUNCTIONS - Shape Generation
# ============================================================================

def add_rectangle(slide, x, y, w, h, fill_color, border_color=None, border_width=1):
    """Add a rectangle shape

    Args:
        slide: Slide object
        x, y, w, h: Position and size in inches
        fill_color: RGBColor for fill
        border_color: RGBColor for border (None = no border)
        border_width: Border width in pt

    Returns:
        Shape object
    """
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color

    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()

    return shape

def add_text_box(slide, x, y, w, h, text, font_size=10, bold=False,
                 color=COLOR_BLACK, align=PP_ALIGN.LEFT, font_name='맑은 고딕'):
    """Add a text box with specified formatting

    Args:
        slide: Slide object
        x, y, w, h: Position and size in inches
        text: Text content
        font_size: Font size in pt
        bold: Bold text
        color: RGBColor for text
        align: Text alignment
        font_name: Font name

    Returns:
        Shape object
    """
    textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    p = text_frame.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align

    return textbox

def add_arrow(slide, x1, y1, x2, y2, color=COLOR_DARK_GRAY, width=2):
    """Add an arrow connector

    Args:
        slide: Slide object
        x1, y1: Start position in inches
        x2, y2: End position in inches
        color: RGBColor for arrow
        width: Line width in pt

    Returns:
        Connector object
    """
    from pptx.enum.shapes import MSO_CONNECTOR

    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1), Inches(y1),
        Inches(x2), Inches(y2)
    )
    connector.line.color.rgb = color
    connector.line.width = Pt(width)

    # Add arrowhead at end
    connector.line.end_arrow_type = 2  # Arrow

    return connector

def add_slide_title(slide, title, slide_num=None):
    """Add standard slide title

    Returns:
        Textbox object
    """
    # Title
    textbox = add_text_box(
        slide, 0.30, 0.30, 10.23, 0.60,
        title, font_size=20, bold=True, color=COLOR_BLACK
    )

    # Slide number (if provided)
    if slide_num:
        add_text_box(
            slide, 10.00, 7.00, 0.50, 0.30,
            str(slide_num), font_size=8, color=COLOR_MED_GRAY,
            align=PP_ALIGN.RIGHT, font_name='Arial'
        )

    return textbox

def add_governing_message(slide, message):
    """Add governing message under title (16pt Bold)

    Returns:
        Textbox object
    """
    return add_text_box(
        slide, 0.30, 1.01, 10.32, 0.63,
        message, font_size=16, bold=True, color=COLOR_MED_GRAY
    )

# ============================================================================
# SLIDE 1: COVER
# ============================================================================

def create_slide_1_cover(prs):
    """Slide 1: Cover - Simple design"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Main title (48pt)
    add_text_box(
        slide, 1.00, 2.50, 8.83, 1.00,
        "전략적 재고운영 Foundation",
        font_size=48, bold=True, color=COLOR_BLACK,
        align=PP_ALIGN.CENTER
    )

    # Subtitle (28pt)
    add_text_box(
        slide, 1.00, 3.70, 8.83, 0.80,
        "Kraljic Matrix와 자재계획 방법론",
        font_size=28, bold=False, color=COLOR_DARK_GRAY,
        align=PP_ALIGN.CENTER
    )

    # Course info (14pt)
    add_text_box(
        slide, 1.00, 5.00, 8.83, 0.50,
        "Session 1 | 전략적 재고운영 및 자재계획수립 과정",
        font_size=14, color=COLOR_MED_GRAY,
        align=PP_ALIGN.CENTER
    )

    print("✓ Slide 1: Cover")
    return slide

# ============================================================================
# SLIDE 2: TOC (15-20 shapes)
# ============================================================================

def create_slide_2_toc(prs):
    """Slide 2: Table of Contents with 7 chapter boxes"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "목차", slide_num=2)
    add_governing_message(
        slide,
        "본 과정은 Kraljic Matrix 기반으로 자재군별 차별화 전략과 계획 방법론을 체계적으로 학습합니다."
    )

    # 7 chapter boxes (alternating colors)
    chapters = [
        "1장 JIT → JIC 패러다임 전환",
        "2장 Kraljic Matrix 프레임워크",
        "3장 차별화 전략",
        "4장 계획 방법론",
        "5장 통합 KPI 프레임워크",
        "6장 산업별 적용 사례",
        "7장 9회차 학습 여정"
    ]

    y_start = 2.00
    box_height = 0.65
    gap = 0.05
    shape_count = 0

    for i, chapter in enumerate(chapters):
        y = y_start + i * (box_height + gap)

        # Alternating background color
        bg_color = COLOR_VERY_LIGHT_GRAY if i % 2 == 0 else COLOR_WHITE

        # Box
        add_rectangle(
            slide, 1.00, y, 8.83, box_height,
            fill_color=bg_color,
            border_color=COLOR_LIGHT_GRAY,
            border_width=1
        )
        shape_count += 1

        # Chapter number (large)
        add_text_box(
            slide, 1.20, y + 0.10, 1.00, 0.45,
            f"{i+1}장", font_size=18, bold=True, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        # Chapter title
        add_text_box(
            slide, 2.40, y + 0.15, 6.50, 0.40,
            chapter.split(' ', 1)[1], font_size=14, bold=False, color=COLOR_BLACK
        )
        shape_count += 1

    print(f"✓ Slide 2: TOC ({shape_count} shapes)")
    return slide

# ============================================================================
# SLIDE 3: CHAPTER 1 DIVIDER (5-10 shapes)
# ============================================================================

def create_slide_3_chapter1_divider(prs):
    """Slide 3: Chapter 1 divider with large number"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    shape_count = 0

    # Large chapter number (72pt)
    add_text_box(
        slide, 0.50, 2.00, 3.00, 2.00,
        "1장", font_size=72, bold=True, color=COLOR_DARK_GRAY,
        align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Chapter title (32pt)
    add_text_box(
        slide, 3.80, 2.50, 6.50, 1.50,
        "JIT → JIC\n패러다임 전환",
        font_size=32, bold=True, color=COLOR_BLACK
    )
    shape_count += 1

    # Decorative line
    from pptx.enum.shapes import MSO_CONNECTOR
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(3.80), Inches(4.20),
        Inches(10.00), Inches(4.20)
    )
    connector.line.color.rgb = COLOR_LIGHT_GRAY
    connector.line.width = Pt(3)
    shape_count += 1

    # Subtitle
    add_text_box(
        slide, 3.80, 4.50, 6.00, 0.80,
        "Just-In-Time에서 Just-In-Case로\n재고 관리 전략의 근본적 변화",
        font_size=14, color=COLOR_MED_GRAY
    )
    shape_count += 1

    print(f"✓ Slide 3: Chapter 1 Divider ({shape_count} shapes)")
    return slide

# ============================================================================
# SLIDE 4: JIT TIMELINE (40-50 shapes) - CRITICAL!
# ============================================================================

def create_slide_4_jit_timeline(prs):
    """Slide 4: JIT Timeline with 40-50 shapes"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.1 JIT의 영광과 몰락", slide_num=4)
    add_governing_message(
        slide,
        "JIT 방식은 40년간 제조업의 표준이었으나 2020년 팬데믹으로 치명적 약점이 드러났습니다."
    )

    shape_count = 0

    # Timeline structure:
    # 1. Horizontal timeline arrow
    # 2. 5 periods: 1970s, 1990s, 2000s, 2010s, 2020
    # 3. Event boxes with arrows
    # 4. Impact indicators

    # Main timeline arrow (horizontal)
    from pptx.enum.shapes import MSO_CONNECTOR
    timeline_y = 3.00
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(1.00), Inches(timeline_y),
        Inches(10.00), Inches(timeline_y)
    )
    connector.line.color.rgb = COLOR_DARK_GRAY
    connector.line.width = Pt(3)
    connector.line.end_arrow_type = 2
    shape_count += 1

    # 5 time periods
    periods = [
        {"year": "1970s", "x": 1.50, "event": "JIT 탄생",
         "desc": "도요타 생산 방식\n무재고 경영", "icon": "↑"},
        {"year": "1990s", "x": 3.25, "event": "글로벌 확산",
         "desc": "미국·유럽 제조업\n표준 채택", "icon": "↑↑"},
        {"year": "2000s", "x": 5.00, "event": "디지털화",
         "desc": "ERP 시스템\n실시간 가시성", "icon": "↑↑↑"},
        {"year": "2010s", "x": 6.75, "event": "최적화",
         "desc": "공급망 최적화\n원가 절감 극대화", "icon": "↑↑↑"},
        {"year": "2020", "x": 8.50, "event": "팬데믹 쇼크",
         "desc": "공급망 마비\nJIT 붕괴", "icon": "↓↓↓"}
    ]

    for period in periods:
        x = period["x"]

        # Circle marker on timeline
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x - 0.10), Inches(timeline_y - 0.10),
            Inches(0.20), Inches(0.20)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = COLOR_DARK_GRAY
        circle.line.fill.background()
        shape_count += 1

        # Year label (below timeline)
        add_text_box(
            slide, x - 0.30, timeline_y + 0.20, 0.60, 0.25,
            period["year"], font_size=10, bold=True,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Event box (above timeline)
        box_y = timeline_y - 1.20
        add_rectangle(
            slide, x - 0.50, box_y, 1.00, 0.50,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Event title
        add_text_box(
            slide, x - 0.45, box_y + 0.05, 0.90, 0.20,
            period["event"], font_size=11, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connecting line (from box to timeline)
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x), Inches(box_y + 0.50),
            Inches(x), Inches(timeline_y - 0.10)
        )
        conn.line.color.rgb = COLOR_LIGHT_GRAY
        conn.line.width = Pt(1)
        shape_count += 1

        # Description box (below timeline)
        desc_y = timeline_y + 0.60
        add_rectangle(
            slide, x - 0.50, desc_y, 1.00, 0.70,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=1
        )
        shape_count += 1

        # Description text (9pt - small!)
        add_text_box(
            slide, x - 0.45, desc_y + 0.08, 0.90, 0.60,
            period["desc"], font_size=9, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Impact indicator
        add_text_box(
            slide, x - 0.25, desc_y + 0.55, 0.50, 0.20,
            period["icon"], font_size=10, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    # Right side: Key insights (Toy Page pattern)
    insights_x = 10.20
    insights_y = 2.00

    # Insights box
    add_rectangle(
        slide, insights_x, insights_y, 0.00, 4.50,  # Width 0 = invisible
        fill_color=COLOR_WHITE
    )

    # "시사점" header
    add_text_box(
        slide, insights_x - 2.20, insights_y, 2.00, 0.30,
        "시사점", font_size=12, bold=True, color=COLOR_BLACK
    )
    shape_count += 1

    # Insight points (9pt body text)
    insights = [
        "JIT는 40년간 제조업 혁신의 상징",
        "원가 절감과 효율성 극대화 달성",
        "2020년 팬데믹으로 근본적 한계 노출",
        "공급망 리스크에 극도로 취약",
        "JIC로의 패러다임 전환 불가피"
    ]

    insight_y = insights_y + 0.40
    for i, insight in enumerate(insights):
        # Bullet point
        add_text_box(
            slide, insights_x - 2.10, insight_y, 0.15, 0.20,
            "•", font_size=10, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        # Insight text
        add_text_box(
            slide, insights_x - 1.90, insight_y, 1.80, 0.30,
            insight, font_size=9, bold=False, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        insight_y += 0.35

    print(f"✓ Slide 4: JIT Timeline ({shape_count} shapes)")
    return slide

# ============================================================================
# SLIDE 5: PANDEMIC WEAKNESSES (30-40 shapes)
# ============================================================================

def create_slide_5_pandemic(prs):
    """Slide 5: Pandemic exposed JIT weaknesses"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.2 팬데믹이 드러낸 JIT의 약점", slide_num=5)
    add_governing_message(
        slide,
        "글로벌 공급망 마비로 JIT의 3대 위험(재고 부족, 공급 중단, 생산 마비)이 현실화되었습니다."
    )

    shape_count = 0

    # Left side: Problem diagram
    # Central problem box
    center_x, center_y = 3.50, 3.80
    add_rectangle(
        slide, center_x, center_y, 2.00, 0.80,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=2
    )
    shape_count += 1

    add_text_box(
        slide, center_x + 0.10, center_y + 0.20, 1.80, 0.40,
        "2020 팬데믹\n공급망 마비", font_size=13, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # 3 major problems (around center)
    problems = [
        {"x": 1.00, "y": 2.50, "title": "재고 부족",
         "details": ["안전재고 없음", "즉각 결품 발생", "생산 차질"]},
        {"x": 1.00, "y": 4.50, "title": "공급 중단",
         "details": ["단일 공급원", "대체품 없음", "긴급 조달 불가"]},
        {"x": 6.00, "y": 3.30, "title": "생산 마비",
         "details": ["라인 가동 중단", "납기 지연", "매출 손실"]}
    ]

    for prob in problems:
        # Problem box
        add_rectangle(
            slide, prob["x"], prob["y"], 1.80, 1.20,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Title
        add_text_box(
            slide, prob["x"] + 0.10, prob["y"] + 0.10, 1.60, 0.30,
            prob["title"], font_size=12, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Details (9pt small text)
        detail_y = prob["y"] + 0.45
        for detail in prob["details"]:
            add_text_box(
                slide, prob["x"] + 0.15, detail_y, 0.10, 0.20,
                "•", font_size=9, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            add_text_box(
                slide, prob["x"] + 0.30, detail_y, 1.45, 0.20,
                detail, font_size=9, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            detail_y += 0.22

        # Arrow to center
        end_x = center_x if prob["x"] < center_x else center_x + 2.00
        end_y = center_y + 0.40

        add_arrow(
            slide,
            prob["x"] + 0.90 if prob["x"] < center_x else prob["x"],
            prob["y"] + 0.60,
            end_x, end_y,
            color=COLOR_MED_GRAY, width=2
        )
        shape_count += 1

    # Right side: Industry impact examples (9pt text)
    impact_x = 7.50
    impact_y = 2.00

    add_text_box(
        slide, impact_x, impact_y, 2.80, 0.30,
        "산업별 피해 사례", font_size=12, bold=True, color=COLOR_BLACK
    )
    shape_count += 1

    industries = [
        {"name": "자동차", "impact": "반도체 부족으로 감산"},
        {"name": "전자", "impact": "부품 결품으로 출시 지연"},
        {"name": "의료", "impact": "마스크·장갑 공급 중단"},
        {"name": "식품", "impact": "포장재 부족으로 생산 차질"}
    ]

    ind_y = impact_y + 0.50
    for ind in industries:
        # Industry box
        add_rectangle(
            slide, impact_x, ind_y, 2.80, 0.65,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=1
        )
        shape_count += 1

        # Industry name (10pt)
        add_text_box(
            slide, impact_x + 0.10, ind_y + 0.08, 0.80, 0.25,
            ind["name"], font_size=10, bold=True, color=COLOR_BLACK
        )
        shape_count += 1

        # Impact (9pt small)
        add_text_box(
            slide, impact_x + 0.10, ind_y + 0.35, 2.60, 0.25,
            ind["impact"], font_size=9, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        ind_y += 0.75

    print(f"✓ Slide 5: Pandemic ({shape_count} shapes)")
    return slide

# ============================================================================
# MAIN GENERATION FUNCTION
# ============================================================================

def main():
    """Generate Part 1 PPTX (First 5 slides for testing)"""
    print("=== Part 1 PPTX Generation (Slides 1-5) ===")
    print("High-quality implementation following S4HANA standards\n")

    prs = create_presentation()

    # Generate first 5 slides
    create_slide_1_cover(prs)
    create_slide_2_toc(prs)
    create_slide_3_chapter1_divider(prs)
    create_slide_4_jit_timeline(prs)
    create_slide_5_pandemic(prs)

    # Save
    output_path = "/home/user/Kraljic_Course/Part1_Session1_StrategicInventory.pptx"
    prs.save(output_path)

    print(f"\n=== Generation Complete ===")
    print(f"Output: {output_path}")
    print(f"Slides generated: 5")
    print(f"\nNext: Run verification script")

    return output_path

if __name__ == "__main__":
    main()
