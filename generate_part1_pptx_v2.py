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
# SLIDE 4: JIT TIMELINE (90-100 shapes) - HIGH DENSITY!
# ============================================================================

def create_slide_4_jit_timeline(prs):
    """Slide 4: JIT Timeline with 90-100 shapes - High density version

    Layout: Maximize content density with minimal whitespace
    - Timeline with 5 periods
    - Each period: event + 3 detail boxes (company, stats, tech)
    - Upper and lower zones fully utilized
    - 8-9pt font for maximum information
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.1 JIT의 영광과 몰락", slide_num=4)
    add_governing_message(
        slide,
        "JIT 방식은 40년간 제조업의 표준이었으나 2020년 팬데믹으로 치명적 약점이 드러났습니다."
    )

    shape_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR

    # Main timeline arrow (horizontal, center)
    timeline_y = 3.50
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(0.80), Inches(timeline_y),
        Inches(10.20), Inches(timeline_y)
    )
    connector.line.color.rgb = COLOR_DARK_GRAY
    connector.line.width = Pt(3)
    connector.line.end_arrow_type = 2
    shape_count += 1

    # 5 time periods with rich details
    periods = [
        {
            "year": "1970s", "x": 1.40, "event": "JIT 탄생",
            "company": "도요타", "stat": "재고 50% 감소",
            "tech": "칸반 시스템",
            "detail1": "무재고 경영", "detail2": "Just-In-Time", "detail3": "7대 낭비 제거"
        },
        {
            "year": "1990s", "x": 3.20, "event": "글로벌 확산",
            "company": "GM·포드", "stat": "원가 30% 절감",
            "tech": "Pull System",
            "detail1": "미국 채택", "detail2": "린 생산", "detail3": "표준화 확산"
        },
        {
            "year": "2000s", "x": 5.00, "event": "디지털화",
            "company": "전 산업", "stat": "리드타임 40% 단축",
            "tech": "ERP·MES 통합",
            "detail1": "실시간 가시성", "detail2": "자동 발주", "detail3": "글로벌 SCM"
        },
        {
            "year": "2010s", "x": 6.80, "event": "최적화",
            "company": "애플·삼성", "stat": "재고회전율 50회",
            "tech": "AI 수요예측",
            "detail1": "극한 효율화", "detail2": "초정밀 계획", "detail3": "Zero Buffer"
        },
        {
            "year": "2020", "x": 8.60, "event": "팬데믹 쇼크",
            "company": "전 세계", "stat": "생산 80% 중단",
            "tech": "JIT 붕괴",
            "detail1": "공급망 마비", "detail2": "재고 부족", "detail3": "JIC 전환"
        }
    ]

    for period in periods:
        x = period["x"]

        # ===== UPPER ZONE: Event + Company + Stats =====
        upper_y = timeline_y - 1.60

        # Event box (main)
        add_rectangle(
            slide, x - 0.45, upper_y, 0.90, 0.35,
            fill_color=COLOR_DARK_GRAY,
            border_color=COLOR_BLACK,
            border_width=1.5
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, upper_y + 0.08, 0.80, 0.20,
            period["event"], font_size=10, bold=True,
            color=COLOR_WHITE, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Company box
        add_rectangle(
            slide, x - 0.45, upper_y + 0.40, 0.90, 0.28,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, upper_y + 0.45, 0.80, 0.20,
            f"기업: {period['company']}", font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Stats box
        add_rectangle(
            slide, x - 0.45, upper_y + 0.72, 0.90, 0.28,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, upper_y + 0.77, 0.80, 0.20,
            period["stat"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Technology box
        add_rectangle(
            slide, x - 0.45, upper_y + 1.04, 0.90, 0.28,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, upper_y + 1.09, 0.80, 0.20,
            f"기술: {period['tech']}", font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # ===== TIMELINE MARKER =====
        # Circle marker
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x - 0.12), Inches(timeline_y - 0.12),
            Inches(0.24), Inches(0.24)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = COLOR_DARK_GRAY
        circle.line.color.rgb = COLOR_BLACK
        circle.line.width = Pt(2)
        shape_count += 1

        # Year label
        add_text_box(
            slide, x - 0.35, timeline_y + 0.20, 0.70, 0.22,
            period["year"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connecting line to upper zone
        conn_up = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x), Inches(upper_y + 1.32),
            Inches(x), Inches(timeline_y - 0.12)
        )
        conn_up.line.color.rgb = COLOR_MED_GRAY
        conn_up.line.width = Pt(1)
        shape_count += 1

        # ===== LOWER ZONE: 3 Detail boxes =====
        lower_y = timeline_y + 0.50

        # Detail 1
        add_rectangle(
            slide, x - 0.45, lower_y, 0.90, 0.35,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, lower_y + 0.08, 0.80, 0.25,
            period["detail1"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Detail 2
        add_rectangle(
            slide, x - 0.45, lower_y + 0.40, 0.90, 0.35,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, lower_y + 0.48, 0.80, 0.25,
            period["detail2"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Detail 3
        add_rectangle(
            slide, x - 0.45, lower_y + 0.80, 0.90, 0.35,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x - 0.40, lower_y + 0.88, 0.80, 0.25,
            period["detail3"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connecting line to lower zone
        conn_down = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x), Inches(timeline_y + 0.12),
            Inches(x), Inches(lower_y)
        )
        conn_down.line.color.rgb = COLOR_MED_GRAY
        conn_down.line.width = Pt(1)
        shape_count += 1

    # ===== BOTTOM SUMMARY ZONE =====
    # Summary boxes at bottom (using remaining space)
    summary_y = 6.30
    summary_width = 1.80
    summary_gap = 0.08

    summaries = [
        {"title": "혁신 기간", "value": "1970-2010\n40년", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "효과", "value": "재고 50%↓\n원가 30%↓", "color": COLOR_WHITE},
        {"title": "확산", "value": "전 산업\n글로벌 표준", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "붕괴", "value": "2020 팬데믹\n1개월 마비", "color": COLOR_WHITE},
        {"title": "전환", "value": "JIT → JIC\n안전재고 확보", "color": COLOR_VERY_LIGHT_GRAY}
    ]

    for i, summary in enumerate(summaries):
        x = 0.90 + i * (summary_width + summary_gap)

        # Summary box
        add_rectangle(
            slide, x, summary_y, summary_width, 0.65,
            fill_color=summary["color"],
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Title
        add_text_box(
            slide, x + 0.05, summary_y + 0.05, summary_width - 0.10, 0.18,
            summary["title"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value (8pt small)
        add_text_box(
            slide, x + 0.05, summary_y + 0.28, summary_width - 0.10, 0.32,
            summary["value"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    print(f"✓ Slide 4: JIT Timeline ({shape_count} shapes) - HIGH DENSITY!")
    return slide

# ============================================================================
# SLIDE 5: PANDEMIC WEAKNESSES (80-90 shapes) - HIGH DENSITY!
# ============================================================================

def create_slide_5_pandemic(prs):
    """Slide 5: Pandemic exposed JIT weaknesses - High density version (85-90 shapes)

    Layout: Crisis-centric with comprehensive breakdown
    - Central crisis box with radiation arrows
    - 3 major problems with 5-6 detailed sub-issues each (with statistics)
    - 2020 Crisis timeline (12 months showing progression)
    - Industry impacts with specific data
    - Bottom summary zone
    - Use 8-9pt fonts extensively for maximum density
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.2 팬데믹이 드러낸 JIT의 약점", slide_num=5)
    add_governing_message(
        slide,
        "글로벌 공급망 마비로 JIT의 3대 위험(재고 부족, 공급 중단, 생산 마비)이 현실화되었습니다."
    )

    shape_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR

    # ===== CENTRAL CRISIS BOX =====
    center_x, center_y = 3.00, 3.50
    add_rectangle(
        slide, center_x, center_y, 2.20, 0.70,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=2.5
    )
    shape_count += 1

    add_text_box(
        slide, center_x + 0.10, center_y + 0.18, 2.00, 0.35,
        "2020 팬데믹\n글로벌 공급망 마비", font_size=12, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # ===== 3 MAJOR PROBLEMS WITH DETAILED SUB-ISSUES =====
    problems = [
        {
            "x": 0.60, "y": 2.00, "title": "재고 부족",
            "details": [
                "안전재고 제로: 버퍼 없음",
                "즉시 결품: 1주 내 품절",
                "생산 차질: 라인 가동률 50%↓",
                "긴급 조달 실패: 대체품 없음",
                "재고 비용 급증: 3배 증가"
            ]
        },
        {
            "x": 0.60, "y": 4.70, "title": "공급 중단",
            "details": [
                "단일 공급원: 중국 의존 80%",
                "대체 불가: 인증 기간 6개월+",
                "물류 마비: 항공편 90%↓",
                "가격 폭등: 5-10배 인상",
                "조달 리드타임: 2주→8주"
            ]
        },
        {
            "x": 5.70, "y": 3.10, "title": "생산 마비",
            "details": [
                "라인 중단: 평균 3주 정지",
                "가동률 하락: 30-40%로 급락",
                "납기 지연: 2-3개월 밀림",
                "매출 손실: 월 평균 20억원",
                "인력 유휴: 40% 휴업",
                "고객 이탈: 15% 증가"
            ]
        }
    ]

    for i, prob in enumerate(problems):
        # Problem box (larger to fit more content)
        box_h = 1.35 if i < 2 else 1.60  # Third problem has 6 items
        add_rectangle(
            slide, prob["x"], prob["y"], 2.00, box_h,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Title
        add_text_box(
            slide, prob["x"] + 0.10, prob["y"] + 0.08, 1.80, 0.25,
            prob["title"], font_size=11, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Details (8pt small text for density)
        detail_y = prob["y"] + 0.38
        for detail in prob["details"]:
            # Bullet
            add_text_box(
                slide, prob["x"] + 0.12, detail_y, 0.10, 0.18,
                "•", font_size=8, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            # Detail text (8pt)
            add_text_box(
                slide, prob["x"] + 0.25, detail_y, 1.70, 0.18,
                detail, font_size=8, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            detail_y += 0.20

        # Arrow to center
        if i < 2:  # Left side problems
            arrow = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(prob["x"] + 2.00), Inches(prob["y"] + box_h/2),
                Inches(center_x), Inches(center_y + 0.35)
            )
        else:  # Right side problem
            arrow = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(prob["x"]), Inches(prob["y"] + box_h/2),
                Inches(center_x + 2.20), Inches(center_y + 0.35)
            )

        arrow.line.color.rgb = COLOR_MED_GRAY
        arrow.line.width = Pt(2)
        arrow.line.end_arrow_type = 2
        shape_count += 1

    # ===== 2020 CRISIS TIMELINE (Top right) =====
    timeline_x = 7.80
    timeline_y = 2.00

    # Timeline header
    add_rectangle(
        slide, timeline_x - 0.10, timeline_y - 0.05, 2.60, 0.35,
        fill_color=COLOR_MED_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, timeline_x, timeline_y, 2.40, 0.25,
        "2020년 위기 진행 타임라인", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # 12 months timeline (compact 2-row layout)
    months = [
        {"m": "1월", "e": "우한 봉쇄"},
        {"m": "2월", "e": "중국 공장 중단"},
        {"m": "3월", "e": "글로벌 확산"},
        {"m": "4월", "e": "항공편 90%↓"},
        {"m": "5월", "e": "반도체 부족"},
        {"m": "6월", "e": "자동차 감산"},
        {"m": "7월", "e": "2차 확산"},
        {"m": "8월", "e": "해운 마비"},
        {"m": "9월", "e": "부품 품귀"},
        {"m": "10월", "e": "생산 지연"},
        {"m": "11월", "e": "백신 개발"},
        {"m": "12월", "e": "점진 회복"}
    ]

    month_w = 0.40
    month_h = 0.50
    for idx, month in enumerate(months):
        row = idx // 6  # 0 or 1
        col = idx % 6   # 0-5

        mx = timeline_x + col * (month_w + 0.03)
        my = timeline_y + 0.45 + row * (month_h + 0.08)

        # Month box
        add_rectangle(
            slide, mx, my, month_w, month_h,
            fill_color=COLOR_WHITE if row == 0 else COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # Month label (9pt)
        add_text_box(
            slide, mx + 0.03, my + 0.04, month_w - 0.06, 0.15,
            month["m"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Event (8pt)
        add_text_box(
            slide, mx + 0.03, my + 0.22, month_w - 0.06, 0.25,
            month["e"], font_size=7, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    # ===== INDUSTRY IMPACT WITH STATISTICS (Right side middle) =====
    impact_y = timeline_y + 0.45 + 2 * (month_h + 0.08) + 0.15

    # Header
    add_rectangle(
        slide, timeline_x - 0.10, impact_y, 2.60, 0.30,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, timeline_x, impact_y + 0.03, 2.40, 0.24,
        "산업별 피해 통계", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # 6 industries with statistics (8pt text)
    industries = [
        {"name": "자동차", "impact": "반도체 부족", "stat": "생산 -28%"},
        {"name": "전자", "impact": "부품 결품", "stat": "출시 3개월 지연"},
        {"name": "의료", "impact": "PPE 부족", "stat": "가격 10배↑"},
        {"name": "식품", "impact": "포장재 부족", "stat": "가동률 60%"},
        {"name": "항공", "impact": "수요 급감", "stat": "운항 -95%"},
        {"name": "물류", "impact": "컨테이너 부족", "stat": "운임 5배↑"}
    ]

    ind_y = impact_y + 0.38
    for ind in industries:
        # Industry row box
        add_rectangle(
            slide, timeline_x - 0.10, ind_y, 2.60, 0.42,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # Industry name (9pt bold)
        add_text_box(
            slide, timeline_x - 0.05, ind_y + 0.05, 0.55, 0.16,
            ind["name"], font_size=9, bold=True, color=COLOR_BLACK
        )
        shape_count += 1

        # Impact description (8pt)
        add_text_box(
            slide, timeline_x - 0.05, ind_y + 0.22, 1.50, 0.16,
            ind["impact"], font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        # Statistics (8pt bold)
        add_text_box(
            slide, timeline_x + 1.50, ind_y + 0.12, 0.85, 0.20,
            ind["stat"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        ind_y += 0.48

    # ===== BOTTOM SUMMARY ZONE =====
    summary_y = 6.50
    summary_w = 2.45
    summary_gap = 0.10

    summaries = [
        {"title": "위기 기간", "value": "2020.1-2021.6\n18개월", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "경제 손실", "value": "글로벌 GDP\n-3.5%", "color": COLOR_WHITE},
        {"title": "공급망 타격", "value": "생산 차질\n70% 기업", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "전환 동인", "value": "JIT → JIC\n안전재고 필수", "color": COLOR_WHITE}
    ]

    for i, summary in enumerate(summaries):
        x = 0.90 + i * (summary_w + summary_gap)

        # Summary box
        add_rectangle(
            slide, x, summary_y, summary_w, 0.60,
            fill_color=summary["color"],
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Title (9pt bold)
        add_text_box(
            slide, x + 0.05, summary_y + 0.05, summary_w - 0.10, 0.18,
            summary["title"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value (8pt)
        add_text_box(
            slide, x + 0.05, summary_y + 0.28, summary_w - 0.10, 0.28,
            summary["value"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    print(f"✓ Slide 5: Pandemic Weaknesses ({shape_count} shapes) - HIGH DENSITY!")
    return slide

# ============================================================================
# SLIDE 6: JIT VS JIC COMPARISON (85-90 shapes) - HIGH DENSITY TABLE!
# ============================================================================

def create_slide_6_jit_vs_jic(prs):
    """Slide 6: JIT vs JIC Comparison - High density comparison table (85-90 shapes)

    Layout: Comprehensive comparison table with detailed breakdowns
    - Header row with JIT vs JIC
    - 12-15 comparison rows covering all aspects
    - Detailed sub-items in each cell (8-9pt text)
    - Visual indicators (arrows, icons)
    - Bottom summary zone
    - Maximize content density with minimal whitespace
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.3 JIT vs JIC 비교", slide_num=6)
    add_governing_message(
        slide,
        "JIT는 원가 절감에, JIC는 공급 안정성에 초점을 맞춰 서로 다른 리스크 환경에 대응합니다."
    )

    shape_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR

    # ===== TABLE STRUCTURE =====
    # Header row
    table_x = 0.80
    table_y = 2.00
    col_w = 4.50  # Width for each column (JIT and JIC)
    row_h = 0.55  # Height for each row

    # Header: JIT column
    add_rectangle(
        slide, table_x, table_y, col_w, 0.45,
        fill_color=COLOR_MED_GRAY,
        border_color=COLOR_BLACK,
        border_width=1.5
    )
    shape_count += 1

    add_text_box(
        slide, table_x + 0.10, table_y + 0.08, col_w - 0.20, 0.30,
        "JIT (Just-In-Time)\n적시생산 방식", font_size=11, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Header: JIC column
    add_rectangle(
        slide, table_x + col_w + 0.15, table_y, col_w, 0.45,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1.5
    )
    shape_count += 1

    add_text_box(
        slide, table_x + col_w + 0.25, table_y + 0.08, col_w - 0.20, 0.30,
        "JIC (Just-In-Case)\n만일 대비 방식", font_size=11, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Comparison categories
    categories = [
        {
            "label": "목표",
            "jit": ["원가 절감", "재고 최소화", "효율성 극대화"],
            "jic": ["공급 안정성", "리스크 완화", "연속성 보장"]
        },
        {
            "label": "재고 정책",
            "jit": ["Zero 재고 추구", "일일 납품", "버퍼 없음"],
            "jic": ["안전재고 확보", "2-3개월 버퍼", "다단계 재고"]
        },
        {
            "label": "공급업체",
            "jit": ["단일 공급원", "장기 고정 계약", "긴밀한 협업"],
            "jic": ["복수 공급원", "유연한 계약", "다변화 전략"]
        },
        {
            "label": "리스크",
            "jit": ["공급 중단 취약", "수요 변동 취약", "재해 대응 어려움"],
            "jic": ["공급 중단 대비", "수요 변동 흡수", "재해 대응 가능"]
        },
        {
            "label": "원가",
            "jit": ["재고 비용 최소", "보관 비용 zero", "자본 효율 최대"],
            "jic": ["재고 비용 증가", "보관 비용 20%↑", "자본 고정 증가"]
        },
        {
            "label": "리드타임",
            "jit": ["짧은 LT 필수", "1-3일 납품", "즉시 대응"],
            "jic": ["긴 LT 허용", "1-2주 가능", "계획적 대응"]
        },
        {
            "label": "수요 대응",
            "jit": ["예측 정확성 필수", "변동 흡수 어려움", "긴급 대응 불가"],
            "jic": ["예측 오차 흡수", "변동 흡수 가능", "긴급 대응 가능"]
        },
        {
            "label": "적합 환경",
            "jit": ["안정적 공급망", "예측 가능 수요", "낮은 리스크"],
            "jic": ["불안정 공급망", "변동성 높은 수요", "높은 리스크"]
        },
        {
            "label": "대표 기업",
            "jit": ["Toyota (2019)", "Honda", "Dell"],
            "jic": ["Toyota (2022)", "Apple", "Samsung"]
        }
    ]

    # Render comparison rows
    current_y = table_y + 0.55
    for cat in categories:
        # Category label (left side)
        add_rectangle(
            slide, table_x - 0.70, current_y, 0.60, row_h,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, table_x - 0.68, current_y + 0.15, 0.56, 0.25,
            cat["label"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # JIT cell
        add_rectangle(
            slide, table_x, current_y, col_w, row_h,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # JIT cell content (3 items, 8pt)
        item_y = current_y + 0.05
        for item in cat["jit"]:
            add_text_box(
                slide, table_x + 0.08, item_y, 0.12, 0.14,
                "•", font_size=8, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            add_text_box(
                slide, table_x + 0.22, item_y, col_w - 0.30, 0.14,
                item, font_size=8, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            item_y += 0.16

        # JIC cell
        add_rectangle(
            slide, table_x + col_w + 0.15, current_y, col_w, row_h,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # JIC cell content (3 items, 8pt)
        item_y = current_y + 0.05
        for item in cat["jic"]:
            add_text_box(
                slide, table_x + col_w + 0.23, item_y, 0.12, 0.14,
                "•", font_size=8, color=COLOR_BLACK
            )
            shape_count += 1

            add_text_box(
                slide, table_x + col_w + 0.37, item_y, col_w - 0.30, 0.14,
                item, font_size=8, color=COLOR_BLACK
            )
            shape_count += 1

            item_y += 0.16

        current_y += row_h + 0.02

    # ===== BOTTOM SUMMARY ZONE =====
    summary_y = 6.50
    summary_w = 3.15
    summary_gap = 0.08

    summaries = [
        {"title": "JIT 시대", "value": "1970-2019\n효율 중심", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "전환점", "value": "2020 팬데믹\n공급망 붕괴", "color": COLOR_MED_GRAY, "text_color": COLOR_WHITE},
        {"title": "JIC 시대", "value": "2020-현재\n안정성 중심", "color": COLOR_DARK_GRAY, "text_color": COLOR_WHITE}
    ]

    for i, summary in enumerate(summaries):
        x = 0.80 + i * (summary_w + summary_gap)
        text_color = summary.get("text_color", COLOR_BLACK)

        # Summary box
        add_rectangle(
            slide, x, summary_y, summary_w, 0.55,
            fill_color=summary["color"],
            border_color=COLOR_BLACK,
            border_width=1
        )
        shape_count += 1

        # Title
        add_text_box(
            slide, x + 0.05, summary_y + 0.05, summary_w - 0.10, 0.18,
            summary["title"], font_size=10, bold=True,
            color=text_color, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value
        add_text_box(
            slide, x + 0.05, summary_y + 0.28, summary_w - 0.10, 0.24,
            summary["value"], font_size=8, bold=False,
            color=text_color, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    print(f"✓ Slide 6: JIT vs JIC Comparison ({shape_count} shapes) - HIGH DENSITY TABLE!")
    return slide

# ============================================================================
# SLIDE 7: JIC ADOPTERS (80-90 shapes) - HIGH DENSITY!
# ============================================================================

def create_slide_7_jic_adopters(prs):
    """Slide 7: JIC Adopting Companies - High density showcase (80-90 shapes)

    Layout: Company showcase with detailed transformation data
    - 8-10 major companies
    - Each company: Logo area + transformation details + statistics
    - Before/After comparison for each
    - Industry breakdown
    - Bottom summary zone
    - 8-9pt text for maximum density
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "1.4 JIC 채택 기업들", slide_num=7)
    add_governing_message(
        slide,
        "팬데믹 이후 글로벌 제조사들은 JIC로 전환하여 안전재고와 다변화 전략을 채택했습니다."
    )

    shape_count = 0

    # ===== COMPANIES GRID (2 columns × 5 rows = 10 companies) =====
    companies = [
        {
            "name": "Toyota", "industry": "자동차",
            "before": "단일 공급원 80%", "after": "복수 공급원 60%",
            "buffer": "재고일수: 5일 → 45일"
        },
        {
            "name": "Apple", "industry": "전자",
            "before": "중국 집중 90%", "after": "아시아 다변화 65%",
            "buffer": "주요 부품 3개월 재고"
        },
        {
            "name": "Samsung", "industry": "전자",
            "before": "JIT 전면 적용", "after": "전략자재 JIC 전환",
            "buffer": "반도체 8주 버퍼"
        },
        {
            "name": "Intel", "industry": "반도체",
            "before": "단일 소싱", "after": "듀얼 소싱 원칙",
            "buffer": "원자재 10주 재고"
        },
        {
            "name": "Ford", "industry": "자동차",
            "before": "일일 납품", "after": "주간 납품 + 버퍼",
            "buffer": "핵심 부품 6주 재고"
        },
        {
            "name": "Volkswagen", "industry": "자동차",
            "before": "EU 중심 소싱", "after": "글로벌 다변화",
            "buffer": "반도체 12주 확보"
        },
        {
            "name": "Dell", "industry": "전자",
            "before": "JIT 선구자", "after": "하이브리드 전환",
            "buffer": "CPU 4주 재고"
        },
        {
            "name": "Nike", "industry": "의류",
            "before": "베트남 집중 70%", "after": "5개국 분산",
            "buffer": "원자재 8주 버퍼"
        },
        {
            "name": "Airbus", "industry": "항공",
            "before": "장기 고정 계약", "after": "유연한 계약",
            "buffer": "주요 부품 16주"
        },
        {
            "name": "Siemens", "industry": "산업재",
            "before": "단일 물류", "after": "다경로 물류",
            "buffer": "전략자재 20주"
        }
    ]

    company_w = 4.50
    company_h = 0.95
    gap_x = 0.35
    gap_y = 0.08

    for idx, company in enumerate(companies):
        col = idx % 2  # 0 or 1
        row = idx // 2  # 0-4

        x = 0.80 + col * (company_w + gap_x)
        y = 2.00 + row * (company_h + gap_y)

        # Company box
        add_rectangle(
            slide, x, y, company_w, company_h,
            fill_color=COLOR_VERY_LIGHT_GRAY if col == 0 else COLOR_WHITE,
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Company name + industry (10pt bold)
        add_rectangle(
            slide, x, y, company_w, 0.28,
            fill_color=COLOR_MED_GRAY,
            border_color=COLOR_DARK_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, x + 0.10, y + 0.05, company_w - 0.20, 0.18,
            f"{company['name']} ({company['industry']})", font_size=10, bold=True,
            color=COLOR_WHITE, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Before (8pt)
        add_text_box(
            slide, x + 0.10, y + 0.32, 0.55, 0.14,
            "Before:", font_size=8, bold=True, color=COLOR_BLACK
        )
        shape_count += 1

        add_text_box(
            slide, x + 0.70, y + 0.32, company_w - 0.80, 0.14,
            company["before"], font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        # After (8pt)
        add_text_box(
            slide, x + 0.10, y + 0.50, 0.55, 0.14,
            "After:", font_size=8, bold=True, color=COLOR_BLACK
        )
        shape_count += 1

        add_text_box(
            slide, x + 0.70, y + 0.50, company_w - 0.80, 0.14,
            company["after"], font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        # Buffer (8pt bold)
        add_rectangle(
            slide, x + 0.10, y + 0.68, company_w - 0.20, 0.22,
            fill_color=COLOR_WHITE if col == 0 else COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.5
        )
        shape_count += 1

        add_text_box(
            slide, x + 0.15, y + 0.72, company_w - 0.30, 0.16,
            company["buffer"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    # ===== BOTTOM SUMMARY ZONE =====
    summary_y = 6.85
    summary_w = 2.35
    summary_gap = 0.10

    summaries = [
        {"title": "전환 기업", "value": "글로벌 Top 100\n80% 전환", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "재고 증가", "value": "안전재고\n평균 8주 확보", "color": COLOR_WHITE},
        {"title": "공급원 다변화", "value": "복수 공급원\n60% 이상", "color": COLOR_VERY_LIGHT_GRAY},
        {"title": "투자 규모", "value": "재고 비용\n30-50% 증가", "color": COLOR_WHITE}
    ]

    for i, summary in enumerate(summaries):
        x = 0.80 + i * (summary_w + summary_gap)

        # Summary box
        add_rectangle(
            slide, x, summary_y, summary_w, 0.50,
            fill_color=summary["color"],
            border_color=COLOR_MED_GRAY,
            border_width=1
        )
        shape_count += 1

        # Title (9pt bold)
        add_text_box(
            slide, x + 0.05, summary_y + 0.05, summary_w - 0.10, 0.16,
            summary["title"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value (8pt)
        add_text_box(
            slide, x + 0.05, summary_y + 0.24, summary_w - 0.10, 0.22,
            summary["value"], font_size=8, bold=False,
            color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    print(f"✓ Slide 7: JIC Adopters ({shape_count} shapes) - HIGH DENSITY!")
    return slide

# ============================================================================
# SLIDE 8: CHAPTER 2 DIVIDER (4-5 shapes) - SIMPLE
# ============================================================================

def create_slide_8_chapter2_divider(prs):
    """Slide 8: Chapter 2 Divider - Simple chapter break (4-5 shapes)

    Layout: Minimalist chapter divider
    - Large chapter number "2장"
    - Chapter title "Kraljic Matrix 프레임워크"
    - Simple monochrome design
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    shape_count = 0

    # Background rectangle (optional for visual depth)
    add_rectangle(
        slide, 0.50, 2.50, 9.80, 2.50,
        fill_color=COLOR_VERY_LIGHT_GRAY,
        border_color=None,
        border_width=0
    )
    shape_count += 1

    # Chapter number "2장"
    add_text_box(
        slide, 2.00, 2.80, 6.80, 0.80,
        "2장", font_size=44, bold=True,
        color=COLOR_DARK_GRAY, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Chapter title
    add_text_box(
        slide, 2.00, 3.80, 6.80, 0.70,
        "Kraljic Matrix 프레임워크", font_size=24, bold=True,
        color=COLOR_BLACK, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Decorative line
    add_rectangle(
        slide, 3.50, 4.70, 3.80, 0.05,
        fill_color=COLOR_DARK_GRAY,
        border_color=None
    )
    shape_count += 1

    print(f"✓ Slide 8: Chapter 2 Divider ({shape_count} shapes)")
    return slide

# ============================================================================
# SLIDE 9: KRALJIC MATRIX BIRTH (70-80 shapes) - TOY PAGE!
# ============================================================================

def create_slide_9_kraljic_birth(prs):
    """Slide 9: Kraljic Matrix Birth - Toy Page layout (70-80 shapes)

    Layout: Toy Page (65% visual + 30% text)
    - Left: Timeline of Kraljic development (1983-present)
    - Right: Key insights and significance
    - Use 8-9pt text extensively
    - High visual impact with arrows and progression
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "2.1 Kraljic Matrix 탄생", slide_num=9)
    add_governing_message(
        slide,
        "1983년 Peter Kraljic이 개발한 2×2 매트릭스는 자재 특성에 따른 차별화 전략의 기초가 되었습니다."
    )

    shape_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR

    # ===== LEFT SIDE (65%): Visual Timeline =====
    left_x = 0.80
    left_w = 6.50

    # Timeline title
    add_rectangle(
        slide, left_x, 2.00, left_w, 0.35,
        fill_color=COLOR_MED_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, left_x + 0.10, 2.05, left_w - 0.20, 0.25,
        "Kraljic Matrix 발전 타임라인 (1983-현재)", font_size=11, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Timeline events (vertical progression)
    events = [
        {
            "year": "1983", "title": "HBR 논문 발표",
            "details": ["Peter Kraljic", "Purchasing Must Become", "Supply Management"],
            "impact": "2×2 매트릭스 최초 제안"
        },
        {
            "year": "1985-90", "title": "학계 확산",
            "details": ["이론 정립", "실증 연구", "교육 과정 포함"],
            "impact": "MBA 필수 교육"
        },
        {
            "year": "1990-2000", "title": "산업 적용",
            "details": ["Fortune 500 채택", "자동차·전자 확산", "컨설팅 방법론화"],
            "impact": "글로벌 표준 확립"
        },
        {
            "year": "2000-2010", "title": "디지털화",
            "details": ["ERP 통합", "자동 분류", "데이터 기반"],
            "impact": "시스템 자동화"
        },
        {
            "year": "2010-2020", "title": "고도화",
            "details": ["AI/ML 접목", "동적 분류", "실시간 모니터링"],
            "impact": "지능형 SCM"
        },
        {
            "year": "2020-현재", "title": "리스크 관리",
            "details": ["공급망 탄력성", "리스크 지표 강화", "시나리오 분석"],
            "impact": "필수 프레임워크"
        }
    ]

    event_y = 2.50
    event_h = 0.70
    for idx, event in enumerate(events):
        # Year marker (left)
        add_rectangle(
            slide, left_x, event_y, 0.90, 0.35,
            fill_color=COLOR_DARK_GRAY,
            border_color=COLOR_BLACK,
            border_width=1
        )
        shape_count += 1

        add_text_box(
            slide, left_x + 0.05, event_y + 0.06, 0.80, 0.24,
            event["year"], font_size=9, bold=True,
            color=COLOR_WHITE, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Title box
        add_rectangle(
            slide, left_x + 1.00, event_y, 2.20, 0.35,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, left_x + 1.10, event_y + 0.06, 2.00, 0.24,
            event["title"], font_size=10, bold=True,
            color=COLOR_BLACK
        )
        shape_count += 1

        # Details (3 items, 8pt)
        detail_x = left_x + 1.00
        detail_y = event_y + 0.40
        for detail in event["details"]:
            add_text_box(
                slide, detail_x + 0.08, detail_y, 0.10, 0.12,
                "•", font_size=7, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            add_text_box(
                slide, detail_x + 0.20, detail_y, 2.00, 0.12,
                detail, font_size=8, color=COLOR_DARK_GRAY
            )
            shape_count += 1

            detail_y += 0.13

        # Impact box (right)
        add_rectangle(
            slide, left_x + 3.30, event_y, 3.00, event_h,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        add_text_box(
            slide, left_x + 3.40, event_y + (event_h - 0.24)/2, 2.80, 0.24,
            event["impact"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connector arrow (if not last)
        if idx < len(events) - 1:
            arrow = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT,
                Inches(left_x + 0.45), Inches(event_y + event_h),
                Inches(left_x + 0.45), Inches(event_y + event_h + 0.08)
            )
            arrow.line.color.rgb = COLOR_MED_GRAY
            arrow.line.width = Pt(2)
            arrow.line.end_arrow_type = 2
            shape_count += 1

        event_y += event_h + 0.08

    # ===== RIGHT SIDE (30%): Text Insights =====
    right_x = 7.50
    right_w = 2.80

    # Section 1: 시사점 (Insights)
    add_rectangle(
        slide, right_x, 2.00, right_w, 0.30,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, 2.04, right_w - 0.20, 0.22,
        "시사점", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    insights = [
        "40년간 검증된 프레임워크",
        "학계와 산업계 공동 인정",
        "시대 변화에 따라 진화",
        "현재까지 가장 널리 사용",
        "디지털 시대에도 유효성 입증"
    ]

    insight_y = 2.40
    for insight in insights:
        add_text_box(
            slide, right_x + 0.08, insight_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, insight_y, right_w - 0.32, 0.16,
            insight, font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        insight_y += 0.20

    # Section 2: 핵심 개념 (Key Concepts)
    concept_y = insight_y + 0.15
    add_rectangle(
        slide, right_x, concept_y, right_w, 0.30,
        fill_color=COLOR_MED_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, concept_y + 0.04, right_w - 0.20, 0.22,
        "핵심 개념", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    concepts = [
        "차별화: 획일적 관리 탈피",
        "2차원 분석: 리스크 × 임팩트",
        "4사분면: 전략·레버리지·병목·일상",
        "맞춤 전략: 자재군별 최적화",
        "동적 관리: 지속적 재분류"
    ]

    concept_text_y = concept_y + 0.40
    for concept in concepts:
        add_text_box(
            slide, right_x + 0.08, concept_text_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_BLACK
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, concept_text_y, right_w - 0.32, 0.16,
            concept, font_size=8, color=COLOR_BLACK
        )
        shape_count += 1

        concept_text_y += 0.20

    # Section 3: 적용 현황 (Current Status)
    status_y = concept_text_y + 0.15
    add_rectangle(
        slide, right_x, status_y, right_w, 0.30,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, status_y + 0.04, right_w - 0.20, 0.22,
        "적용 현황", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    statuses = [
        "Fortune 500: 95% 사용",
        "제조업: 필수 방법론",
        "공공 조달: 정부 채택",
        "교육: MBA 핵심 과목",
        "인증: APICS/ISM 포함"
    ]

    status_text_y = status_y + 0.40
    for status in statuses:
        add_text_box(
            slide, right_x + 0.08, status_text_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, status_text_y, right_w - 0.32, 0.16,
            status, font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        status_text_y += 0.20

    print(f"✓ Slide 9: Kraljic Birth ({shape_count} shapes) - TOY PAGE!")
    return slide

# ============================================================================
# SLIDE 10: KRALJIC AXES (75-85 shapes) - TOY PAGE!
# ============================================================================

def create_slide_10_kraljic_axes(prs):
    """Slide 10: Kraljic Matrix Axes - Toy Page layout (75-85 shapes)

    Layout: Toy Page (65% visual + 30% text)
    - Left: Visual representation of the two axes with detailed indicators
    - Right: Evaluation criteria and measurement methods
    - Use 8-9pt text extensively
    - High visual impact with axis diagrams
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "2.2 2×2 매트릭스의 두 축", slide_num=10)
    add_governing_message(
        slide,
        "공급 리스크(X축)와 구매 임팩트(Y축) 두 기준으로 자재를 4개 군으로 분류합니다."
    )

    shape_count = 0
    from pptx.enum.shapes import MSO_CONNECTOR

    # ===== LEFT SIDE (65%): Axis Visualization =====
    left_x = 0.80
    left_w = 6.50

    # Y-AXIS: Purchase Impact (Vertical)
    y_axis_x = left_x + 0.30
    y_axis_y_start = 2.20
    y_axis_y_end = 6.60

    # Y-axis line
    y_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(y_axis_x), Inches(y_axis_y_start),
        Inches(y_axis_x), Inches(y_axis_y_end)
    )
    y_line.line.color.rgb = COLOR_DARK_GRAY
    y_line.line.width = Pt(3)
    y_line.line.begin_arrow_type = 2  # Arrow at start (top)
    shape_count += 1

    # Y-axis label
    add_text_box(
        slide, y_axis_x - 0.50, y_axis_y_start - 0.25, 1.00, 0.20,
        "구매 임팩트 (Y축)", font_size=11, bold=True,
        color=COLOR_BLACK, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Y-axis indicators (5 levels with HIGH density)
    y_indicators = [
        {"level": "매우 높음", "desc": "연매출 10% 이상", "value": "100억원+", "y": 2.50},
        {"level": "높음", "desc": "연매출 5-10%", "value": "50-100억원", "y": 3.40},
        {"level": "중간", "desc": "연매출 2-5%", "value": "10-50억원", "y": 4.30},
        {"level": "낮음", "desc": "연매출 1-2%", "value": "5-10억원", "y": 5.20},
        {"level": "매우 낮음", "desc": "연매출 1% 미만", "value": "5억원 이하", "y": 6.10}
    ]

    for indicator in y_indicators:
        # Level box
        add_rectangle(
            slide, y_axis_x + 0.15, indicator["y"], 1.20, 0.70,
            fill_color=COLOR_VERY_LIGHT_GRAY,
            border_color=COLOR_MED_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # Level label (9pt bold)
        add_text_box(
            slide, y_axis_x + 0.20, indicator["y"] + 0.05, 1.10, 0.16,
            indicator["level"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Description (8pt)
        add_text_box(
            slide, y_axis_x + 0.20, indicator["y"] + 0.24, 1.10, 0.14,
            indicator["desc"], font_size=8, color=COLOR_DARK_GRAY,
            align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value (8pt bold)
        add_text_box(
            slide, y_axis_x + 0.20, indicator["y"] + 0.42, 1.10, 0.22,
            indicator["value"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connector to axis
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(y_axis_x), Inches(indicator["y"] + 0.35),
            Inches(y_axis_x + 0.15), Inches(indicator["y"] + 0.35)
        )
        conn.line.color.rgb = COLOR_MED_GRAY
        conn.line.width = Pt(1)
        shape_count += 1

    # X-AXIS: Supply Risk (Horizontal)
    x_axis_x_start = left_x + 2.00
    x_axis_x_end = left_x + 6.50
    x_axis_y = 6.70

    # X-axis line
    x_line = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x_axis_x_start), Inches(x_axis_y),
        Inches(x_axis_x_end), Inches(x_axis_y)
    )
    x_line.line.color.rgb = COLOR_DARK_GRAY
    x_line.line.width = Pt(3)
    x_line.line.end_arrow_type = 2  # Arrow at end (right)
    shape_count += 1

    # X-axis label
    add_text_box(
        slide, x_axis_x_end - 0.40, x_axis_y + 0.15, 0.80, 0.20,
        "공급 리스크 (X축)", font_size=11, bold=True,
        color=COLOR_BLACK, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # X-axis indicators (5 levels with HIGH density)
    x_indicators = [
        {"level": "매우 낮음", "desc": "공급원 10개+", "value": "선택 다양", "x": 2.10},
        {"level": "낮음", "desc": "공급원 5-10개", "value": "대체 용이", "x": 3.00},
        {"level": "중간", "desc": "공급원 3-5개", "value": "대체 가능", "x": 3.90},
        {"level": "높음", "desc": "공급원 1-2개", "value": "대체 어려움", "x": 4.80},
        {"level": "매우 높음", "desc": "공급원 1개", "value": "대체 불가", "x": 5.70}
    ]

    for indicator in x_indicators:
        # Level box
        add_rectangle(
            slide, indicator["x"], x_axis_y - 1.05, 0.80, 0.90,
            fill_color=COLOR_WHITE,
            border_color=COLOR_LIGHT_GRAY,
            border_width=0.75
        )
        shape_count += 1

        # Level label (9pt bold)
        add_text_box(
            slide, indicator["x"] + 0.05, x_axis_y - 1.00, 0.70, 0.16,
            indicator["level"], font_size=9, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Description (8pt)
        add_text_box(
            slide, indicator["x"] + 0.05, x_axis_y - 0.78, 0.70, 0.24,
            indicator["desc"], font_size=7, color=COLOR_DARK_GRAY,
            align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value (8pt bold)
        add_text_box(
            slide, indicator["x"] + 0.05, x_axis_y - 0.48, 0.70, 0.18,
            indicator["value"], font_size=8, bold=True,
            color=COLOR_BLACK, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Connector to axis
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(indicator["x"] + 0.40), Inches(x_axis_y - 0.15),
            Inches(indicator["x"] + 0.40), Inches(x_axis_y)
        )
        conn.line.color.rgb = COLOR_MED_GRAY
        conn.line.width = Pt(1)
        shape_count += 1

    # ===== RIGHT SIDE (30%): Evaluation Criteria =====
    right_x = 7.50
    right_w = 2.80

    # Section 1: 구매 임팩트 평가
    add_rectangle(
        slide, right_x, 2.00, right_w, 0.30,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, 2.04, right_w - 0.20, 0.22,
        "구매 임팩트 평가", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    impact_criteria = [
        "연간 구매 금액",
        "매출 대비 비중",
        "수익성 영향도",
        "전략적 중요도",
        "대체 비용",
        "품질 영향도"
    ]

    criteria_y = 2.40
    for criterion in impact_criteria:
        add_text_box(
            slide, right_x + 0.08, criteria_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, criteria_y, right_w - 0.32, 0.16,
            criterion, font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        criteria_y += 0.19

    # Section 2: 공급 리스크 평가
    risk_y = criteria_y + 0.12
    add_rectangle(
        slide, right_x, risk_y, right_w, 0.30,
        fill_color=COLOR_MED_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, risk_y + 0.04, right_w - 0.20, 0.22,
        "공급 리스크 평가", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    risk_criteria = [
        "공급업체 수",
        "대체 가능성",
        "납기 리드타임",
        "품질 안정성",
        "지역 집중도",
        "기술 의존도"
    ]

    risk_criteria_y = risk_y + 0.40
    for criterion in risk_criteria:
        add_text_box(
            slide, right_x + 0.08, risk_criteria_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_BLACK
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, risk_criteria_y, right_w - 0.32, 0.16,
            criterion, font_size=8, color=COLOR_BLACK
        )
        shape_count += 1

        risk_criteria_y += 0.19

    # Section 3: 측정 방법
    method_y = risk_criteria_y + 0.12
    add_rectangle(
        slide, right_x, method_y, right_w, 0.30,
        fill_color=COLOR_DARK_GRAY,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, right_x + 0.10, method_y + 0.04, right_w - 0.20, 0.22,
        "측정 방법", font_size=10, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    methods = [
        "정량 데이터: ERP 추출",
        "정성 평가: 전문가 점수",
        "가중치 적용: 중요도 반영",
        "스코어링: 0-100점",
        "매트릭스 매핑: 자동 분류",
        "주기적 재평가: 분기/반기"
    ]

    method_text_y = method_y + 0.40
    for method in methods:
        add_text_box(
            slide, right_x + 0.08, method_text_y, 0.12, 0.16,
            "•", font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        add_text_box(
            slide, right_x + 0.22, method_text_y, right_w - 0.32, 0.16,
            method, font_size=8, color=COLOR_DARK_GRAY
        )
        shape_count += 1

        method_text_y += 0.19

    print(f"✓ Slide 10: Kraljic Axes ({shape_count} shapes) - TOY PAGE!")
    return slide

# ============================================================================
# SLIDE 11: KRALJIC MATRIX DOOR CHART (100-120 shapes) - CRITICAL!!!
# ============================================================================

def create_slide_11_kraljic_door_chart(prs):
    """Slide 11: Kraljic Matrix Door Chart - THE CRITICAL SLIDE! (100-120 shapes)

    Layout: Door Chart pattern with maximum density
    - 2×2 Matrix with 4 colored quadrants
    - Each quadrant: 15-20 detail items (8pt text)
    - Axis labels and spectrum indicators
    - Strategic recommendations for each quadrant
    - Use 70-80% of shapes in 9pt or smaller
    - This is THE most important slide - maximum information density!
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_slide_title(slide, "2.3 Kraljic Matrix", slide_num=11)
    add_governing_message(
        slide,
        "Kraljic Matrix는 공급 리스크와 구매 금액을 기준으로 자재를 4개 군으로 분류하여 차별화 전략을 수립합니다."
    )

    shape_count = 0

    # ===== MATRIX DIMENSIONS =====
    matrix_x = 1.50
    matrix_y = 2.20
    quad_w = 3.80
    quad_h = 2.00
    gap = 0.10

    # Define Kraljic colors (ONLY used in this slide!)
    COLOR_STRATEGIC = RGBColor(142, 68, 173)   # Purple
    COLOR_BOTTLENECK = RGBColor(230, 126, 34)  # Orange
    COLOR_LEVERAGE = RGBColor(39, 174, 96)     # Green
    COLOR_ROUTINE = RGBColor(149, 165, 166)    # Gray

    # ===== AXIS LABELS =====
    # Y-axis label (left)
    add_text_box(
        slide, 0.50, matrix_y + quad_h, 0.80, 2*quad_h + gap,
        "구매 임팩트\n(Purchase Impact)\n→", font_size=12, bold=True,
        color=COLOR_BLACK, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # X-axis label (bottom)
    add_text_box(
        slide, matrix_x, matrix_y + 2*quad_h + 2*gap + 0.10, 2*quad_w + gap, 0.40,
        "공급 리스크 (Supply Risk) →", font_size=12, bold=True,
        color=COLOR_BLACK, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # ===== QUADRANT 1: STRATEGIC (TOP RIGHT) =====
    q1_x = matrix_x + quad_w + gap
    q1_y = matrix_y

    add_rectangle(
        slide, q1_x, q1_y, quad_w, quad_h,
        fill_color=COLOR_STRATEGIC,
        border_color=COLOR_BLACK,
        border_width=2
    )
    shape_count += 1

    # Quadrant title
    add_text_box(
        slide, q1_x + 0.10, q1_y + 0.08, quad_w - 0.20, 0.28,
        "전략자재 (Strategic)", font_size=13, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Strategic details (15-18 items, 8pt)
    strategic_items = [
        "특성: 고리스크 + 고임팩트",
        "공급원: 소수 (1-3개)",
        "구매금액: 매출 10% 이상",
        "사례: 반도체, 핵심 원자재",
        "--- 전략 ---",
        "장기 파트너십 구축",
        "협력적 관계 강화",
        "공동 기술 개발",
        "리스크 공유 계약",
        "--- 계획 ---",
        "하이브리드: LTP + MRP",
        "예측 + 수요 결합",
        "전략적 안전재고",
        "--- KPI ---",
        "공급 안정성 95%+",
        "품질 불량률 0.1% 이하",
        "납기 준수율 98%+",
        "협력 지수 4.5/5.0"
    ]

    detail_y = q1_y + 0.42
    for item in strategic_items:
        if item.startswith("---"):  # Section divider
            add_text_box(
                slide, q1_x + 0.12, detail_y, quad_w - 0.24, 0.12,
                item, font_size=8, bold=True,
                color=COLOR_WHITE, align=PP_ALIGN.CENTER
            )
            shape_count += 1
            detail_y += 0.14
        else:
            add_text_box(
                slide, q1_x + 0.12, detail_y, 0.10, 0.10,
                "•", font_size=7, color=COLOR_WHITE
            )
            shape_count += 1

            add_text_box(
                slide, q1_x + 0.24, detail_y, quad_w - 0.36, 0.10,
                item, font_size=8, color=COLOR_WHITE
            )
            shape_count += 1

            detail_y += 0.11

    # ===== QUADRANT 2: BOTTLENECK (TOP LEFT) =====
    q2_x = matrix_x
    q2_y = matrix_y

    add_rectangle(
        slide, q2_x, q2_y, quad_w, quad_h,
        fill_color=COLOR_BOTTLENECK,
        border_color=COLOR_BLACK,
        border_width=2
    )
    shape_count += 1

    # Quadrant title
    add_text_box(
        slide, q2_x + 0.10, q2_y + 0.08, quad_w - 0.20, 0.28,
        "병목자재 (Bottleneck)", font_size=13, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Bottleneck details
    bottleneck_items = [
        "특성: 고리스크 + 저임팩트",
        "공급원: 매우 소수 (1-2개)",
        "구매금액: 매출 2% 미만",
        "사례: 특수 부품, 인증 자재",
        "--- 전략 ---",
        "공급 안정성 최우선",
        "안전재고 충분 확보",
        "대체품 개발 추진",
        "복수 공급원 발굴",
        "--- 계획 ---",
        "ROP (재주문점) 방식",
        "Min-Max 재고 관리",
        "높은 안전재고율",
        "--- KPI ---",
        "재고 가용률 98%+",
        "결품률 0.5% 이하",
        "리드타임 준수 95%+",
        "비상 재고 8주+"
    ]

    detail_y = q2_y + 0.42
    for item in bottleneck_items:
        if item.startswith("---"):
            add_text_box(
                slide, q2_x + 0.12, detail_y, quad_w - 0.24, 0.12,
                item, font_size=8, bold=True,
                color=COLOR_WHITE, align=PP_ALIGN.CENTER
            )
            shape_count += 1
            detail_y += 0.14
        else:
            add_text_box(
                slide, q2_x + 0.12, detail_y, 0.10, 0.10,
                "•", font_size=7, color=COLOR_WHITE
            )
            shape_count += 1

            add_text_box(
                slide, q2_x + 0.24, detail_y, quad_w - 0.36, 0.10,
                item, font_size=8, color=COLOR_WHITE
            )
            shape_count += 1

            detail_y += 0.11

    # ===== QUADRANT 3: LEVERAGE (BOTTOM RIGHT) =====
    q3_x = matrix_x + quad_w + gap
    q3_y = matrix_y + quad_h + gap

    add_rectangle(
        slide, q3_x, q3_y, quad_w, quad_h,
        fill_color=COLOR_LEVERAGE,
        border_color=COLOR_BLACK,
        border_width=2
    )
    shape_count += 1

    # Quadrant title
    add_text_box(
        slide, q3_x + 0.10, q3_y + 0.08, quad_w - 0.20, 0.28,
        "레버리지자재 (Leverage)", font_size=13, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Leverage details
    leverage_items = [
        "특성: 저리스크 + 고임팩트",
        "공급원: 다수 (10개+)",
        "구매금액: 매출 5-10%",
        "사례: 표준 부품, 원자재",
        "--- 전략 ---",
        "경쟁 입찰 활용",
        "가격 협상 중점",
        "물량 레버리지 활용",
        "단기 계약 체결",
        "--- 계획 ---",
        "MRP (자재소요계획)",
        "수요 기반 발주",
        "최소 안전재고",
        "--- KPI ---",
        "원가 절감률 5%+",
        "재고 회전율 12회+",
        "가격 경쟁력 상위 10%",
        "조달 효율 90%+"
    ]

    detail_y = q3_y + 0.42
    for item in leverage_items:
        if item.startswith("---"):
            add_text_box(
                slide, q3_x + 0.12, detail_y, quad_w - 0.24, 0.12,
                item, font_size=8, bold=True,
                color=COLOR_WHITE, align=PP_ALIGN.CENTER
            )
            shape_count += 1
            detail_y += 0.14
        else:
            add_text_box(
                slide, q3_x + 0.12, detail_y, 0.10, 0.10,
                "•", font_size=7, color=COLOR_WHITE
            )
            shape_count += 1

            add_text_box(
                slide, q3_x + 0.24, detail_y, quad_w - 0.36, 0.10,
                item, font_size=8, color=COLOR_WHITE
            )
            shape_count += 1

            detail_y += 0.11

    # ===== QUADRANT 4: ROUTINE (BOTTOM LEFT) =====
    q4_x = matrix_x
    q4_y = matrix_y + quad_h + gap

    add_rectangle(
        slide, q4_x, q4_y, quad_w, quad_h,
        fill_color=COLOR_ROUTINE,
        border_color=COLOR_BLACK,
        border_width=2
    )
    shape_count += 1

    # Quadrant title
    add_text_box(
        slide, q4_x + 0.10, q4_y + 0.08, quad_w - 0.20, 0.28,
        "일상자재 (Routine)", font_size=13, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Routine details
    routine_items = [
        "특성: 저리스크 + 저임팩트",
        "공급원: 매우 다수 (20개+)",
        "구매금액: 매출 1% 미만",
        "사례: 소모품, MRO",
        "--- 전략 ---",
        "프로세스 효율화",
        "자동 발주 시스템",
        "통합 구매 (카탈로그)",
        "관리 비용 최소화",
        "--- 계획 ---",
        "Min-Max 자동 발주",
        "VMI (공급자 관리 재고)",
        "E-Procurement 활용",
        "--- KPI ---",
        "처리 시간 단축 50%+",
        "발주 비용 최소화",
        "자동화율 80%+",
        "사용자 만족도 4.0/5.0"
    ]

    detail_y = q4_y + 0.42
    for item in routine_items:
        if item.startswith("---"):
            add_text_box(
                slide, q4_x + 0.12, detail_y, quad_w - 0.24, 0.12,
                item, font_size=8, bold=True,
                color=COLOR_WHITE, align=PP_ALIGN.CENTER
            )
            shape_count += 1
            detail_y += 0.14
        else:
            add_text_box(
                slide, q4_x + 0.12, detail_y, 0.10, 0.10,
                "•", font_size=7, color=COLOR_WHITE
            )
            shape_count += 1

            add_text_box(
                slide, q4_x + 0.24, detail_y, quad_w - 0.36, 0.10,
                item, font_size=8, color=COLOR_WHITE
            )
            shape_count += 1

            detail_y += 0.11

    # ===== SUMMARY TABLE (Right side) =====
    summary_x = 9.50
    summary_y = 2.20
    summary_w = 0.75
    summary_h = 4.20

    # Summary table header
    add_rectangle(
        slide, summary_x - 0.10, summary_y - 0.05, summary_w + 0.20, 0.30,
        fill_color=COLOR_BLACK,
        border_color=COLOR_BLACK,
        border_width=1
    )
    shape_count += 1

    add_text_box(
        slide, summary_x - 0.05, summary_y - 0.01, summary_w + 0.10, 0.22,
        "비중", font_size=9, bold=True,
        color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )
    shape_count += 1

    # Quadrant percentages
    percentages = [
        {"label": "전략", "value": "15-20%", "y": summary_y + 0.35, "color": COLOR_STRATEGIC},
        {"label": "병목", "value": "5-10%", "y": summary_y + 1.05, "color": COLOR_BOTTLENECK},
        {"label": "레버", "value": "50-60%", "y": summary_y + 2.45, "color": COLOR_LEVERAGE},
        {"label": "일상", "value": "20-25%", "y": summary_y + 3.15, "color": COLOR_ROUTINE}
    ]

    for pct in percentages:
        # Percentage box
        add_rectangle(
            slide, summary_x - 0.10, pct["y"], summary_w + 0.20, 0.55,
            fill_color=pct["color"],
            border_color=COLOR_BLACK,
            border_width=1
        )
        shape_count += 1

        # Label
        add_text_box(
            slide, summary_x - 0.05, pct["y"] + 0.05, summary_w + 0.10, 0.18,
            pct["label"], font_size=9, bold=True,
            color=COLOR_WHITE, align=PP_ALIGN.CENTER
        )
        shape_count += 1

        # Value
        add_text_box(
            slide, summary_x - 0.05, pct["y"] + 0.28, summary_w + 0.10, 0.22,
            pct["value"], font_size=8, bold=True,
            color=COLOR_WHITE, align=PP_ALIGN.CENTER
        )
        shape_count += 1

    print(f"✓ Slide 11: Kraljic Door Chart ({shape_count} shapes) - CRITICAL DOOR CHART!")
    return slide

# ============================================================================
# MAIN GENERATION FUNCTION
# ============================================================================

def main():
    """Generate Part 1 PPTX (Slides 1-11) - INCLUDING CRITICAL DOOR CHART!"""
    print("=== Part 1 PPTX Generation (Slides 1-11) ===")
    print("High-quality implementation following S4HANA standards")
    print("Including CRITICAL Kraljic Matrix Door Chart (Slide 11)\n")

    prs = create_presentation()

    # Chapter 1: JIT → JIC Paradigm Shift
    create_slide_1_cover(prs)
    create_slide_2_toc(prs)
    create_slide_3_chapter1_divider(prs)
    create_slide_4_jit_timeline(prs)
    create_slide_5_pandemic(prs)
    create_slide_6_jit_vs_jic(prs)
    create_slide_7_jic_adopters(prs)

    # Chapter 2: Kraljic Matrix Framework
    create_slide_8_chapter2_divider(prs)
    create_slide_9_kraljic_birth(prs)
    create_slide_10_kraljic_axes(prs)
    create_slide_11_kraljic_door_chart(prs)

    # Save
    output_path = "/home/user/Kraljic_Course/Part1_Session1_StrategicInventory.pptx"
    prs.save(output_path)

    print(f"\n=== Generation Complete ===")
    print(f"Output: {output_path}")
    print(f"Slides generated: 11 (including Door Chart!)")
    print(f"\nNext: Run verification script")

    return output_path

if __name__ == "__main__":
    main()
