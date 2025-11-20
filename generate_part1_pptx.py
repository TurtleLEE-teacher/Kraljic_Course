#!/usr/bin/env python3
"""
Generate Part 1 PPTX using python-pptx
S4HANA design system: monochrome colors, 10.83" × 7.5", governing messages
"""

import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# S4HANA 디자인 시스템
SLIDE_WIDTH = Inches(10.83)
SLIDE_HEIGHT = Inches(7.5)

# 모노크롬 색상 팔레트
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_DARK_GRAY = RGBColor(51, 51, 51)
COLOR_MED_GRAY = RGBColor(102, 102, 102)
COLOR_LIGHT_GRAY = RGBColor(204, 204, 204)
COLOR_VERY_LIGHT_GRAY = RGBColor(230, 230, 230)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_ACCENT = RGBColor(26, 82, 118)  # Dark blue for minimal use

# Kraljic Matrix colors (ONLY for matrix diagram)
COLOR_STRATEGIC = RGBColor(142, 68, 173)  # Purple
COLOR_BOTTLENECK = RGBColor(230, 126, 34)  # Orange
COLOR_LEVERAGE = RGBColor(39, 174, 96)  # Green
COLOR_ROUTINE = RGBColor(149, 165, 166)  # Gray

# Font sizes (based on S4HANA analysis)
FONT_SIZE_COVER_TITLE = Pt(48)
FONT_SIZE_TITLE = Pt(20)
FONT_SIZE_GOVERNING = Pt(16)  # NOT 14pt Italic - use 16pt Bold
FONT_SIZE_HEADING = Pt(14)
FONT_SIZE_BODY = Pt(10)  # PRIMARY font size (65% of all text)
FONT_SIZE_BULLET = Pt(12)  # For bullet points (20-25% of text)
FONT_SIZE_CAPTION = Pt(8)


def create_presentation():
    """Create presentation with S4HANA dimensions"""
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs


def add_cover_slide(prs, data):
    """Add cover slide"""
    blank_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(1.0), Inches(2.5), Inches(8.83), Inches(1.0)
    )
    title_frame = title_box.text_frame
    title_frame.text = data['title']
    title_frame.paragraphs[0].font.size = FONT_SIZE_COVER_TITLE
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = COLOR_BLACK
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(
        Inches(1.0), Inches(3.7), Inches(8.83), Inches(0.6)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = data['subtitle']
    subtitle_frame.paragraphs[0].font.size = Pt(28)
    subtitle_frame.paragraphs[0].font.color.rgb = COLOR_MED_GRAY
    subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Course info
    course_box = slide.shapes.add_textbox(
        Inches(1.0), Inches(5.0), Inches(8.83), Inches(1.0)
    )
    course_frame = course_box.text_frame
    course_frame.text = f"{data['course']}\n{data['date']}\n{data['instructor']}"
    for para in course_frame.paragraphs:
        para.font.size = Pt(14)
        para.font.color.rgb = COLOR_DARK_GRAY
        para.alignment = PP_ALIGN.CENTER

    return slide


def add_title_and_governing(slide, title_text, governing_text=None, slide_number=None):
    """Add title and governing message (REQUIRED for all content slides)"""
    # Title
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3), Inches(9.83), Inches(0.5)
    )
    title_frame = title_box.text_frame
    title_frame.text = title_text
    title_frame.paragraphs[0].font.size = FONT_SIZE_TITLE
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = COLOR_DARK_GRAY

    # Governing message (16pt Bold, NOT 14pt Italic!)
    if governing_text:
        gov_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.9), Inches(9.83), Inches(0.6)
        )
        gov_frame = gov_box.text_frame
        gov_frame.text = governing_text
        gov_frame.paragraphs[0].font.size = FONT_SIZE_GOVERNING
        gov_frame.paragraphs[0].font.bold = True  # BOLD not italic
        gov_frame.paragraphs[0].font.color.rgb = COLOR_MED_GRAY

    # Slide number
    if slide_number:
        num_box = slide.shapes.add_textbox(
            Inches(9.8), Inches(7.0), Inches(0.8), Inches(0.3)
        )
        num_frame = num_box.text_frame
        num_frame.text = str(slide_number)
        num_frame.paragraphs[0].font.size = FONT_SIZE_CAPTION
        num_frame.paragraphs[0].font.color.rgb = COLOR_MED_GRAY
        num_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT


def add_list_bullets_slide(prs, data):
    """Add bullet list slide"""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_title_and_governing(
        slide,
        data.get('title', ''),
        data.get('governingMessage'),
        data.get('slideNumber')
    )

    # Introduction (optional)
    y_position = Inches(1.6)
    if data.get('introduction'):
        intro_box = slide.shapes.add_textbox(
            Inches(0.8), y_position, Inches(9.0), Inches(0.4)
        )
        intro_frame = intro_box.text_frame
        intro_frame.text = data['introduction']
        intro_frame.paragraphs[0].font.size = FONT_SIZE_BODY
        intro_frame.paragraphs[0].font.color.rgb = COLOR_DARK_GRAY
        y_position += Inches(0.6)

    # Bullet points
    items = data.get('items', [])
    bullet_box = slide.shapes.add_textbox(
        Inches(0.8), y_position, Inches(9.0), Inches(5.0)
    )
    text_frame = bullet_box.text_frame
    text_frame.word_wrap = True

    for i, item in enumerate(items):
        p = text_frame.paragraphs[i] if i == 0 else text_frame.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = FONT_SIZE_BULLET
        p.font.color.rgb = COLOR_DARK_GRAY
        p.space_before = Pt(6) if i > 0 else Pt(0)

    return slide


def add_content_2col_slide(prs, data):
    """Add 2-column comparison slide"""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_title_and_governing(
        slide,
        data.get('title', ''),
        data.get('governingMessage'),
        data.get('slideNumber')
    )

    # Left column
    left_box = slide.shapes.add_textbox(
        Inches(0.8), Inches(1.6), Inches(4.5), Inches(5.0)
    )
    left_frame = left_box.text_frame
    left_frame.word_wrap = True

    # Left title
    left_title = left_frame.paragraphs[0]
    left_title.text = data.get('leftTitle', '')
    left_title.font.size = FONT_SIZE_HEADING
    left_title.font.bold = True
    left_title.font.color.rgb = COLOR_DARK_GRAY

    # Left content
    left_content = left_frame.add_paragraph()
    left_content.text = data.get('leftContent', '')
    left_content.font.size = FONT_SIZE_BODY
    left_content.font.color.rgb = COLOR_DARK_GRAY
    left_content.space_before = Pt(12)

    # Right column
    right_box = slide.shapes.add_textbox(
        Inches(5.5), Inches(1.6), Inches(4.5), Inches(5.0)
    )
    right_frame = right_box.text_frame
    right_frame.word_wrap = True

    # Right title
    right_title = right_frame.paragraphs[0]
    right_title.text = data.get('rightTitle', '')
    right_title.font.size = FONT_SIZE_HEADING
    right_title.font.bold = True
    right_title.font.color.rgb = COLOR_DARK_GRAY

    # Right content
    right_content = right_frame.add_paragraph()
    right_content.text = data.get('rightContent', '')
    right_content.font.size = FONT_SIZE_BODY
    right_content.font.color.rgb = COLOR_DARK_GRAY
    right_content.space_before = Pt(12)

    return slide


def add_kraljic_matrix_slide(prs, data):
    """Add Kraljic Matrix 2x2 diagram (ONLY place to use colors)"""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_title_and_governing(
        slide,
        data.get('title', ''),
        data.get('governingMessage'),
        data.get('slideNumber')
    )

    # Introduction
    if data.get('introduction'):
        intro_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.6), Inches(9.0), Inches(0.3)
        )
        intro_frame = intro_box.text_frame
        intro_frame.text = data['introduction']
        intro_frame.paragraphs[0].font.size = FONT_SIZE_BODY
        intro_frame.paragraphs[0].font.color.rgb = COLOR_DARK_GRAY

    # Matrix dimensions
    matrix_left = Inches(1.5)
    matrix_top = Inches(2.2)
    quadrant_width = Inches(4.0)
    quadrant_height = Inches(2.2)

    # Strategic (Top-Right) - Purple
    strategic_shape = slide.shapes.add_shape(
        1,  # Rectangle
        matrix_left + quadrant_width, matrix_top,
        quadrant_width, quadrant_height
    )
    strategic_shape.fill.solid()
    strategic_shape.fill.fore_color.rgb = COLOR_STRATEGIC
    strategic_shape.line.color.rgb = COLOR_WHITE
    strategic_shape.line.width = Pt(2)

    strategic_text = strategic_shape.text_frame
    strategic_text.text = "전략자재\nStrategic"
    for para in strategic_text.paragraphs:
        para.font.size = FONT_SIZE_HEADING
        para.font.bold = True
        para.font.color.rgb = COLOR_WHITE
        para.alignment = PP_ALIGN.CENTER

    # Add strategic items
    if data.get('strategic'):
        for item in data['strategic']:
            p = strategic_text.add_paragraph()
            p.text = item
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_WHITE

    # Bottleneck (Top-Left) - Orange
    bottleneck_shape = slide.shapes.add_shape(
        1,  # Rectangle
        matrix_left, matrix_top,
        quadrant_width, quadrant_height
    )
    bottleneck_shape.fill.solid()
    bottleneck_shape.fill.fore_color.rgb = COLOR_BOTTLENECK
    bottleneck_shape.line.color.rgb = COLOR_WHITE
    bottleneck_shape.line.width = Pt(2)

    bottleneck_text = bottleneck_shape.text_frame
    bottleneck_text.text = "병목자재\nBottleneck"
    for para in bottleneck_text.paragraphs:
        para.font.size = FONT_SIZE_HEADING
        para.font.bold = True
        para.font.color.rgb = COLOR_WHITE
        para.alignment = PP_ALIGN.CENTER

    if data.get('bottleneck'):
        for item in data['bottleneck']:
            p = bottleneck_text.add_paragraph()
            p.text = item
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_WHITE

    # Leverage (Bottom-Right) - Green
    leverage_shape = slide.shapes.add_shape(
        1,  # Rectangle
        matrix_left + quadrant_width, matrix_top + quadrant_height,
        quadrant_width, quadrant_height
    )
    leverage_shape.fill.solid()
    leverage_shape.fill.fore_color.rgb = COLOR_LEVERAGE
    leverage_shape.line.color.rgb = COLOR_WHITE
    leverage_shape.line.width = Pt(2)

    leverage_text = leverage_shape.text_frame
    leverage_text.text = "레버리지자재\nLeverage"
    for para in leverage_text.paragraphs:
        para.font.size = FONT_SIZE_HEADING
        para.font.bold = True
        para.font.color.rgb = COLOR_WHITE
        para.alignment = PP_ALIGN.CENTER

    if data.get('leverage'):
        for item in data['leverage']:
            p = leverage_text.add_paragraph()
            p.text = item
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_WHITE

    # Routine (Bottom-Left) - Gray
    routine_shape = slide.shapes.add_shape(
        1,  # Rectangle
        matrix_left, matrix_top + quadrant_height,
        quadrant_width, quadrant_height
    )
    routine_shape.fill.solid()
    routine_shape.fill.fore_color.rgb = COLOR_ROUTINE
    routine_shape.line.color.rgb = COLOR_WHITE
    routine_shape.line.width = Pt(2)

    routine_text = routine_shape.text_frame
    routine_text.text = "일상자재\nRoutine"
    for para in routine_text.paragraphs:
        para.font.size = FONT_SIZE_HEADING
        para.font.bold = True
        para.font.color.rgb = COLOR_WHITE
        para.alignment = PP_ALIGN.CENTER

    if data.get('routine'):
        for item in data['routine']:
            p = routine_text.add_paragraph()
            p.text = item
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_WHITE

    # Axis labels
    # X-axis (Purchase Impact)
    x_label = slide.shapes.add_textbox(
        matrix_left, matrix_top + quadrant_height * 2 + Inches(0.1),
        quadrant_width * 2, Inches(0.3)
    )
    x_text = x_label.text_frame
    x_text.text = "구매 임팩트 (Purchase Impact) →"
    x_text.paragraphs[0].font.size = FONT_SIZE_BODY
    x_text.paragraphs[0].font.bold = True
    x_text.paragraphs[0].font.color.rgb = COLOR_DARK_GRAY
    x_text.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Y-axis (Supply Risk)
    y_label = slide.shapes.add_textbox(
        matrix_left - Inches(1.2), matrix_top,
        Inches(1.0), quadrant_height * 2
    )
    y_text = y_label.text_frame
    y_text.text = "공급 리스크\n(Supply Risk)\n↑"
    for para in y_text.paragraphs:
        para.font.size = FONT_SIZE_BODY
        para.font.bold = True
        para.font.color.rgb = COLOR_DARK_GRAY
        para.alignment = PP_ALIGN.CENTER

    # Caption
    if data.get('caption'):
        caption_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(6.8), Inches(9.0), Inches(0.3)
        )
        caption_text = caption_box.text_frame
        caption_text.text = data['caption']
        caption_text.paragraphs[0].font.size = FONT_SIZE_CAPTION
        caption_text.paragraphs[0].font.color.rgb = COLOR_MED_GRAY
        caption_text.paragraphs[0].alignment = PP_ALIGN.CENTER

    return slide


def main():
    """Generate PPTX from JSON"""
    # Load JSON data
    with open('skill/data/part1-session1-complete.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Create presentation
    prs = create_presentation()

    print(f"Generating Part 1 PPTX: {data['title']}")
    print(f"Total slides: {len(data['slides'])}")

    # Generate each slide
    for slide_data in data['slides']:
        layout = slide_data['layout']
        slide_content = slide_data['data']

        print(f"  Creating slide {slide_data['id']}: {layout} - {slide_content.get('title', 'Cover')}")

        if layout == 'cover':
            add_cover_slide(prs, slide_content)
        elif layout == 'list-bullets':
            add_list_bullets_slide(prs, slide_content)
        elif layout == 'content-2col':
            add_content_2col_slide(prs, slide_content)
        elif layout == 'diagram-kraljic':
            add_kraljic_matrix_slide(prs, slide_content)
        else:
            print(f"    WARNING: Unknown layout '{layout}', skipping")

    # Save PPTX
    output_path = 'Part1_Session1_StrategicInventory.pptx'
    prs.save(output_path)
    print(f"\n✓ PPTX saved: {output_path}")
    print(f"  Slides: {len(prs.slides)}")
    print(f"  Dimensions: {SLIDE_WIDTH/914400:.2f}\" × {SLIDE_HEIGHT/914400:.2f}\"")
    print(f"  Design: S4HANA monochrome + Kraljic colors")


if __name__ == '__main__':
    main()
