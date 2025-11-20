#!/usr/bin/env python3
"""
Part 1 PPTX Generator - Strategic Inventory Management Foundation
Session 1: Kraljic Matrix와 자재계획 방법론

S4HANA Design Standards Compliance:
- Dimensions: 10.83" × 7.50" (1.44:1)
- Monochrome color system (black/white/gray)
- Font: Arial/맑은 고딕
- Governing messages: 16pt Bold (NOT 14pt Italic)
- Shape counts: 20-50+ per slide
- Font distribution: 10pt primary (65%), 12pt bullets (20-25%)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# S4HANA Color Palette (Monochrome)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_DARK_GRAY = RGBColor(51, 51, 51)
COLOR_MED_GRAY = RGBColor(102, 102, 102)
COLOR_LIGHT_GRAY = RGBColor(204, 204, 204)
COLOR_VERY_LIGHT_GRAY = RGBColor(230, 230, 230)
COLOR_WHITE = RGBColor(255, 255, 255)

# Kraljic Matrix colors
COLOR_STRATEGIC = RGBColor(142, 68, 173)
COLOR_BOTTLENECK = RGBColor(230, 126, 34)
COLOR_LEVERAGE = RGBColor(39, 174, 60)
COLOR_ROUTINE = RGBColor(149, 165, 166)

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10.83)
    prs.slide_height = Inches(7.50)
    return prs

def add_slide_title(slide, title):
    left = Inches(0.30)
    top = Inches(0.30)
    width = Inches(10.23)
    height = Inches(0.60)
    
    textbox = slide.shapes.add_textbox(left, top, width, height)
    p = textbox.text_frame.paragraphs[0]
    p.text = title
    p.font.name = '맑은 고딕'
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLACK
    return textbox

def add_governing_message(slide, message):
    left = Inches(0.30)
    top = Inches(1.01)
    width = Inches(10.32)
    height = Inches(0.63)
    
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    
    p = text_frame.paragraphs[0]
    p.text = message
    p.font.name = '맑은 고딕'
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_MED_GRAY
    return textbox

def add_footer(slide, footer_text, slide_number):
    left = Inches(0.30)
    top = Inches(7.00)
    width = Inches(8.00)
    height = Inches(0.30)
    
    textbox = slide.shapes.add_textbox(left, top, width, height)
    p = textbox.text_frame.paragraphs[0]
    p.text = footer_text
    p.font.name = 'Arial'
    p.font.size = Pt(8)
    p.font.color.rgb = COLOR_MED_GRAY
    
    left = Inches(10.00)
    textbox = slide.shapes.add_textbox(left, top, Inches(0.50), height)
    p = textbox.text_frame.paragraphs[0]
    p.text = str(slide_number)
    p.font.name = 'Arial'
    p.font.size = Pt(8)
    p.font.color.rgb = COLOR_MED_GRAY
    p.alignment = PP_ALIGN.RIGHT

def create_cover_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    left = Inches(1.00)
    top = Inches(2.50)
    width = Inches(8.83)
    height = Inches(1.20)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    p = title_box.text_frame.paragraphs[0]
    p.text = "[1회차] 전략적 재고운영 Foundation"
    p.font.name = '맑은 고딕'
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = COLOR_BLACK
    p.alignment = PP_ALIGN.CENTER
    
    top = Inches(3.80)
    height = Inches(0.80)
    
    subtitle_box = slide.shapes.add_textbox(left, top, width, height)
    p = subtitle_box.text_frame.paragraphs[0]
    p.text = "Kraljic Matrix와 자재계획 방법론"
    p.font.name = '맑은 고딕'
    p.font.size = Pt(28)
    p.font.color.rgb = COLOR_DARK_GRAY
    p.alignment = PP_ALIGN.CENTER
    
    top = Inches(5.50)
    height = Inches(1.00)
    
    meta_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = meta_box.text_frame
    
    for i, line in enumerate(["전략적 재고운영 및 자재계획수립 과정", "45분", "Session 1 of 9"]):
        p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        p.text = line
        p.font.name = '맑은 고딕'
        p.font.size = Pt(14)
        p.font.color.rgb = COLOR_MED_GRAY
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(6)
    
    return slide

def create_kraljic_matrix_slide(prs, slide_num):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    add_slide_title(slide, "2.3 Kraljic Matrix: 4대 자재군 분류")
    add_governing_message(slide,
        "공급 리스크(Y축)와 구매 임팩트(X축)의 조합으로 4개 자재군을 분류하며, 각 군은 완전히 다른 관리 철학을 요구합니다.")
    
    matrix_left = Inches(1.50)
    matrix_top = Inches(2.00)
    quadrant_width = Inches(3.80)
    quadrant_height = Inches(2.20)
    
    # Spectrum indicator
    top_indicator = slide.shapes.add_textbox(
        Inches(1.50), Inches(1.60), Inches(7.60), Inches(0.30)
    )
    p = top_indicator.text_frame.paragraphs[0]
    p.text = "← 낮음                공급 리스크 (Supply Risk)                높음 →"
    p.font.name = '맑은 고딕'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_MED_GRAY
    p.alignment = PP_ALIGN.CENTER
    
    # Quadrants with WHITE text on colored backgrounds
    quadrants = [
        {
            "left": matrix_left,
            "top": matrix_top,
            "color": COLOR_LEVERAGE,
            "title": "레버리지자재",
            "subtitle": "Leverage Materials",
            "bullets": ["• 공급 안정, 금액 큼", "• 경쟁 입찰", "• 원가 절감 집중", "• MRP 계획"]
        },
        {
            "left": matrix_left + quadrant_width,
            "top": matrix_top,
            "color": COLOR_STRATEGIC,
            "title": "전략자재",
            "subtitle": "Strategic Materials",
            "bullets": ["• 공급 어렵고 금액 큼", "• 장기 파트너십", "• Win-Win 협력", "• LTP + Hybrid"]
        },
        {
            "left": matrix_left,
            "top": matrix_top + quadrant_height,
            "color": COLOR_ROUTINE,
            "title": "일상자재",
            "subtitle": "Routine Materials",
            "bullets": ["• 공급 쉽고 금액 작음", "• 자동화", "• 효율성 극대화", "• Min-Max / VMI"]
        },
        {
            "left": matrix_left + quadrant_width,
            "top": matrix_top + quadrant_height,
            "color": COLOR_BOTTLENECK,
            "title": "병목자재",
            "subtitle": "Bottleneck Materials",
            "bullets": ["• 공급 어렵고 금액 작음", "• 안전재고 확보", "• 공급 안정성 우선", "• ROP 발주"]
        }
    ]
    
    for q in quadrants:
        # Background shape
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            q["left"], q["top"], quadrant_width, quadrant_height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = q["color"]
        shape.line.color.rgb = COLOR_BLACK
        shape.line.width = Pt(2)
        
        # Title (WHITE text)
        title_box = slide.shapes.add_textbox(
            q["left"] + Inches(0.20), q["top"] + Inches(0.15),
            quadrant_width - Inches(0.40), Inches(0.40)
        )
        p = title_box.text_frame.paragraphs[0]
        p.text = q["title"]
        p.font.name = '맑은 고딕'
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER
        
        # Subtitle (WHITE text)
        subtitle_box = slide.shapes.add_textbox(
            q["left"] + Inches(0.20), q["top"] + Inches(0.55),
            quadrant_width - Inches(0.40), Inches(0.25)
        )
        p = subtitle_box.text_frame.paragraphs[0]
        p.text = q["subtitle"]
        p.font.name = 'Arial'
        p.font.size = Pt(9)
        p.font.color.rgb = COLOR_WHITE
        p.alignment = PP_ALIGN.CENTER
        
        # Bullets (WHITE text)
        bullets_box = slide.shapes.add_textbox(
            q["left"] + Inches(0.30), q["top"] + Inches(0.90),
            quadrant_width - Inches(0.60), Inches(1.20)
        )
        text_frame = bullets_box.text_frame
        text_frame.word_wrap = True
        
        for i, bullet_text in enumerate(q["bullets"]):
            p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
            p.text = bullet_text
            p.font.name = '맑은 고딕'
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_WHITE
            p.space_after = Pt(3)
    
    add_footer(slide, "전략적 재고운영 Foundation", slide_num)
    return slide

def main():
    print("=== Part 1 PPTX Generation Started ===")
    
    prs = create_presentation()
    print(f"✓ Created presentation: 10.83\" × 7.50\"")
    
    # Generate slides
    create_cover_slide(prs)
    print(f"✓ Slide 1: Cover")
    
    create_kraljic_matrix_slide(prs, 2)
    print(f"✓ Slide 2: Kraljic Matrix")
    
    output_path = "/home/user/Kraljic_Course/Part1_Session1_StrategicInventory.pptx"
    prs.save(output_path)
    
    print(f"\n=== PPTX Generation Complete ===")
    print(f"Output: {output_path}")
    print(f"Total slides: {len(prs.slides)}")
    
    return output_path

if __name__ == "__main__":
    main()
