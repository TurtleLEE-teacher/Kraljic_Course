#!/usr/bin/env python3
"""
SVG Generator for Part 2 PPTX Visual Enhancement
Creates professional diagrams that can be edited in PowerPoint
"""

import os

# ============================================================================
# COLOR CONSTANTS (Monochrome S4HANA)
# ============================================================================
COLOR_BLACK = "#000000"
COLOR_DARK_GRAY = "#333333"
COLOR_MED_GRAY = "#666666"
COLOR_LIGHT_GRAY = "#CCCCCC"
COLOR_VERY_LIGHT_GRAY = "#E6E6E6"
COLOR_WHITE = "#FFFFFF"
COLOR_ACCENT = "#1A5276"  # Dark blue

# ============================================================================
# SVG UTILITY FUNCTIONS
# ============================================================================

def create_svg_header(width, height):
    """Create SVG header with standard styling"""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<svg width="{width}" height="{height}" xmlns="http://www.w3.org/2000/svg">
<defs>
    <style>
        .box {{ fill: {COLOR_VERY_LIGHT_GRAY}; stroke: {COLOR_MED_GRAY}; stroke-width: 2; }}
        .box-dark {{ fill: {COLOR_MED_GRAY}; stroke: {COLOR_DARK_GRAY}; stroke-width: 2; }}
        .arrow {{ fill: {COLOR_MED_GRAY}; }}
        .line {{ stroke: {COLOR_MED_GRAY}; stroke-width: 2; fill: none; }}
        .text {{ font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 14px; fill: {COLOR_DARK_GRAY}; }}
        .text-bold {{ font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 16px; font-weight: bold; fill: {COLOR_BLACK}; }}
        .text-small {{ font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 11px; fill: {COLOR_MED_GRAY}; }}
        .text-white {{ font-family: 'Malgun Gothic', Arial, sans-serif; font-size: 14px; fill: {COLOR_WHITE}; }}
        .circle {{ fill: {COLOR_ACCENT}; }}
        .bg-accent {{ fill: {COLOR_ACCENT}; }}
    </style>
</defs>
'''

def create_svg_footer():
    """Create SVG footer"""
    return '</svg>'

def create_rounded_rect(x, y, width, height, rx, css_class, text=None, text_class="text"):
    """Create rounded rectangle with optional text"""
    svg = f'<rect class="{css_class}" x="{x}" y="{y}" width="{width}" height="{height}" rx="{rx}"/>\n'
    if text:
        text_x = x + width / 2
        text_y = y + height / 2 + 5  # Vertically center (approx)
        svg += f'<text class="{text_class}" x="{text_x}" y="{text_y}" text-anchor="middle">{text}</text>\n'
    return svg

def create_arrow_horizontal(x1, y, x2):
    """Create horizontal arrow from x1 to x2"""
    arrow_head_size = 10
    svg = f'<line class="line" x1="{x1}" y1="{y}" x2="{x2 - arrow_head_size}" y2="{y}"/>\n'
    svg += f'<polygon class="arrow" points="{x2-arrow_head_size},{y-5} {x2},{y} {x2-arrow_head_size},{y+5}"/>\n'
    return svg

def create_arrow_vertical(x, y1, y2):
    """Create vertical arrow from y1 to y2"""
    arrow_head_size = 10
    svg = f'<line class="line" x1="{x}" y1="{y1}" x2="{x}" y2="{y2 - arrow_head_size}"/>\n'
    svg += f'<polygon class="arrow" points="{x-5},{y2-arrow_head_size} {x},{y2} {x+5},{y2-arrow_head_size}"/>\n'
    return svg

def create_circle_with_number(x, y, radius, number):
    """Create circle with number inside"""
    svg = f'<circle class="circle" cx="{x}" cy="{y}" r="{radius}"/>\n'
    svg += f'<text class="text-white" x="{x}" y="{y+5}" text-anchor="middle" font-weight="bold">{number}</text>\n'
    return svg

# ============================================================================
# DIAGRAM GENERATORS
# ============================================================================

def generate_bottleneck_process_flow():
    """Slide 5: Bottleneck Strategy Process Flow"""
    width = 900
    height = 250

    svg = create_svg_header(width, height)

    steps = [
        "공급선 다변화",
        "이중 공급 체계",
        "장기 계약",
        "관계 강화"
    ]

    box_width = 180
    box_height = 100
    gap = 40
    start_x = 20
    y = 75

    for i, step in enumerate(steps):
        x = start_x + i * (box_width + gap)

        # Step box
        svg += create_rounded_rect(x, y, box_width, box_height, 10, "box", step, "text-bold")

        # Number circle
        svg += create_circle_with_number(x + 20, y + 20, 15, str(i + 1))

        # Arrow (except last)
        if i < len(steps) - 1:
            arrow_start = x + box_width + 5
            arrow_end = x + box_width + gap - 5
            svg += create_arrow_horizontal(arrow_start, y + box_height / 2, arrow_end)

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide5_bottleneck_process.svg'
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

def generate_leverage_bidding_flow():
    """Slide 9: Leverage Competitive Bidding Flow"""
    width = 700
    height = 400

    svg = create_svg_header(width, height)

    # Top section: Process flow
    process = ["RFQ 발송", "경쟁 입찰", "TCO 분석", "공급업체 선정"]
    box_width = 140
    box_height = 60
    gap = 30
    start_x = 50
    y = 50

    for i, step in enumerate(process):
        x = start_x + i * (box_width + gap)
        svg += create_rounded_rect(x, y, box_width, box_height, 8, "box-dark", step, "text-white")

        if i < len(process) - 1:
            arrow_start = x + box_width + 3
            arrow_end = x + box_width + gap - 3
            svg += create_arrow_horizontal(arrow_start, y + box_height / 2, arrow_end)

    # Bottom section: Key points
    key_points = [
        "• 표준화된 견적서",
        "• 다수 공급업체",
        "• 가격 경쟁 유도",
        "• 물량 통합"
    ]

    y_point = 180
    for i, point in enumerate(key_points):
        svg += create_rounded_rect(80, y_point, 540, 40, 5, "box", None)
        svg += f'<text class="text" x="100" y="{y_point + 25}">{point}</text>\n'
        y_point += 50

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide9_leverage_bidding.svg'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

def generate_tco_comparison():
    """Slide 11: TCO Analysis Comparison"""
    width = 800
    height = 350

    svg = create_svg_header(width, height)

    # Title
    svg += '<text class="text-bold" x="400" y="30" text-anchor="middle">TCO = 구매가 + 물류비 + 관세 + 품질비용 + 재고비용 + 관리비용</text>\n'

    # Two comparison boxes
    # Left: 국내 공급업체
    svg += create_rounded_rect(50, 80, 320, 240, 10, "box", None)
    svg += '<text class="text-bold" x="210" y="110" text-anchor="middle">국내 공급업체</text>\n'

    domestic = [
        "구매가: ₩100",
        "물류비: ₩5",
        "관세: ₩0",
        "품질비용: ₩2",
        "재고비용: ₩3",
        "관리비용: ₩2"
    ]

    y_pos = 140
    for item in domestic:
        svg += f'<text class="text" x="70" y="{y_pos}">{item}</text>\n'
        y_pos += 30

    svg += '<text class="text-bold" x="210" y="{}" text-anchor="middle" fill="{}">총 TCO: ₩112</text>\n'.format(y_pos + 10, COLOR_ACCENT)

    # Right: 해외 공급업체
    svg += create_rounded_rect(430, 80, 320, 240, 10, "box", None)
    svg += '<text class="text-bold" x="590" y="110" text-anchor="middle">해외 공급업체</text>\n'

    overseas = [
        "구매가: ₩85",
        "물류비: ₩15",
        "관세: ₩8",
        "품질비용: ₩5",
        "재고비용: ₩8",
        "관리비용: ₩4"
    ]

    y_pos = 140
    for item in overseas:
        svg += f'<text class="text" x="450" y="{y_pos}">{item}</text>\n'
        y_pos += 30

    svg += '<text class="text-bold" x="590" y="{}" text-anchor="middle" fill="{}">총 TCO: ₩125</text>\n'.format(y_pos + 10, COLOR_ACCENT)

    # Arrow showing winner
    svg += '<text class="text-bold" x="400" y="340" text-anchor="middle" fill="#27AE60">← 국내 공급업체 선정 (TCO 우위)</text>\n'

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide11_tco_comparison.svg'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

def generate_partnership_diagram():
    """Slide 12: Strategic Partnership Relationship"""
    width = 700
    height = 400

    svg = create_svg_header(width, height)

    # Center: Partnership box
    center_x, center_y = 350, 200
    svg += create_rounded_rect(center_x - 100, center_y - 40, 200, 80, 10, "box-dark", "전략적 파트너십", "text-white")

    # Three pillars around center
    pillars = [
        {"x": 100, "y": 50, "label": "목표 공유", "detail": "원가절감\n품질향상\n기술혁신"},
        {"x": 500, "y": 50, "label": "이익 공유", "detail": "절감액\n50/50 분배"},
        {"x": 300, "y": 320, "label": "리스크 공유", "detail": "가격변동\n공동대응"}
    ]

    for pillar in pillars:
        # Pillar box
        svg += create_rounded_rect(pillar["x"], pillar["y"], 140, 60, 8, "box", pillar["label"], "text-bold")

        # Detail text
        details = pillar["detail"].split("\n")
        detail_y = pillar["y"] + 80
        for detail in details:
            svg += f'<text class="text-small" x="{pillar["x"] + 70}" y="{detail_y}" text-anchor="middle">{detail}</text>\n'
            detail_y += 18

        # Line to center
        line_x1 = pillar["x"] + 70
        line_y1 = pillar["y"] + 60 if pillar["y"] < center_y else pillar["y"]
        svg += f'<line class="line" x1="{line_x1}" y1="{line_y1}" x2="{center_x}" y2="{center_y}"/>\n'

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide12_partnership.svg'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

def generate_eprocurement_architecture():
    """Slide 15: E-Procurement System Architecture"""
    width = 800
    height = 350

    svg = create_svg_header(width, height)

    # System flow (vertical)
    layers = [
        {"y": 30, "label": "카탈로그 구매", "detail": "사전 등록 품목 선택"},
        {"y": 110, "label": "자동 발주", "detail": "재고 부족 시 자동 생성"},
        {"y": 190, "label": "승인 자동화", "detail": "일정 금액 이하 자동 승인"},
        {"y": 270, "label": "3-Way Matching", "detail": "PO-GR-IR 자동 매칭"}
    ]

    box_width = 600
    box_height = 60
    start_x = 100

    for i, layer in enumerate(layers):
        svg += create_rounded_rect(start_x, layer["y"], box_width, box_height, 8, "box", None)
        svg += f'<text class="text-bold" x="{start_x + 30}" y="{layer["y"] + 28}">{layer["label"]}</text>\n'
        svg += f'<text class="text-small" x="{start_x + 30}" y="{layer["y"] + 48}">{layer["detail"]}</text>\n'

        # Arrow to next (except last)
        if i < len(layers) - 1:
            arrow_x = start_x + box_width / 2
            arrow_y1 = layer["y"] + box_height + 3
            arrow_y2 = layers[i + 1]["y"] - 3
            svg += create_arrow_vertical(arrow_x, arrow_y1, arrow_y2)

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide15_eprocurement.svg'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

def generate_toyota_three_pillars():
    """Slide 21: Toyota 3 Core Strategies"""
    width = 800
    height = 400

    svg = create_svg_header(width, height)

    # Title
    svg += '<text class="text-bold" x="400" y="30" text-anchor="middle" font-size="18">Toyota SRM 3대 핵심 전략</text>\n'

    # Three pillars
    pillars = [
        {
            "x": 50,
            "title": "상호 신뢰\n파트너십",
            "items": ["장기 계약", "투명한 정보", "공정한 가격"]
        },
        {
            "x": 300,
            "title": "Kaizen\n지속적 개선",
            "items": ["교육 지원", "현장 지원", "공동 해결"]
        },
        {
            "x": 550,
            "title": "성장 비전\n공유",
            "items": ["장기 예측", "투자 지원", "공동 R&D"]
        }
    ]

    for i, pillar in enumerate(pillars):
        # Pillar number
        svg += create_circle_with_number(pillar["x"] + 100, 80, 25, str(i + 1))

        # Pillar title box
        svg += create_rounded_rect(pillar["x"], 120, 200, 70, 10, "box-dark", pillar["title"], "text-white")

        # Items
        item_y = 220
        for item in pillar["items"]:
            svg += create_rounded_rect(pillar["x"] + 10, item_y, 180, 35, 5, "box", item, "text")
            item_y += 45

    svg += create_svg_footer()

    output_path = '/home/user/Kraljic_Course/SVG_ASSETS/slide21_toyota_pillars.svg'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)

    return output_path

# ============================================================================
# MAIN GENERATOR
# ============================================================================

def generate_all_svgs():
    """Generate all SVG diagrams"""
    print("=" * 80)
    print("GENERATING SVG DIAGRAMS FOR PART 2 ENHANCEMENT")
    print("=" * 80)
    print()

    diagrams = [
        ("Slide 5: Bottleneck Process Flow", generate_bottleneck_process_flow),
        ("Slide 9: Leverage Bidding Flow", generate_leverage_bidding_flow),
        ("Slide 11: TCO Comparison", generate_tco_comparison),
        ("Slide 12: Partnership Diagram", generate_partnership_diagram),
        ("Slide 15: E-Procurement Architecture", generate_eprocurement_architecture),
        ("Slide 21: Toyota Three Pillars", generate_toyota_three_pillars),
    ]

    generated = []
    for name, func in diagrams:
        print(f"✓ Generating {name}...")
        path = func()
        generated.append(path)
        print(f"  → {path}")

    print()
    print("=" * 80)
    print(f"✅ Successfully generated {len(generated)} SVG diagrams")
    print("=" * 80)
    print()

    return generated

if __name__ == "__main__":
    generate_all_svgs()
