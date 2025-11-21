#!/usr/bin/env python3
"""
Additional SVG Generator for Part 2 - 5 new diagrams
Generates professional SVG diagrams for PPTX insertion
"""

def generate_matrix_door_chart():
    """슬라이드 6: 자재군별 소싱 전략 매트릭스 (7×4 표)"""
    svg = '''<svg width="800" height="500" xmlns="http://www.w3.org/2000/svg">
  <!-- Title -->
  <text x="400" y="30" font-family="Malgun Gothic, Arial" font-size="20" font-weight="bold"
        text-anchor="middle" fill="#333">자재군별 소싱 전략 매트릭스</text>

  <!-- Header Row -->
  <rect x="50" y="60" width="150" height="50" fill="#E6E6E6" stroke="#666" stroke-width="1"/>
  <text x="125" y="90" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#333">구분</text>

  <rect x="200" y="60" width="150" height="50" fill="#E67E22" stroke="#666" stroke-width="1"/>
  <text x="275" y="90" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#FFF">🔴 병목자재</text>

  <rect x="350" y="60" width="150" height="50" fill="#27AE60" stroke="#666" stroke-width="1"/>
  <text x="425" y="90" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#FFF">🟢 레버리지</text>

  <rect x="500" y="60" width="150" height="50" fill="#8E44AD" stroke="#666" stroke-width="1"/>
  <text x="575" y="90" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#FFF">🟣 전략자재</text>

  <rect x="650" y="60" width="150" height="50" fill="#95A5A6" stroke="#666" stroke-width="1"/>
  <text x="725" y="90" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#FFF">⚪ 일상자재</text>

  <!-- Row 1: 핵심 목표 -->
  <rect x="50" y="110" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="140" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">핵심 목표</text>

  <rect x="200" y="110" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="140" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">공급 안정성</text>

  <rect x="350" y="110" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="140" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">원가 경쟁력</text>

  <rect x="500" y="110" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="140" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">상호 성장</text>

  <rect x="650" y="110" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="140" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">효율성</text>

  <!-- Row 2: 소싱 전략 -->
  <rect x="50" y="160" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="190" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">소싱 전략</text>

  <rect x="200" y="160" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="190" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">공급선 다변화</text>

  <rect x="350" y="160" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="190" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">경쟁 촉진</text>

  <rect x="500" y="160" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="185" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">전략적</text>
  <text x="575" y="197" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">파트너십</text>

  <rect x="650" y="160" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="185" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">통합 &amp;</text>
  <text x="725" y="197" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">자동화</text>

  <!-- Row 3: 공급업체 수 -->
  <rect x="50" y="210" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="240" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">공급업체 수</text>

  <rect x="200" y="210" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="240" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">2~3개</text>

  <rect x="350" y="210" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="240" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">5개 이상</text>

  <rect x="500" y="210" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="240" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">1~2개 (전략적)</text>

  <rect x="650" y="210" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="240" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">1~2개 (통합)</text>

  <!-- Row 4: 계약 기간 -->
  <rect x="50" y="260" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="290" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">계약 기간</text>

  <rect x="200" y="260" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="285" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">중장기</text>
  <text x="275" y="297" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">(1~3년)</text>

  <rect x="350" y="260" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="285" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">단기</text>
  <text x="425" y="297" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">(6개월~1년)</text>

  <rect x="500" y="260" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="285" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">장기</text>
  <text x="575" y="297" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">(3~5년)</text>

  <rect x="650" y="260" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="285" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">중기</text>
  <text x="725" y="297" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">(1~2년)</text>

  <!-- Row 5: 관계 유형 -->
  <rect x="50" y="310" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="340" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">관계 유형</text>

  <rect x="200" y="310" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="340" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">협력적</text>

  <rect x="350" y="310" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="340" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">거래적</text>

  <rect x="500" y="310" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="340" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">파트너십</text>

  <rect x="650" y="310" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="340" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">효율적</text>

  <!-- Row 6: 협상 방식 -->
  <rect x="50" y="360" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="390" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">협상 방식</text>

  <rect x="200" y="360" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="390" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">안정성 중심</text>

  <rect x="350" y="360" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="390" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">가격 경쟁</text>

  <rect x="500" y="360" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="390" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">Win-Win</text>

  <rect x="650" y="360" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="390" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">표준화</text>

  <!-- Row 7: 정보 공유 -->
  <rect x="50" y="410" width="150" height="50" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="125" y="440" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">정보 공유</text>

  <rect x="200" y="410" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="275" y="440" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">중간 수준</text>

  <rect x="350" y="410" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="425" y="440" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">제한적</text>

  <rect x="500" y="410" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="575" y="440" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">고도 공유</text>

  <rect x="650" y="410" width="150" height="50" fill="#FFF" stroke="#666" stroke-width="1"/>
  <text x="725" y="440" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#333">최소화</text>
</svg>'''

    with open('SVG_ASSETS/slide6_matrix_door_chart.svg', 'w', encoding='utf-8') as f:
        f.write(svg)
    print("✅ Generated: slide6_matrix_door_chart.svg")


def generate_bottleneck_multi_sourcing():
    """슬라이드 9: 병목자재 공급선 다변화 프로세스"""
    svg = '''<svg width="700" height="400" xmlns="http://www.w3.org/2000/svg">
  <!-- Title -->
  <text x="350" y="30" font-family="Malgun Gothic, Arial" font-size="18" font-weight="bold"
        text-anchor="middle" fill="#333">병목자재 공급선 다변화 (Dual Sourcing)</text>

  <!-- Step 1: 메인 공급업체 -->
  <rect x="50" y="80" width="140" height="80" rx="10" fill="#E67E22" stroke="#333" stroke-width="2"/>
  <circle cx="80" cy="105" r="15" fill="#FFF" stroke="#333" stroke-width="2"/>
  <text x="80" y="112" font-family="Arial" font-size="16" font-weight="bold"
        text-anchor="middle" fill="#E67E22">1</text>
  <text x="120" y="125" font-family="Malgun Gothic" font-size="13" font-weight="bold" fill="#FFF">메인 공급업체</text>
  <text x="120" y="142" font-family="Malgun Gothic" font-size="10" fill="#FFF">70-80% 물량</text>

  <!-- Arrow 1 -->
  <path d="M 190 120 L 230 120" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrowhead)"/>

  <!-- Step 2: 백업 공급업체 -->
  <rect x="230" y="80" width="140" height="80" rx="10" fill="#E67E22" stroke="#333" stroke-width="2"/>
  <circle cx="260" cy="105" r="15" fill="#FFF" stroke="#333" stroke-width="2"/>
  <text x="260" y="112" font-family="Arial" font-size="16" font-weight="bold"
        text-anchor="middle" fill="#E67E22">2</text>
  <text x="300" y="125" font-family="Malgun Gothic" font-size="13" font-weight="bold" fill="#FFF">백업 공급업체</text>
  <text x="300" y="142" font-family="Malgun Gothic" font-size="10" fill="#FFF">20-30% 물량</text>

  <!-- Arrow 2 -->
  <path d="M 370 120 L 410 120" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrowhead)"/>

  <!-- Step 3: 지역 분산 -->
  <rect x="410" y="80" width="140" height="80" rx="10" fill="#E67E22" stroke="#333" stroke-width="2"/>
  <circle cx="440" cy="105" r="15" fill="#FFF" stroke="#333" stroke-width="2"/>
  <text x="440" y="112" font-family="Arial" font-size="16" font-weight="bold"
        text-anchor="middle" fill="#E67E22">3</text>
  <text x="480" y="125" font-family="Malgun Gothic" font-size="13" font-weight="bold" fill="#FFF">지역 분산</text>
  <text x="480" y="142" font-family="Malgun Gothic" font-size="10" fill="#FFF">다른 국가/지역</text>

  <!-- Benefits Section -->
  <rect x="50" y="200" width="500" height="160" rx="10" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="300" y="225" font-family="Malgun Gothic" font-size="14" font-weight="bold"
        text-anchor="middle" fill="#333">✓ 4가지 다변화 방법</text>

  <!-- Benefit 1 -->
  <rect x="70" y="245" width="220" height="45" rx="5" fill="#FFF" stroke="#CCC" stroke-width="1"/>
  <text x="180" y="262" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">Dual Sourcing</text>
  <text x="180" y="277" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">메인 + 백업 체계</text>

  <!-- Benefit 2 -->
  <rect x="310" y="245" width="220" height="45" rx="5" fill="#FFF" stroke="#CCC" stroke-width="1"/>
  <text x="420" y="262" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">지역적 분산</text>
  <text x="420" y="277" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">다른 지역/국가 확보</text>

  <!-- Benefit 3 -->
  <rect x="70" y="300" width="220" height="45" rx="5" fill="#FFF" stroke="#CCC" stroke-width="1"/>
  <text x="180" y="317" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">기술 이전</text>
  <text x="180" y="332" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">신규 공급업체 육성</text>

  <!-- Benefit 4 -->
  <rect x="310" y="300" width="220" height="45" rx="5" fill="#FFF" stroke="#CCC" stroke-width="1"/>
  <text x="420" y="317" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">대체재 개발</text>
  <text x="420" y="332" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">설계 변경 검토</text>

  <!-- Arrow marker definition -->
  <defs>
    <marker id="arrowhead" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">
      <polygon points="0 0, 10 3, 0 6" fill="#666"/>
    </marker>
  </defs>
</svg>'''

    with open('SVG_ASSETS/slide9_bottleneck_multi_sourcing.svg', 'w', encoding='utf-8') as f:
        f.write(svg)
    print("✅ Generated: slide9_bottleneck_multi_sourcing.svg")


def generate_consolidation_before_after():
    """슬라이드 16: 레버리지 통합 구매 Before-After"""
    svg = '''<svg width="700" height="350" xmlns="http://www.w3.org/2000/svg">
  <!-- Title -->
  <text x="350" y="30" font-family="Malgun Gothic, Arial" font-size="18" font-weight="bold"
        text-anchor="middle" fill="#333">통합 구매를 통한 협상력 강화</text>

  <!-- BEFORE Section -->
  <rect x="50" y="60" width="280" height="250" rx="10" fill="#FFEBEE" stroke="#E74C3C" stroke-width="2"/>
  <text x="190" y="85" font-family="Malgun Gothic" font-size="14" font-weight="bold"
        text-anchor="middle" fill="#E74C3C">❌ BEFORE: 분산 구매</text>

  <!-- 10 suppliers circles -->
  <circle cx="100" cy="120" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="100" y="125" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">A사</text>

  <circle cx="150" cy="120" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="150" y="125" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">B사</text>

  <circle cx="200" cy="120" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="200" y="125" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">C사</text>

  <circle cx="250" cy="120" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="250" y="125" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">D사</text>

  <circle cx="100" cy="170" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="100" y="175" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">E사</text>

  <circle cx="150" cy="170" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="150" y="175" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">F사</text>

  <circle cx="200" cy="170" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="200" y="175" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">G사</text>

  <circle cx="250" cy="170" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="250" y="175" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">H사</text>

  <circle cx="125" cy="220" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="125" y="225" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">I사</text>

  <circle cx="225" cy="220" r="18" fill="#FFF" stroke="#E74C3C" stroke-width="1.5"/>
  <text x="225" y="225" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#333">J사</text>

  <!-- Problems -->
  <text x="190" y="260" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#666">
    ⚠️ 개별 물량 작음 → 협상력 약함
  </text>
  <text x="190" y="280" font-family="Malgun Gothic" font-size="10" text-anchor="middle" fill="#666">
    ⚠️ 관리 비용 높음
  </text>

  <!-- Arrow -->
  <path d="M 340 180 L 360 180" stroke="#27AE60" stroke-width="3" fill="none" marker-end="url(#greenarrow)"/>
  <text x="350" y="165" font-family="Arial" font-size="16" font-weight="bold" fill="#27AE60">→</text>

  <!-- AFTER Section -->
  <rect x="370" y="60" width="280" height="250" rx="10" fill="#E8F8F5" stroke="#27AE60" stroke-width="2"/>
  <text x="510" y="85" font-family="Malgun Gothic" font-size="14" font-weight="bold"
        text-anchor="middle" fill="#27AE60">✅ AFTER: 통합 구매</text>

  <!-- 3-5 suppliers (larger circles) -->
  <circle cx="420" cy="140" r="28" fill="#FFF" stroke="#27AE60" stroke-width="2"/>
  <text x="420" y="138" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#333">A사</text>
  <text x="420" y="152" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">60%</text>

  <circle cx="510" cy="140" r="25" fill="#FFF" stroke="#27AE60" stroke-width="2"/>
  <text x="510" y="138" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#333">B사</text>
  <text x="510" y="152" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">30%</text>

  <circle cx="600" cy="140" r="20" fill="#FFF" stroke="#27AE60" stroke-width="2"/>
  <text x="600" y="143" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#333">C사</text>
  <text x="600" y="155" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">10%</text>

  <!-- Benefits -->
  <text x="510" y="210" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#27AE60">✓ 개별 물량 3-10배 증가</text>
  <text x="510" y="230" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#27AE60">✓ 협상력 대폭 향상</text>
  <text x="510" y="250" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#27AE60">✓ 관리 효율성 개선</text>
  <text x="510" y="270" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#27AE60">✓ 단가 10-20% 절감</text>

  <!-- Arrow markers -->
  <defs>
    <marker id="greenarrow" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">
      <polygon points="0 0, 10 3, 0 6" fill="#27AE60"/>
    </marker>
  </defs>
</svg>'''

    with open('SVG_ASSETS/slide16_consolidation_before_after.svg', 'w', encoding='utf-8') as f:
        f.write(svg)
    print("✅ Generated: slide16_consolidation_before_after.svg")


def generate_supplier_consolidation():
    """슬라이드 28: 일상자재 공급업체 통합 (원스톱 쇼핑)"""
    svg = '''<svg width="650" height="400" xmlns="http://www.w3.org/2000/svg">
  <!-- Title -->
  <text x="325" y="30" font-family="Malgun Gothic, Arial" font-size="18" font-weight="bold"
        text-anchor="middle" fill="#333">일상자재: 원스톱 쇼핑 (One-Stop Shopping)</text>

  <!-- Central Supplier -->
  <rect x="225" y="80" width="200" height="100" rx="15" fill="#95A5A6" stroke="#333" stroke-width="3"/>
  <text x="325" y="115" font-family="Malgun Gothic" font-size="16" font-weight="bold"
        text-anchor="middle" fill="#FFF">통합 공급업체</text>
  <text x="325" y="135" font-family="Malgun Gothic" font-size="12"
        text-anchor="middle" fill="#FFF">(MRO 전문업체)</text>
  <text x="325" y="155" font-family="Arial" font-size="11"
        text-anchor="middle" fill="#FFF">1~2개 업체로 통합</text>

  <!-- Category 1 -->
  <ellipse cx="120" cy="250" rx="70" ry="35" fill="#E6E6E6" stroke="#666" stroke-width="1.5"/>
  <text x="120" y="245" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">사무용품</text>
  <text x="120" y="260" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">문구, 종이 등</text>
  <path d="M 165 235 L 235 165" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrow)"/>

  <!-- Category 2 -->
  <ellipse cx="250" cy="290" rx="70" ry="35" fill="#E6E6E6" stroke="#666" stroke-width="1.5"/>
  <text x="250" y="285" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">청소용품</text>
  <text x="250" y="300" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">세제, 도구 등</text>
  <path d="M 280 265 L 305 180" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrow)"/>

  <!-- Category 3 -->
  <ellipse cx="400" cy="290" rx="70" ry="35" fill="#E6E6E6" stroke="#666" stroke-width="1.5"/>
  <text x="400" y="285" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">전기/전자</text>
  <text x="400" y="300" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">전구, 배터리 등</text>
  <path d="M 370 265 L 345 180" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrow)"/>

  <!-- Category 4 -->
  <ellipse cx="530" cy="250" rx="70" ry="35" fill="#E6E6E6" stroke="#666" stroke-width="1.5"/>
  <text x="530" y="245" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">소모성 공구</text>
  <text x="530" y="260" font-family="Malgun Gothic" font-size="9"
        text-anchor="middle" fill="#666">드릴, 톱 등</text>
  <path d="M 485 235 L 415 165" stroke="#666" stroke-width="2" fill="none" marker-end="url(#arrow)"/>

  <!-- Benefits box -->
  <rect x="50" y="340" width="550" height="50" rx="8" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="325" y="360" font-family="Malgun Gothic" font-size="11" font-weight="bold"
        text-anchor="middle" fill="#333">✓ 발주 간소화  |  ✓ 관리 비용 감소  |  ✓ 월간 통합 결제  |  ✓ E-Procurement 연계</text>

  <!-- Arrow marker -->
  <defs>
    <marker id="arrow" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">
      <polygon points="0 0, 10 3, 0 6" fill="#666"/>
    </marker>
  </defs>
</svg>'''

    with open('SVG_ASSETS/slide28_supplier_consolidation.svg', 'w', encoding='utf-8') as f:
        f.write(svg)
    print("✅ Generated: slide28_supplier_consolidation.svg")


def generate_scorecard_template():
    """슬라이드 34: Supplier Scorecard 템플릿"""
    svg = '''<svg width="750" height="450" xmlns="http://www.w3.org/2000/svg">
  <!-- Title -->
  <text x="375" y="30" font-family="Malgun Gothic, Arial" font-size="18" font-weight="bold"
        text-anchor="middle" fill="#333">Supplier Scorecard 평가 체계</text>

  <!-- 5 evaluation categories in a radial layout -->

  <!-- Center: Total Score -->
  <circle cx="375" cy="230" r="50" fill="#1A5276" stroke="#333" stroke-width="2"/>
  <text x="375" y="220" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#FFF">총점</text>
  <text x="375" y="242" font-family="Arial" font-size="24" font-weight="bold"
        text-anchor="middle" fill="#FFF">100</text>

  <!-- Category 1: Quality (30%) - Top -->
  <rect x="305" y="80" width="140" height="70" rx="10" fill="#E74C3C" stroke="#333" stroke-width="2"/>
  <text x="375" y="102" font-family="Malgun Gothic" font-size="13" font-weight="bold"
        text-anchor="middle" fill="#FFF">1. 품질 (30%)</text>
  <text x="375" y="120" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">불량률 (PPM)</text>
  <text x="375" y="135" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">검사 통과율</text>
  <path d="M 375 180 L 375 150" stroke="#666" stroke-width="2" fill="none" marker-end="url(#ar)"/>

  <!-- Category 2: Delivery (30%) - Top Right -->
  <rect x="520" y="130" width="140" height="70" rx="10" fill="#E67E22" stroke="#333" stroke-width="2"/>
  <text x="590" y="152" font-family="Malgun Gothic" font-size="13" font-weight="bold"
        text-anchor="middle" fill="#FFF">2. 납기 (30%)</text>
  <text x="590" y="170" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">납기 준수율 (OTD)</text>
  <text x="590" y="185" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">리드타임 안정성</text>
  <path d="M 425 230 L 520 165" stroke="#666" stroke-width="2" fill="none" marker-end="url(#ar)"/>

  <!-- Category 3: Price (20%) - Bottom Right -->
  <rect x="520" y="260" width="140" height="70" rx="10" fill="#F39C12" stroke="#333" stroke-width="2"/>
  <text x="590" y="282" font-family="Malgun Gothic" font-size="13" font-weight="bold"
        text-anchor="middle" fill="#FFF">3. 가격 (20%)</text>
  <text x="590" y="300" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">시장가 대비 수준</text>
  <text x="590" y="315" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">원가 절감 기여도</text>
  <path d="M 425 230 L 520 295" stroke="#666" stroke-width="2" fill="none" marker-end="url(#ar)"/>

  <!-- Category 4: Collaboration (10%) - Bottom Left -->
  <rect x="90" y="260" width="140" height="70" rx="10" fill="#3498DB" stroke="#333" stroke-width="2"/>
  <text x="160" y="282" font-family="Malgun Gothic" font-size="13" font-weight="bold"
        text-anchor="middle" fill="#FFF">4. 협력 (10%)</text>
  <text x="160" y="300" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">정보 공유 수준</text>
  <text x="160" y="315" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">개선 제안 건수</text>
  <path d="M 325 230 L 230 295" stroke="#666" stroke-width="2" fill="none" marker-end="url(#ar)"/>

  <!-- Category 5: Risk (10%) - Top Left -->
  <rect x="90" y="130" width="140" height="70" rx="10" fill="#9B59B6" stroke="#333" stroke-width="2"/>
  <text x="160" y="152" font-family="Malgun Gothic" font-size="13" font-weight="bold"
        text-anchor="middle" fill="#FFF">5. 리스크 (10%)</text>
  <text x="160" y="170" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">재무 건전성</text>
  <text x="160" y="185" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#FFF">지속가능성</text>
  <path d="M 325 230 L 230 165" stroke="#666" stroke-width="2" fill="none" marker-end="url(#ar)"/>

  <!-- Grade classification -->
  <rect x="150" y="370" width="450" height="60" rx="8" fill="#F0F0F0" stroke="#666" stroke-width="1"/>
  <text x="375" y="390" font-family="Malgun Gothic" font-size="12" font-weight="bold"
        text-anchor="middle" fill="#333">등급 분류</text>
  <text x="220" y="410" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#27AE60">A (90+)</text>
  <text x="320" y="410" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#3498DB">B (70-89)</text>
  <text x="430" y="410" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#F39C12">C (50-69)</text>
  <text x="530" y="410" font-family="Malgun Gothic" font-size="10"
        text-anchor="middle" fill="#E74C3C">D (&lt;50)</text>

  <!-- Arrow marker -->
  <defs>
    <marker id="ar" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto">
      <polygon points="0 0, 10 3, 0 6" fill="#666"/>
    </marker>
  </defs>
</svg>'''

    with open('SVG_ASSETS/slide34_scorecard_template.svg', 'w', encoding='utf-8') as f:
        f.write(svg)
    print("✅ Generated: slide34_scorecard_template.svg")


if __name__ == "__main__":
    print("Generating 5 additional SVG diagrams...\n")

    generate_matrix_door_chart()
    generate_bottleneck_multi_sourcing()
    generate_consolidation_before_after()
    generate_supplier_consolidation()
    generate_scorecard_template()

    print("\n" + "=" * 60)
    print("✅ All 5 additional SVGs generated!")
    print("=" * 60)
    print("\nTotal SVG count: 11 (6 existing + 5 new)")
    print("\nFiles in SVG_ASSETS/:")
    print("  1. slide5_bottleneck_process.svg (existing)")
    print("  2. slide6_matrix_door_chart.svg (NEW)")
    print("  3. slide9_bottleneck_multi_sourcing.svg (NEW)")
    print("  4. slide9_leverage_bidding.svg (existing)")
    print("  5. slide11_tco_comparison.svg (existing)")
    print("  6. slide12_partnership.svg (existing)")
    print("  7. slide15_eprocurement.svg (existing)")
    print("  8. slide16_consolidation_before_after.svg (NEW)")
    print("  9. slide21_toyota_pillars.svg (existing)")
    print(" 10. slide28_supplier_consolidation.svg (NEW)")
    print(" 11. slide34_scorecard_template.svg (NEW)")
