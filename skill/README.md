# pptx-mslee v2.0 - HTML→PPTX 교육자료 생성 도구

**고품질 교육 프레젠테이션 자동 생성 시스템**

## 주요 변경사항 (v2.0)

### ✨ HTML 기반 시스템으로 전면 개편
- **이전 (v1.0)**: PptxGenJS 직접 사용 → 레이아웃 제한, 텍스트 오버플로우 발생
- **현재 (v2.0)**: Handlebars 템플릿 → HTML 생성 → html2pptx 변환 → 고품질 PPTX

### 🎯 품질 향상
- ✅ 텍스트 오버플로우 자동 방지
- ✅ 3-Color Rule 엄격 적용
- ✅ 10px 그리드 자동 정렬
- ✅ MECE, Why-How-So What 프레임워크 자동 적용
- ✅ 디자인 일관성 보장

### 🚀 자동화
- 72장 슬라이드를 수일 작업 → **5분 자동 생성**
- JSON 데이터 → 자동 HTML 생성 → PPTX 변환
- 품질 보고서 자동 생성

## 빠른 시작

### 1. 의존성 설치
```bash
cd ~/.claude/skills/pptx-mslee
npm install
```

### 2. 샘플 PPTX 생성
```bash
node scripts/generate-course.js data/test-sample-2slides.json --debug
```

### 3. 출력 확인
```bash
# 생성된 PPTX 파일
ls -lh output/test-sample-2slides.pptx

# PowerPoint에서 열기
start output/test-sample-2slides.pptx  # Windows
open output/test-sample-2slides.pptx   # macOS
```

## 핵심 기능

### 지원하는 레이아웃
1. **cover**: 표지 슬라이드 (세션별 색상 그라디언트)
2. **content-2col**: 2단 본문 (좌우 비교, Why-How-So What)
3. **list-bullets**: 불릿 리스트 (최대 6개 항목)

### 자동 변환 기능
- **불릿 텍스트 → `<ul>` 리스트**: 자동 변환
- **세션별 색상**: 1~7회차 자동 적용
- **디자인 검증**: html2pptx 검증 자동 통과

## 사용법

### JSON 데이터 구조
```json
{
  "course": "전략적 재고운영 및 자재계획 수립",
  "session": 1,
  "title": "SCM 개념과 Kraljic Matrix",
  "totalSlides": 2,
  "slides": [
    {
      "id": 1,
      "layout": "cover",
      "data": {
        "title": "1회차: SCM 개념",
        "subtitle": "전략적 재고운영의 기초",
        "course": "전략적 재고운영 교육",
        "date": "2025",
        "instructor": "강사명"
      }
    },
    {
      "id": 2,
      "layout": "content-2col",
      "data": {
        "title": "개선 전 vs 개선 후",
        "sessionBadge": "1회차",
        "leftTitle": "Before",
        "leftContent": "• 디자인 일관성 부족\n• 템플릿 없음",
        "rightTitle": "After",
        "rightContent": "• 3-Color Rule 적용\n• 자동 생성",
        "footer": "pptx-mslee v2.0",
        "slideNumber": 2
      }
    }
  ]
}
```

### 명령어 옵션
```bash
# 기본 생성
node scripts/generate-course.js data/session1.json

# 디버그 모드 (HTML 파일 유지)
node scripts/generate-course.js data/session1.json --debug

# 품질 보고서 생성
node scripts/generate-course.js data/session1.json --report

# 배치 처리
node scripts/generate-course.js data/*.json --batch
```

## 시스템 요구사항

- **Node.js**: v18.0.0 이상
- **npm**: 9.0.0 이상
- **의존성**:
  - `pptxgenjs`: ^3.12.0
  - `handlebars`: ^4.7.8
  - `@ant/html2pptx`: ^0.1.0
  - `sharp`: ^0.33.0
  - `chalk`: ^5.3.0

## 디렉토리 구조

```
pptx-mslee/
├── scripts/
│   ├── edu-pptx-builder.js      # v2.0 HTML 기반 빌더
│   └── generate-course.js        # 생성 스크립트
├── templates/education-course/
│   ├── layouts/
│   │   ├── cover.hbs             # 표지 템플릿
│   │   ├── content-2col.hbs      # 2단 본문 템플릿
│   │   └── list-bullets.hbs      # 불릿 리스트 템플릿
│   ├── partials/
│   │   ├── common-styles.hbs     # 공통 CSS
│   │   ├── header.hbs            # 헤더 partial
│   │   └── footer.hbs            # 푸터 partial
│   └── styles/
│       ├── variables.css         # CSS 변수
│       └── theme-strategic-edu.css
├── data/
│   └── test-sample-2slides.json  # 샘플 데이터
├── output/
│   ├── *.pptx                    # 생성된 PPTX
│   └── temp-html/                # 디버그용 HTML (--debug 시)
├── docs/
│   ├── QUICK-START.md            # 빠른 시작 가이드
│   └── TEMPLATE-GUIDE.md         # 템플릿 개발 가이드
├── SKILL.md                      # 스킬 문서 (v2.0)
├── html2pptx.md                  # html2pptx 사용 가이드
└── package.json
```

## 문서

- **SKILL.md**: 전체 기능 및 API 문서
- **QUICK-START.md**: 5분 빠른 시작
- **html2pptx.md**: HTML→PPTX 변환 가이드
- **TEMPLATE-GUIDE.md**: 템플릿 개발 가이드

## 버전 히스토리

### v2.0.0 (2025-01-04)
- ✅ HTML 기반 시스템으로 전면 개편
- ✅ Handlebars 템플릿 엔진 통합
- ✅ html2pptx 변환 파이프라인 구축
- ✅ 자동 불릿 리스트 변환
- ✅ 품질 검증 자동화
- ✅ 3-Color Rule, MECE, Why-How-So What 적용

### v1.0.0 (2024-11-03)
- PptxGenJS 직접 사용 버전
- 기본 레이아웃 3종 (cover, content-2col, list-bullets)
- 세션별 색상 시스템

## 라이선스

MIT License

## 문의

Issues: GitHub Issues
