import pptxgen from "pptxgenjs";
import { html2pptx } from "@ant/html2pptx";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function createPresentation() {
  try {
    console.log("S4HANA 스타일 개선 PPTX 생성 시작...\n");

    // Create a new pptx presentation
    const pptx = new pptxgen();

    // Define custom layout matching S4HANA (780pt x 540pt = 10.83" x 7.5")
    pptx.defineLayout({ name: 'S4HANA', width: 10.83, height: 7.5 });
    pptx.layout = 'S4HANA';

    pptx.author = "전략적 재고운영 교육";
    pptx.title = "Part1 Foundation - S4HANA Style (개선판)";

    const slideDir = __dirname;

    // ==================== Slide 1: Kraljic Matrix ====================
    console.log("슬라이드 1/5: Kraljic Matrix 개요...");
    const { slide: slide1 } = await html2pptx(
      path.join(slideDir, 'slide1-improved.html'),
      pptx
    );

    // Add 2x2 Matrix Table
    const matrixData = [
      [
        { text: "", options: { fill: { color: "F5F5F5" } } },
        { text: "낮은 공급리스크", options: { bold: true, fill: { color: "F5F5F5" }, align: "center" } },
        { text: "높은 공급리스크", options: { bold: true, fill: { color: "F5F5F5" }, align: "center" } }
      ],
      [
        { text: "높은\n영향도", options: { bold: true, fill: { color: "F5F5F5" }, align: "center", valign: "middle" } },
        { text: "레버리지자재\n경쟁입찰, 가격협상", options: { align: "center", valign: "middle", fontSize: 11 } },
        { text: "전략자재\n장기 파트너십 구축", options: { align: "center", valign: "middle", fontSize: 11 } }
      ],
      [
        { text: "낮은\n영향도", options: { bold: true, fill: { color: "F5F5F5" }, align: "center", valign: "middle" } },
        { text: "일상자재\n프로세스 자동화", options: { align: "center", valign: "middle", fontSize: 11 } },
        { text: "병목자재\n대체품 확보", options: { align: "center", valign: "middle", fontSize: 11 } }
      ]
    ];

    slide1.addTable(matrixData, {
      x: 5.5, y: 2.5, w: 4.8, h: 3.5,
      border: { pt: 1, color: "CCCCCC" },
      colW: [1.2, 1.8, 1.8],
      rowH: [0.6, 1.45, 1.45],
      fontSize: 12,
      color: "333333"
    });

    console.log("  ✓ 완료 (표 추가됨)\n");

    // ==================== Slide 2: 전략자재 ====================
    console.log("슬라이드 2/5: 전략자재 특성...");
    await html2pptx(path.join(slideDir, 'slide2-improved.html'), pptx);
    console.log("  ✓ 완료\n");

    // ==================== Slide 3: 병목 vs 레버리지 ====================
    console.log("슬라이드 3/5: 병목자재 vs 레버리지자재...");
    const { slide: slide3 } = await html2pptx(
      path.join(slideDir, 'slide3-improved.html'),
      pptx
    );

    // Add comparison table
    const comparisonData = [
      [
        { text: "구분", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "병목자재", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "레버리지자재", options: { bold: true, fill: { color: "F5F5F5" } } }
      ],
      [
        { text: "공급 리스크", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "높음", options: {} },
        { text: "낮음", options: {} }
      ],
      [
        { text: "사업 영향도", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "낮음", options: {} },
        { text: "높음", options: {} }
      ],
      [
        { text: "공급업체 수", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "소수/독점", options: {} },
        { text: "다수", options: {} }
      ],
      [
        { text: "대체 가능성", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "낮음", options: {} },
        { text: "높음", options: {} }
      ],
      [
        { text: "관리 전략", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "대체품 확보, 안전재고", options: { fontSize: 11 } },
        { text: "경쟁입찰, 가격협상", options: { fontSize: 11 } }
      ],
      [
        { text: "재고관리", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "ROP 방식", options: {} },
        { text: "MRP 방식", options: {} }
      ],
      [
        { text: "핵심 목표", options: { bold: true, fill: { color: "F5F5F5" } } },
        { text: "공급 안정성 확보", options: { fontSize: 11 } },
        { text: "원가 절감", options: { fontSize: 11 } }
      ]
    ];

    slide3.addTable(comparisonData, {
      x: 0.5, y: 1.8, w: 9.8, h: 3.2,
      border: { pt: 1, color: "CCCCCC" },
      colW: [2.4, 3.7, 3.7],
      align: "center",
      valign: "middle",
      fontSize: 12,
      color: "333333"
    });

    console.log("  ✓ 완료 (표 추가됨)\n");

    // ==================== Slide 4: ROP vs MRP ====================
    console.log("슬라이드 4/5: ROP vs MRP 비교...");
    await html2pptx(path.join(slideDir, 'slide4-improved.html'), pptx);
    console.log("  ✓ 완료 (2컬럼 비교 레이아웃)\n");

    // ==================== Slide 5: 일상자재 ====================
    console.log("슬라이드 5/5: 일상자재 효율화...");
    await html2pptx(path.join(slideDir, 'slide5-improved.html'), pptx);
    console.log("  ✓ 완료\n");

    // Save the presentation
    const outputPath = path.join(__dirname, '..', '..', 'output', 'Part1_본문5장_S4HANA스타일_개선판.pptx');
    console.log(`PPTX 저장 중: ${outputPath}`);

    await pptx.writeFile({ fileName: outputPath });

    console.log('\n' + '='.repeat(60));
    console.log('✅ PPTX 생성 완료!');
    console.log('='.repeat(60));
    console.log(`파일 위치: ${outputPath}`);
    console.log(`슬라이드 수: 5개`);
    console.log(`크기: 780pt × 540pt (13:9 비율)`);
    console.log(`스타일: S4HANA 모노톤`);
    console.log('\n개선 사항:');
    console.log('  ✓ 거버닝 메시지 추가 (모든 슬라이드)');
    console.log('  ✓ 표 추가 (Slide 1, 3)');
    console.log('  ✓ ROP vs MRP 2컬럼 비교 레이아웃 (Slide 4)');
    console.log('  ✓ 콘텐츠 밀도 증가 (80% 이상 공간 활용)');
    console.log('  ✓ 여백 최적화 (padding 30-40px)');
    console.log('='.repeat(60));

  } catch (error) {
    console.error('\n❌ 오류 발생:', error.message);
    if (error.stack) {
      console.error('\nStack trace:');
      console.error(error.stack);
    }
    process.exit(1);
  }
}

createPresentation();
