import pptxgen from "pptxgenjs";
import { html2pptx } from "@ant/html2pptx";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function createPresentation() {
  try {
    console.log("S4HANA 스타일 PPTX 생성 시작...");

    // Create a new pptx presentation
    const pptx = new pptxgen();

    // Define custom layout matching S4HANA (780pt x 540pt = 10.83" x 7.5")
    pptx.defineLayout({ name: 'S4HANA', width: 10.83, height: 7.5 });
    pptx.layout = 'S4HANA';

    pptx.author = "전략적 재고운영 교육";
    pptx.title = "Part1 Foundation - S4HANA Style";

    // HTML 파일 경로
    const slideDir = __dirname;
    const slides = [
      { file: 'slide1.html', title: 'Kraljic Matrix 개요' },
      { file: 'slide2.html', title: '전략자재 특성' },
      { file: 'slide3.html', title: '병목자재 vs 레버리지자재' },
      { file: 'slide4.html', title: '재고관리 방법론 비교' },
      { file: 'slide5.html', title: '일상자재 효율화' }
    ];

    // Convert each HTML to slide
    for (let i = 0; i < slides.length; i++) {
      const slidePath = path.join(slideDir, slides[i].file);
      console.log(`슬라이드 ${i + 1}/5 변환 중: ${slides[i].title}...`);

      try {
        await html2pptx(slidePath, pptx);
        console.log(`  ✓ 완료`);
      } catch (error) {
        console.error(`  ✗ 실패: ${error.message}`);
        throw error;
      }
    }

    // Save the presentation
    const outputPath = path.join(__dirname, '..', '..', 'output', 'Part1_본문5장_S4HANA스타일.pptx');
    console.log(`\nPPTX 저장 중: ${outputPath}`);

    await pptx.writeFile({ fileName: outputPath });

    console.log('\n✅ PPTX 생성 완료!');
    console.log(`파일 위치: ${outputPath}`);
    console.log(`슬라이드 수: ${slides.length}개`);
    console.log(`크기: 780pt × 540pt (13:9 비율)`);
    console.log(`스타일: S4HANA 모노톤`);

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
