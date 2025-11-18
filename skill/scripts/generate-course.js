/**
 * Generate Course - 교육 과정 PPTX 자동 생성 스크립트
 *
 * 사용법:
 * node scripts/generate-course.js data/session1-scm-kraljic.json
 * node scripts/generate-course.js data/session1-scm-kraljic.json --output custom-output.pptx
 * node scripts/generate-course.js data/*.json --batch
 *
 * 옵션:
 * --output, -o: 출력 파일명 지정
 * --batch, -b: 여러 JSON 파일 일괄 처리
 * --debug, -d: 디버그 모드
 * --report, -r: 품질 보고서 생성
 */

import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import chalk from 'chalk';
import EduPptxBuilder from './edu-pptx-builder.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * 명령줄 인자 파싱
 */
function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    inputFiles: [],
    output: null,
    batch: false,
    debug: false,
    report: false
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    if (arg === '--output' || arg === '-o') {
      options.output = args[++i];
    } else if (arg === '--batch' || arg === '-b') {
      options.batch = true;
    } else if (arg === '--debug' || arg === '-d') {
      options.debug = true;
    } else if (arg === '--report' || arg === '-r') {
      options.report = true;
    } else if (!arg.startsWith('-')) {
      options.inputFiles.push(arg);
    }
  }

  return options;
}

/**
 * JSON 데이터 파일 로드
 * @param {string} filepath - JSON 파일 경로
 * @returns {Promise<object>} JSON 데이터
 */
async function loadJsonData(filepath) {
  try {
    const content = await fs.readFile(filepath, 'utf-8');
    return JSON.parse(content);
  } catch (error) {
    console.error(chalk.red(`Error loading JSON file: ${filepath}`));
    console.error(chalk.red(error.message));
    process.exit(1);
  }
}

/**
 * 슬라이드 데이터를 PPTX로 변환
 * @param {object} courseData - 교육 과정 데이터
 * @param {object} options - 옵션
 * @returns {Promise<EduPptxBuilder>} 빌더 인스턴스
 */
async function generatePPTX(courseData, options = {}) {
  const builder = new EduPptxBuilder('strategic-edu', {
    debug: options.debug
  });

  console.log(chalk.blue('\nGenerating PPTX...'));
  console.log(chalk.gray(`Course: ${courseData.course || 'Untitled'}`));
  console.log(chalk.gray(`Session: ${courseData.session || 'N/A'}`));
  console.log(chalk.gray(`Total Slides: ${courseData.slides?.length || 0}`));

  if (!courseData.slides || courseData.slides.length === 0) {
    console.warn(chalk.yellow('Warning: No slides found in data'));
    return builder;
  }

  // 슬라이드 생성 (HTML 기반)
  for (const slideData of courseData.slides) {
    try {
      const layout = slideData.layout;
      const data = { session: courseData.session, ...slideData.data };

      // 통합 addSlide() 메서드 사용 (HTML 기반)
      await builder.addSlide(layout, data);
      console.log(chalk.green('✓'), `${layout} slide added:`, data.title || `Slide ${slideData.id}`);

    } catch (error) {
      console.error(chalk.red(`✗ Error generating slide ${slideData.id} (${slideData.layout}):`), error.message);
      if (options.debug) {
        console.error(chalk.gray(error.stack));
      }
    }
  }

  return builder;
}

/**
 * 출력 파일명 생성
 * @param {string} inputFile - 입력 파일명
 * @param {string} customOutput - 사용자 지정 출력 파일명
 * @returns {string} 출력 파일명
 */
function generateOutputFilename(inputFile, customOutput) {
  if (customOutput) {
    return customOutput.endsWith('.pptx') ? customOutput : `${customOutput}.pptx`;
  }

  const basename = path.basename(inputFile, '.json');
  return `${basename}.pptx`;
}

/**
 * 품질 보고서 저장
 * @param {object} report - 품질 보고서
 * @param {string} outputPath - 출력 경로
 */
async function saveQualityReport(report, outputPath) {
  const reportPath = outputPath.replace('.pptx', '-report.json');

  try {
    await fs.writeFile(reportPath, JSON.stringify(report, null, 2));
    console.log(chalk.green('Quality report saved:'), chalk.cyan(reportPath));
  } catch (error) {
    console.error(chalk.red('Error saving quality report:'), error.message);
  }
}

/**
 * 메인 함수
 */
async function main() {
  console.log(chalk.bold.blue('\n=== EduPptxBuilder - Course Generator ===\n'));

  const options = parseArgs();

  // 입력 파일 검증
  if (options.inputFiles.length === 0) {
    console.error(chalk.red('Error: No input files specified'));
    console.log(chalk.gray('\nUsage:'));
    console.log(chalk.gray('  node generate-course.js <input.json> [options]'));
    console.log(chalk.gray('  node generate-course.js data/*.json --batch'));
    console.log(chalk.gray('\nOptions:'));
    console.log(chalk.gray('  --output, -o <file>  Output filename'));
    console.log(chalk.gray('  --batch, -b          Process multiple files'));
    console.log(chalk.gray('  --debug, -d          Enable debug mode'));
    console.log(chalk.gray('  --report, -r         Generate quality report'));
    process.exit(1);
  }

  // 출력 디렉토리 확인/생성
  const outputDir = path.join(__dirname, '..', 'output');
  try {
    await fs.mkdir(outputDir, { recursive: true });
  } catch (error) {
    console.error(chalk.red('Error creating output directory:'), error.message);
    process.exit(1);
  }

  // 파일 처리
  let successCount = 0;
  let failCount = 0;

  for (const inputFile of options.inputFiles) {
    try {
      console.log(chalk.blue(`\nProcessing: ${inputFile}`));

      // 데이터 로드
      const courseData = await loadJsonData(inputFile);

      // PPTX 생성
      const builder = await generatePPTX(courseData, options);

      // 저장
      const outputFilename = generateOutputFilename(inputFile, options.output);
      const outputPath = path.join(outputDir, outputFilename);

      const result = await builder.save(outputPath);

      if (result.success) {
        successCount++;
        console.log(chalk.green('✓ Success:'), chalk.cyan(outputPath));

        // 품질 보고서 생성
        if (options.report) {
          const report = builder.getQualityReport();
          await saveQualityReport(report, outputPath);
        }
      } else {
        failCount++;
        console.error(chalk.red('✗ Failed:'), inputFile);
      }

    } catch (error) {
      failCount++;
      console.error(chalk.red(`✗ Error processing ${inputFile}:`), error.message);
      if (options.debug) {
        console.error(error.stack);
      }
    }
  }

  // 요약
  console.log(chalk.bold.blue('\n=== Summary ==='));
  console.log(chalk.green(`Success: ${successCount}`));
  if (failCount > 0) {
    console.log(chalk.red(`Failed: ${failCount}`));
  }
  console.log();

  process.exit(failCount > 0 ? 1 : 0);
}

// 스크립트 실행
main().catch(error => {
  console.error(chalk.red('\nFatal error:'), error);
  process.exit(1);
});
