/**
 * EduPptxBuilder v2.0 - HTML→PPTX 기반 교육자료 생성 도구
 *
 * 주요 변경사항:
 * - Handlebars 템플릿 엔진 기반 HTML 생성
 * - html2pptx를 통한 고품질 PPTX 변환
 * - 3-Color Rule, MECE, Why-How-So What 프레임워크 자동 적용
 * - 텍스트 오버플로우 자동 방지
 *
 * 사용법:
 * ```javascript
 * import EduPptxBuilder from './edu-pptx-builder.js';
 *
 * const builder = new EduPptxBuilder('strategic-edu');
 * await builder.addSlide('cover', coverData);
 * await builder.addSlide('content-2col', contentData);
 * await builder.save('output.pptx');
 * ```
 */

import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import Handlebars from 'handlebars';
import chalk from 'chalk';
import PptxGenJS from 'pptxgenjs';
import { html2pptx } from '@ant/html2pptx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * EduPptxBuilder v2.0 - HTML→PPTX 기반
 */
class EduPptxBuilder {
  /**
   * @param {string} theme - 테마 이름 (기본값: 'strategic-edu')
   * @param {object} options - 추가 옵션
   */
  constructor(theme = 'strategic-edu', options = {}) {
    this.theme = theme;
    this.options = {
      debug: options.debug || false,
      strictValidation: options.strictValidation || true,
      ...options
    };

    this.slides = [];
    this.htmlSlides = [];
    this.warnings = [];
    this.errors = [];

    this.templateDir = path.join(__dirname, '..', 'templates', 'education-course');
    this.partialsDir = path.join(this.templateDir, 'partials');
    this.layoutsDir = path.join(this.templateDir, 'layouts');

    if (this.options.debug) {
      console.log(chalk.blue('EduPptxBuilder v2.0 (HTML-based) initialized'));
      console.log(chalk.gray(`  Theme: ${theme}`));
      console.log(chalk.gray(`  Template Dir: ${this.templateDir}`));
    }
  }

  /**
   * Handlebars helpers 등록
   */
  _registerHelpers() {
    // 불릿 텍스트를 <ul> 리스트로 변환
    Handlebars.registerHelper('bulletListToUl', function(text) {
      if (!text) return '';

      // "• " 또는 "- "로 시작하는 라인들을 <li>로 변환
      const lines = text.split('\n').filter(line => line.trim());
      const listItems = lines.map(line => {
        // 불릿 기호 제거
        const cleaned = line.replace(/^[•\-\*]\s*/, '');
        return `  <li>${cleaned}</li>`;
      }).join('\n');

      return new Handlebars.SafeString(`<ul>\n${listItems}\n</ul>`);
    });

    if (this.options.debug) {
      console.log(chalk.green('✓'), 'Handlebars helpers registered');
    }
  }

  /**
   * Handlebars partials 등록
   */
  async _registerPartials() {
    try {
      // Helpers 먼저 등록
      this._registerHelpers();

      const partials = ['common-styles', 'header', 'footer'];

      for (const partialName of partials) {
        const partialPath = path.join(this.partialsDir, `${partialName}.hbs`);
        const partialContent = await fs.readFile(partialPath, 'utf-8');
        Handlebars.registerPartial(partialName, partialContent);
      }

      if (this.options.debug) {
        console.log(chalk.green('✓'), 'Handlebars partials registered:', partials.join(', '));
      }
    } catch (error) {
      this.errors.push(`Failed to register partials: ${error.message}`);
      throw error;
    }
  }

  /**
   * 레이아웃 템플릿 로드 및 컴파일
   * @param {string} layoutName - 레이아웃 이름 (예: 'cover', 'content-2col')
   * @returns {Promise<Function>} Handlebars 템플릿 함수
   */
  async _loadLayout(layoutName) {
    try {
      const layoutPath = path.join(this.layoutsDir, `${layoutName}.hbs`);
      const layoutContent = await fs.readFile(layoutPath, 'utf-8');
      return Handlebars.compile(layoutContent);
    } catch (error) {
      this.errors.push(`Failed to load layout '${layoutName}': ${error.message}`);
      throw error;
    }
  }

  /**
   * HTML 생성
   * @param {string} layoutName - 레이아웃 이름
   * @param {object} data - 슬라이드 데이터
   * @returns {Promise<string>} 생성된 HTML
   */
  async generateHTML(layoutName, data) {
    try {
      // Partials 등록 (첫 호출 시 한 번만)
      if (!Handlebars.partials['common-styles']) {
        await this._registerPartials();
      }

      // 레이아웃 템플릿 로드 및 컴파일
      const template = await this._loadLayout(layoutName);

      // 데이터 컴파일 (세션 색상 자동 적용)
      const html = template({
        ...data,
        session: data.session || 1
      });

      if (this.options.debug) {
        console.log(chalk.green('✓'), `HTML generated for layout: ${layoutName}`);
      }

      return html;
    } catch (error) {
      this.errors.push(`HTML generation failed for '${layoutName}': ${error.message}`);
      throw error;
    }
  }

  /**
   * 슬라이드 추가 (통합 메서드)
   * @param {string} layout - 레이아웃 이름 ('cover', 'content-2col', 'list-bullets')
   * @param {object} data - 슬라이드 데이터
   */
  async addSlide(layout, data) {
    try {
      const html = await this.generateHTML(layout, data);

      this.htmlSlides.push({
        layout,
        data,
        html
      });

      this.slides.push({ type: layout, data });

      if (this.options.debug) {
        console.log(chalk.green('✓'), `Slide added: ${layout} - ${data.title || 'Untitled'}`);
      }
    } catch (error) {
      this.errors.push(`Failed to add slide '${layout}': ${error.message}`);
      console.error(chalk.red('Error adding slide:'), error);
    }
  }

  /**
   * HTML → PPTX 변환 (html2pptx 사용)
   * @param {string} outputPath - PPTX 파일 출력 경로
   * @returns {Promise<void>}
   */
  async convertToPPTX(outputPath) {
    try {
      // PptxGenJS 인스턴스 생성
      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'EDUCATION', width: 10, height: 5.625 });
      pptx.layout = 'EDUCATION';

      // 임시 HTML 파일 디렉토리
      const tempDir = path.join(__dirname, '..', 'output', 'temp-html');
      await fs.mkdir(tempDir, { recursive: true });

      // 각 HTML 슬라이드를 파일로 저장하고 html2pptx로 변환
      for (let i = 0; i < this.htmlSlides.length; i++) {
        const htmlSlide = this.htmlSlides[i];
        const tempHtmlPath = path.join(tempDir, `slide-${i + 1}.html`);

        // HTML 파일 저장
        await fs.writeFile(tempHtmlPath, htmlSlide.html, 'utf-8');

        // html2pptx로 슬라이드 추가
        await html2pptx(tempHtmlPath, pptx);

        if (this.options.debug) {
          console.log(chalk.green('✓'), `Converted slide ${i + 1}/${this.htmlSlides.length}`);
        }
      }

      // PPTX 파일 저장
      await pptx.writeFile({ fileName: outputPath });

      // 임시 파일 정리
      if (!this.options.debug) {
        await fs.rm(tempDir, { recursive: true, force: true });
      } else {
        console.log(chalk.gray(`  Temp HTML files kept for debugging: ${tempDir}`));
      }

      if (this.options.debug) {
        console.log(chalk.green('✓ PPTX conversion completed'));
      }

    } catch (error) {
      this.errors.push(`PPTX conversion failed: ${error.message}`);
      throw error;
    }
  }

  /**
   * PPTX 파일 저장
   * @param {string} filename - 출력 파일명
   */
  async save(filename) {
    try {
      if (this.htmlSlides.length === 0) {
        console.warn(chalk.yellow('Warning: No slides to save'));
        return {
          success: false,
          message: 'No slides generated',
          slideCount: 0,
          warnings: this.warnings,
          errors: this.errors
        };
      }

      // HTML → PPTX 변환
      await this.convertToPPTX(filename);

      if (this.options.debug) {
        console.log(chalk.green('\n✓ PPTX saved successfully:'), chalk.cyan(filename));
        console.log(chalk.gray(`  Total slides: ${this.slides.length}`));
        console.log(chalk.gray(`  Warnings: ${this.warnings.length}`));
        console.log(chalk.gray(`  Errors: ${this.errors.length}`));
      }

      return {
        success: true,
        filename,
        slideCount: this.slides.length,
        warnings: this.warnings,
        errors: this.errors
      };
    } catch (error) {
      console.error(chalk.red('Error saving PPTX:'), error);
      this.errors.push(`Save failed: ${error.message}`);
      throw error;
    }
  }

  /**
   * 품질 보고서 생성
   * @returns {object} 품질 보고서
   */
  getQualityReport() {
    return {
      totalSlides: this.slides.length,
      warnings: this.warnings,
      errors: this.errors,
      slideTypes: this.slides.reduce((acc, s) => {
        acc[s.type] = (acc[s.type] || 0) + 1;
        return acc;
      }, {}),
      timestamp: new Date().toISOString()
    };
  }

  // ========================================
  // 레거시 호환성 메서드 (v1.0 → v2.0 마이그레이션)
  // ========================================

  /**
   * 표지 슬라이드 추가 (레거시)
   * @deprecated Use addSlide('cover', data) instead
   */
  async addCoverSlide(data) {
    console.warn(chalk.yellow('Warning:'), 'addCoverSlide() is deprecated. Use addSlide("cover", data) instead.');
    return this.addSlide('cover', data);
  }

  /**
   * 2단 본문 슬라이드 추가 (레거시)
   * @deprecated Use addSlide('content-2col', data) instead
   */
  async addContent2ColSlide(data) {
    console.warn(chalk.yellow('Warning:'), 'addContent2ColSlide() is deprecated. Use addSlide("content-2col", data) instead.');
    return this.addSlide('content-2col', data);
  }

  /**
   * 불릿 리스트 슬라이드 추가 (레거시)
   * @deprecated Use addSlide('list-bullets', data) instead
   */
  async addBulletListSlide(data) {
    console.warn(chalk.yellow('Warning:'), 'addBulletListSlide() is deprecated. Use addSlide("list-bullets", data) instead.');
    return this.addSlide('list-bullets', data);
  }
}

export default EduPptxBuilder;
