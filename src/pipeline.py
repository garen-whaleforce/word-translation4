"""
Pipeline Module - 端到端處理管線

流程:
1. PDF 解析 → 結構化 JSON
2. 翻譯 (termbase 保護 + LLM)
3. Word 模板回填
4. 生成驗證報告
"""
from __future__ import annotations
import json
import logging
from pathlib import Path
from typing import Optional, Union, Dict, Any
from dataclasses import dataclass, field

from .cb_parser import CBParser, ParseResult, ClauseItem, OverviewItem
from .termbase import Termbase, load_termbase_from_json, create_combined_termbase
from .translation_service import TranslationService, TranslationQAReport
from .word_filler import WordFiller, FillResult
from .config import settings

logger = logging.getLogger(__name__)


@dataclass
class PipelineConfig:
    """管線配置"""
    template_path: Optional[Path] = None
    glossary_path: Optional[Path] = None
    tm_path: Optional[Path] = None
    bulk_model: str = "gemini-2.5-flash"
    refine_model: str = "gpt-5.2"
    enable_refinement: bool = True
    dry_run: bool = False


@dataclass
class PipelineResult:
    """管線執行結果"""
    # 輸入資訊
    pdf_filename: str
    template_used: str = ""

    # 處理結果
    parse_result: Optional[ParseResult] = None
    fill_result: Optional[FillResult] = None
    qa_report: Optional[TranslationQAReport] = None

    # 輸出檔案
    output_docx: str = ""
    extracted_json: str = ""
    qa_report_json: str = ""

    # 統計
    total_clauses: int = 0
    translated_segments: int = 0

    # 錯誤
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pdf_filename": self.pdf_filename,
            "template_used": self.template_used,
            "output_docx": self.output_docx,
            "extracted_json": self.extracted_json,
            "qa_report_json": self.qa_report_json,
            "total_clauses": self.total_clauses,
            "translated_segments": self.translated_segments,
            "errors": self.errors,
            "warnings": self.warnings,
            "parse_result": self.parse_result.to_dict() if self.parse_result else None,
            "fill_result": self.fill_result.to_dict() if self.fill_result else None,
            "qa_report": self.qa_report.to_dict() if self.qa_report else None
        }


class Pipeline:
    """端到端處理管線"""

    def __init__(self, config: Optional[PipelineConfig] = None):
        """
        初始化管線

        Args:
            config: 管線配置
        """
        self.config = config or PipelineConfig()
        self.termbase: Optional[Termbase] = None
        self.translation_service: Optional[TranslationService] = None

        # 載入術語庫
        self._load_termbase()

    def _load_termbase(self):
        """載入術語庫"""
        glossary_path = self.config.glossary_path or Path("rules/en_zh_glossary_preferred.json")
        tm_path = self.config.tm_path or Path("rules/en_zh_translation_memory.csv")

        if glossary_path.exists():
            try:
                from .termbase import create_combined_termbase
                self.termbase = create_combined_termbase(glossary_path, tm_path)
                logger.info(f"Loaded termbase with {len(self.termbase)} terms")
            except Exception as e:
                logger.warning(f"Failed to load termbase: {e}")
                self.termbase = Termbase()
        else:
            self.termbase = Termbase()

        # 初始化翻譯服務
        self.translation_service = TranslationService(
            termbase=self.termbase,
            bulk_model=self.config.bulk_model,
            refine_model=self.config.refine_model,
            dry_run=self.config.dry_run
        )

    def process(
        self,
        pdf_path: Union[str, Path],
        output_dir: Union[str, Path],
        template_path: Optional[Union[str, Path]] = None
    ) -> PipelineResult:
        """
        執行完整處理管線

        Args:
            pdf_path: PDF 檔案路徑
            output_dir: 輸出目錄
            template_path: Word 模板路徑 (可選)

        Returns:
            PipelineResult
        """
        pdf_path = Path(pdf_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        result = PipelineResult(pdf_filename=pdf_path.name)

        # 決定模板
        template_path = template_path or self.config.template_path
        if template_path:
            result.template_used = str(template_path)

        try:
            # 1. 解析 PDF
            logger.info(f"Step 1: Parsing PDF: {pdf_path}")
            parse_result = self._parse_pdf(pdf_path)
            result.parse_result = parse_result
            result.total_clauses = len(parse_result.clauses)
            result.errors.extend(parse_result.errors)
            result.warnings.extend(parse_result.warnings)

            # 保存 extracted JSON
            extracted_json_path = output_dir / f"{pdf_path.stem}_extracted.json"
            with open(extracted_json_path, 'w', encoding='utf-8') as f:
                json.dump(parse_result.to_dict(), f, ensure_ascii=False, indent=2)
            result.extracted_json = str(extracted_json_path)
            logger.info(f"Saved extracted JSON: {extracted_json_path}")

            # 2. 翻譯
            logger.info("Step 2: Translating content...")
            translated_result, qa_report = self._translate(parse_result)
            result.qa_report = qa_report
            result.translated_segments = qa_report.translated_segments

            # 3. 填充 Word 模板 (如果有模板)
            if template_path and Path(template_path).exists():
                logger.info(f"Step 3: Filling Word template: {template_path}")
                output_docx_path = output_dir / f"{pdf_path.stem}_output.docx"
                fill_result = self._fill_template(
                    template_path,
                    translated_result,
                    output_docx_path
                )
                result.fill_result = fill_result
                result.output_docx = str(output_docx_path)
                result.errors.extend(fill_result.errors)
                result.warnings.extend(fill_result.warnings)
            else:
                result.warnings.append("No template provided, skipping Word generation")

            # 4. 保存 QA 報告
            qa_report_path = output_dir / f"{pdf_path.stem}_qa_report.json"
            with open(qa_report_path, 'w', encoding='utf-8') as f:
                json.dump(qa_report.to_dict(), f, ensure_ascii=False, indent=2)
            result.qa_report_json = str(qa_report_path)
            logger.info(f"Saved QA report: {qa_report_path}")

        except Exception as e:
            logger.error(f"Pipeline error: {e}")
            result.errors.append(str(e))

        return result

    def _parse_pdf(self, pdf_path: Path) -> ParseResult:
        """解析 PDF"""
        parser = CBParser(pdf_path)
        return parser.parse()

    def _translate(self, parse_result: ParseResult) -> tuple:
        """翻譯內容"""
        # 收集需要翻譯的文字
        texts_to_translate = []

        # Overview items
        for item in parse_result.overview_of_energy_sources:
            texts_to_translate.append(item.description)
            texts_to_translate.append(item.safeguards)

        # Energy diagram
        if parse_result.energy_source_diagram_text:
            texts_to_translate.append(parse_result.energy_source_diagram_text)

        # Clauses
        for clause in parse_result.clauses:
            texts_to_translate.append(clause.requirement_test)
            texts_to_translate.append(clause.result_remark)

        # 批量翻譯
        results = self.translation_service.translate_batch(
            texts_to_translate,
            enable_refinement=self.config.enable_refinement
        )

        # 建立翻譯映射
        translation_map = {
            r.original_text: r.translated_text
            for r in results
        }

        # 更新 parse_result 中的翻譯
        translated_result = self._apply_translations(parse_result, translation_map)

        # 生成 QA 報告
        qa_report = self.translation_service.generate_qa_report(results)

        return translated_result, qa_report

    def _apply_translations(
        self,
        parse_result: ParseResult,
        translation_map: Dict[str, str]
    ) -> ParseResult:
        """將翻譯應用到解析結果"""
        # 複製結果
        result = ParseResult(
            filename=parse_result.filename,
            trf_no=parse_result.trf_no,
            test_report_no=parse_result.test_report_no,
            errors=parse_result.errors.copy(),
            warnings=parse_result.warnings.copy()
        )

        # 翻譯 Overview items
        for item in parse_result.overview_of_energy_sources:
            result.overview_of_energy_sources.append(OverviewItem(
                hazard_clause=item.hazard_clause,
                description=translation_map.get(item.description, item.description),
                safeguards=translation_map.get(item.safeguards, item.safeguards),
                remarks=item.remarks
            ))

        # 翻譯 Energy diagram
        result.energy_source_diagram_text = translation_map.get(
            parse_result.energy_source_diagram_text,
            parse_result.energy_source_diagram_text
        )

        # 翻譯 Clauses
        for clause in parse_result.clauses:
            result.clauses.append(ClauseItem(
                clause_id=clause.clause_id,
                requirement_test=translation_map.get(clause.requirement_test, clause.requirement_test),
                result_remark=translation_map.get(clause.result_remark, clause.result_remark),
                verdict=clause.verdict,
                page_number=clause.page_number
            ))

        return result

    def _fill_template(
        self,
        template_path: Union[str, Path],
        data: ParseResult,
        output_path: Path
    ) -> FillResult:
        """填充模板"""
        filler = WordFiller(template_path)
        return filler.fill(data, output_path)


def run_pipeline(
    pdf_path: Union[str, Path],
    output_dir: Union[str, Path],
    template_path: Optional[Union[str, Path]] = None,
    dry_run: bool = False
) -> PipelineResult:
    """
    便捷函數：執行完整管線

    Args:
        pdf_path: PDF 檔案路徑
        output_dir: 輸出目錄
        template_path: Word 模板路徑
        dry_run: 是否為乾跑模式

    Returns:
        PipelineResult
    """
    config = PipelineConfig(
        template_path=Path(template_path) if template_path else None,
        dry_run=dry_run
    )
    pipeline = Pipeline(config)
    return pipeline.process(pdf_path, output_dir, template_path)
