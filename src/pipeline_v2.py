#!/usr/bin/env python3
"""
Pipeline v2 - 整合 segmenter, energy diagram, appended tables

流程:
1. PDF 區塊偵測 (segmenter)
2. 提取 Overview + Energy Diagram
3. 解析 Clause 並按章節分群
4. 提取附表
5. 翻譯 (可選)
6. 回填 Word 模板
7. 生成 QA 報告
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Internal imports
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.parse_cb_pdf_v2 import CBParserV2, ParseResultV2
from src.extract.segmenter import PDFSegmenter, SegmentResult
from src.extract.energy_diagram import EnergyDiagramExtractor, EnergyDiagramResult
from src.extract.appended_tables import AppendedTablesExtractor, AppendedTablesResult
from scripts.fill_word_ast_b import ASTBWordFiller, FillResult

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class QAMetrics:
    """QA 指標"""
    overview_count: int = 0
    chapter_count: int = 0
    clause_count: int = 0
    appended_table_count: int = 0
    energy_diagram_found: bool = False

    # Verdict 統計
    verdict_p: int = 0
    verdict_na: int = 0
    verdict_fail: int = 0
    verdict_other: int = 0

    # 翻譯統計
    translated_segments: int = 0
    translation_errors: int = 0

    # 回填統計
    tables_filled: int = 0
    clauses_filled: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class QAReport:
    """QA 報告"""
    filename: str
    generated_at: str = ""
    metrics: QAMetrics = field(default_factory=QAMetrics)

    # 區塊資訊
    segments: Optional[Dict[str, Any]] = None

    # 錯誤與警告
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    # 比對結果 (與參考 DOC 比較)
    similarity_score: float = 0.0
    diff_items: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "filename": self.filename,
            "generated_at": self.generated_at,
            "metrics": self.metrics.to_dict(),
            "segments": self.segments,
            "errors": self.errors,
            "warnings": self.warnings,
            "similarity_score": self.similarity_score,
            "diff_items": self.diff_items
        }


@dataclass
class PipelineV2Result:
    """Pipeline v2 執行結果"""
    # 輸入
    pdf_filename: str
    template_used: str = ""

    # 中間結果
    segment_result: Optional[SegmentResult] = None
    parse_result: Optional[ParseResultV2] = None
    energy_diagram: Optional[EnergyDiagramResult] = None
    appended_tables: Optional[AppendedTablesResult] = None
    fill_result: Optional[FillResult] = None

    # 輸出檔案
    output_docx: str = ""
    extracted_json: str = ""
    qa_report_json: str = ""
    energy_diagram_png: str = ""

    # QA 報告
    qa_report: Optional[QAReport] = None

    # 統計
    total_clauses: int = 0
    chapters_count: int = 0
    appended_tables_count: int = 0

    # 錯誤
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pdf_filename": self.pdf_filename,
            "template_used": self.template_used,
            "output_docx": self.output_docx,
            "extracted_json": self.extracted_json,
            "qa_report_json": self.qa_report_json,
            "energy_diagram_png": self.energy_diagram_png,
            "total_clauses": self.total_clauses,
            "chapters_count": self.chapters_count,
            "appended_tables_count": self.appended_tables_count,
            "errors": self.errors,
            "warnings": self.warnings,
            "qa_report": self.qa_report.to_dict() if self.qa_report else None
        }


class PipelineV2:
    """Pipeline v2 - 完整處理管線"""

    def __init__(
        self,
        template_path: Optional[Path] = None,
        translate_func: Optional[Callable[[str], str]] = None,
        dry_run: bool = False
    ):
        """
        初始化 Pipeline

        Args:
            template_path: Word 模板路徑
            translate_func: 翻譯函數 (text -> translated_text)
            dry_run: 乾跑模式 (不執行翻譯)
        """
        self.template_path = Path(template_path) if template_path else None
        self.translate_func = translate_func
        self.dry_run = dry_run

    def process(
        self,
        pdf_path: Path,
        output_dir: Path
    ) -> PipelineV2Result:
        """
        執行完整處理管線

        Args:
            pdf_path: PDF 檔案路徑
            output_dir: 輸出目錄

        Returns:
            PipelineV2Result
        """
        pdf_path = Path(pdf_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        result = PipelineV2Result(pdf_filename=pdf_path.name)
        if self.template_path:
            result.template_used = str(self.template_path)

        try:
            # 1. PDF 區塊偵測
            logger.info("Step 1: PDF 區塊偵測...")
            segmenter = PDFSegmenter(pdf_path)
            result.segment_result = segmenter.segment()

            # 保存 segments.json
            segments_path = output_dir / f"{pdf_path.stem}_segments.json"
            with open(segments_path, 'w', encoding='utf-8') as f:
                json.dump(result.segment_result.to_dict(), f, ensure_ascii=False, indent=2)

            # 2. 解析 PDF (v2)
            logger.info("Step 2: 解析 PDF...")
            parser = CBParserV2(pdf_path)
            result.parse_result = parser.parse()
            result.total_clauses = len(result.parse_result.clauses)
            result.chapters_count = len(result.parse_result.clause_tables)
            result.errors.extend(result.parse_result.errors)
            result.warnings.extend(result.parse_result.warnings)

            # 保存 extracted.json
            extracted_path = output_dir / f"{pdf_path.stem}_extracted.json"
            with open(extracted_path, 'w', encoding='utf-8') as f:
                json.dump(result.parse_result.to_dict(), f, ensure_ascii=False, indent=2)
            result.extracted_json = str(extracted_path)

            # 3. 提取 Energy Diagram
            logger.info("Step 3: 提取 Energy Diagram...")
            energy_extractor = EnergyDiagramExtractor(pdf_path)
            result.energy_diagram = energy_extractor.extract(output_dir=output_dir)
            if result.energy_diagram.diagram_image_path:
                result.energy_diagram_png = result.energy_diagram.diagram_image_path

            # 4. 提取附表
            logger.info("Step 4: 提取附表...")
            appended_extractor = AppendedTablesExtractor(pdf_path)
            result.appended_tables = appended_extractor.extract()
            result.appended_tables_count = len(result.appended_tables.tables)

            # 保存 appended_tables.json
            appended_path = output_dir / f"{pdf_path.stem}_appended_tables.json"
            with open(appended_path, 'w', encoding='utf-8') as f:
                json.dump(result.appended_tables.to_dict(), f, ensure_ascii=False, indent=2)

            # 5. 填充 Word 模板
            if self.template_path and self.template_path.exists():
                logger.info(f"Step 5: 填充 Word 模板: {self.template_path.name}...")
                output_docx_path = output_dir / f"{pdf_path.stem}_output.docx"

                filler = ASTBWordFiller(self.template_path)
                result.fill_result = filler.fill(
                    parse_result=result.parse_result,
                    output_path=output_docx_path,
                    translate_func=self.translate_func if not self.dry_run else None,
                    appended_tables=result.appended_tables.tables if result.appended_tables else None
                )
                result.output_docx = str(output_docx_path)
                result.errors.extend(result.fill_result.errors)
                result.warnings.extend(result.fill_result.warnings)
            else:
                result.warnings.append("No template provided, skipping Word generation")

            # 6. 生成 QA 報告
            logger.info("Step 6: 生成 QA 報告...")
            result.qa_report = self._generate_qa_report(result)

            # 保存 QA 報告
            qa_path = output_dir / f"{pdf_path.stem}_qa_report.json"
            with open(qa_path, 'w', encoding='utf-8') as f:
                json.dump(result.qa_report.to_dict(), f, ensure_ascii=False, indent=2)
            result.qa_report_json = str(qa_path)

            logger.info("Pipeline 完成!")

        except Exception as e:
            logger.error(f"Pipeline 錯誤: {e}")
            result.errors.append(str(e))

        return result

    def _generate_qa_report(self, result: PipelineV2Result) -> QAReport:
        """生成 QA 報告"""
        report = QAReport(
            filename=result.pdf_filename,
            generated_at=datetime.utcnow().isoformat()
        )

        metrics = report.metrics

        # 區塊資訊
        if result.segment_result:
            report.segments = result.segment_result.to_dict().get("summary")

        # Overview
        if result.parse_result:
            metrics.overview_count = len(result.parse_result.overview_items)
            metrics.chapter_count = len(result.parse_result.clause_tables)
            metrics.clause_count = len(result.parse_result.clauses)

            # Verdict 統計
            for clause in result.parse_result.clauses:
                v = clause.verdict.upper() if clause.verdict else ""
                if v == "P":
                    metrics.verdict_p += 1
                elif v in ("N/A", "NA"):
                    metrics.verdict_na += 1
                elif v in ("F", "FAIL"):
                    metrics.verdict_fail += 1
                else:
                    metrics.verdict_other += 1

        # Energy Diagram
        if result.energy_diagram:
            metrics.energy_diagram_found = result.energy_diagram.found

        # Appended Tables
        if result.appended_tables:
            metrics.appended_table_count = len(result.appended_tables.tables)

        # 回填統計
        if result.fill_result:
            metrics.tables_filled = result.fill_result.tables_filled
            metrics.clauses_filled = result.fill_result.clauses_filled

        # 錯誤與警告
        report.errors = result.errors.copy()
        report.warnings = result.warnings.copy()

        return report


def run_pipeline_v2(
    pdf_path: Path,
    output_dir: Path,
    template_path: Optional[Path] = None,
    translate_func: Optional[Callable[[str], str]] = None,
    dry_run: bool = False
) -> PipelineV2Result:
    """
    便捷函數：執行 Pipeline v2

    Args:
        pdf_path: PDF 檔案路徑
        output_dir: 輸出目錄
        template_path: Word 模板路徑
        translate_func: 翻譯函數
        dry_run: 乾跑模式

    Returns:
        PipelineV2Result
    """
    pipeline = PipelineV2(
        template_path=template_path,
        translate_func=translate_func,
        dry_run=dry_run
    )
    return pipeline.process(pdf_path, output_dir)


def main():
    """CLI 入口"""
    import argparse

    parser = argparse.ArgumentParser(description='CB PDF 處理 Pipeline v2')
    parser.add_argument('pdf', help='PDF 檔案路徑')
    parser.add_argument('--template', '-t', default='templates/AST-B.docx', help='Word 模板路徑')
    parser.add_argument('--output', '-o', default='output', help='輸出目錄')
    parser.add_argument('--translate', action='store_true', help='啟用翻譯')
    parser.add_argument('--dry-run', action='store_true', help='乾跑模式')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    template_path = Path(args.template) if args.template else None
    output_dir = Path(args.output) / pdf_path.stem

    # 翻譯函數
    translate_func = None
    if args.translate and not args.dry_run:
        try:
            from scripts.translate_and_compare import PDFTranslator
            translator = PDFTranslator()

            def translate_func(text):
                if not text or not text.strip():
                    return text
                results = translator.translate_batch([text])
                return results[0] if results else text
        except Exception as e:
            logger.warning(f"無法載入翻譯器: {e}")

    result = run_pipeline_v2(
        pdf_path=pdf_path,
        output_dir=output_dir,
        template_path=template_path,
        translate_func=translate_func,
        dry_run=args.dry_run
    )

    print(f"\n{'='*60}")
    print("Pipeline v2 執行結果")
    print(f"{'='*60}")
    print(f"PDF: {result.pdf_filename}")
    print(f"模板: {result.template_used}")
    print(f"條款數: {result.total_clauses}")
    print(f"章節數: {result.chapters_count}")
    print(f"附表數: {result.appended_tables_count}")

    if result.qa_report:
        m = result.qa_report.metrics
        print(f"\nVerdict 統計:")
        print(f"  P: {m.verdict_p}")
        print(f"  N/A: {m.verdict_na}")
        print(f"  Fail: {m.verdict_fail}")
        print(f"  Other: {m.verdict_other}")

    print(f"\n輸出檔案:")
    print(f"  Word: {result.output_docx}")
    print(f"  JSON: {result.extracted_json}")
    print(f"  QA Report: {result.qa_report_json}")

    if result.errors:
        print(f"\n錯誤 ({len(result.errors)}):")
        for e in result.errors[:5]:
            print(f"  - {e}")

    if result.warnings:
        print(f"\n警告 ({len(result.warnings)}):")
        for w in result.warnings[:5]:
            print(f"  - {w}")


if __name__ == '__main__':
    main()
