#!/usr/bin/env python3
"""
PDF 區塊偵測器 (Segmenter)

將 CB PDF 分割為以下區塊:
- overview_pages: OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS
- energy_diagram_pages: ENERGY SOURCE DIAGRAM
- clause_checklist_pages: Clause Requirement + Test Result - Remark Verdict
- appended_table_pages: TABLE: 附表
"""
import re
import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Set
from dataclasses import dataclass, field, asdict

import pdfplumber

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class PageSegment:
    """頁面區塊"""
    page_number: int
    segment_type: str
    title: str = ""
    confidence: float = 1.0


@dataclass
class SegmentResult:
    """分段結果"""
    filename: str
    total_pages: int = 0

    # 區塊頁碼
    overview_pages: List[int] = field(default_factory=list)
    energy_diagram_pages: List[int] = field(default_factory=list)
    clause_checklist_pages: List[int] = field(default_factory=list)
    appended_table_pages: List[int] = field(default_factory=list)

    # 詳細區塊資訊
    segments: List[PageSegment] = field(default_factory=list)

    # 偵測到的附表
    detected_tables: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "filename": self.filename,
            "total_pages": self.total_pages,
            "overview_pages": self.overview_pages,
            "energy_diagram_pages": self.energy_diagram_pages,
            "clause_checklist_pages": self.clause_checklist_pages,
            "appended_table_pages": self.appended_table_pages,
            "segments": [asdict(s) for s in self.segments],
            "detected_tables": self.detected_tables,
            "summary": {
                "overview_count": len(self.overview_pages),
                "energy_diagram_count": len(self.energy_diagram_pages),
                "clause_checklist_count": len(self.clause_checklist_pages),
                "appended_table_count": len(self.appended_table_pages),
            }
        }


class PDFSegmenter:
    """PDF 區塊偵測器"""

    # 區塊偵測模式
    OVERVIEW_PATTERNS = [
        r"OVERVIEW\s+OF\s+ENERGY\s+SOURCES",
        r"OVERVIEW\s+OF\s+SAFEGUARDS",
        r"Energy\s+sources?\s+and\s+safeguards",
    ]

    ENERGY_DIAGRAM_PATTERNS = [
        r"ENERGY\s+SOURCE\s+DIAGRAM",
        r"Energy\s+source\s+diagram",
        r"ES\s*/\s*PS\s*/\s*MS",
    ]

    CLAUSE_CHECKLIST_PATTERNS = [
        r"Clause\s+Requirement\s*\+?\s*Test",
        r"Result\s*-?\s*Remark\s+Verdict",
        r"^\s*Clause\s+.*Verdict\s*$",
    ]

    # TABLE: 模式
    TABLE_PATTERN = re.compile(
        r"(?:TABLE|Table):\s*([^\n]+)",
        re.IGNORECASE
    )

    # Clause ID 模式 (用於識別條款表)
    CLAUSE_ID_PATTERN = re.compile(r"^([A-Z]\.)?(\d+)(\.\d+)*$")

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        self.result = SegmentResult(filename=self.pdf_path.name)

    def segment(self) -> SegmentResult:
        """執行分段"""
        logger.info(f"開始分段: {self.pdf_path.name}")

        with pdfplumber.open(self.pdf_path) as pdf:
            self.result.total_pages = len(pdf.pages)

            for page_num, page in enumerate(pdf.pages):
                page_number = page_num + 1
                text = page.extract_text() or ""

                # 偵測各區塊
                self._detect_overview(page_number, text)
                self._detect_energy_diagram(page_number, text)
                self._detect_clause_checklist(page_number, text)
                self._detect_appended_tables(page_number, text)

        # 排序
        self.result.overview_pages.sort()
        self.result.energy_diagram_pages.sort()
        self.result.clause_checklist_pages.sort()
        self.result.appended_table_pages.sort()
        self.result.detected_tables = sorted(set(self.result.detected_tables))

        logger.info(f"分段完成: {self.result.total_pages} 頁")
        logger.info(f"  Overview: {len(self.result.overview_pages)} 頁")
        logger.info(f"  Energy Diagram: {len(self.result.energy_diagram_pages)} 頁")
        logger.info(f"  Clause Checklist: {len(self.result.clause_checklist_pages)} 頁")
        logger.info(f"  Appended Tables: {len(self.result.appended_table_pages)} 頁")
        logger.info(f"  偵測到附表: {len(self.result.detected_tables)} 個")

        return self.result

    def _detect_overview(self, page_number: int, text: str):
        """偵測 Overview 區塊"""
        for pattern in self.OVERVIEW_PATTERNS:
            if re.search(pattern, text, re.IGNORECASE):
                if page_number not in self.result.overview_pages:
                    self.result.overview_pages.append(page_number)
                    self.result.segments.append(PageSegment(
                        page_number=page_number,
                        segment_type="overview",
                        title="OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS"
                    ))
                break

    def _detect_energy_diagram(self, page_number: int, text: str):
        """偵測 Energy Diagram 區塊"""
        for pattern in self.ENERGY_DIAGRAM_PATTERNS:
            if re.search(pattern, text, re.IGNORECASE):
                if page_number not in self.result.energy_diagram_pages:
                    self.result.energy_diagram_pages.append(page_number)
                    self.result.segments.append(PageSegment(
                        page_number=page_number,
                        segment_type="energy_diagram",
                        title="ENERGY SOURCE DIAGRAM"
                    ))
                break

    def _detect_clause_checklist(self, page_number: int, text: str):
        """偵測 Clause Checklist 區塊"""
        # 方法1: 檢查表頭模式
        for pattern in self.CLAUSE_CHECKLIST_PATTERNS:
            if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
                if page_number not in self.result.clause_checklist_pages:
                    self.result.clause_checklist_pages.append(page_number)
                    self.result.segments.append(PageSegment(
                        page_number=page_number,
                        segment_type="clause_checklist",
                        title="Clause Checklist"
                    ))
                return

        # 方法2: 檢查是否有 Clause ID + Verdict 的組合
        lines = text.split('\n')
        clause_count = 0
        verdict_count = 0

        for line in lines:
            parts = line.split()
            if parts:
                # 檢查是否有 Clause ID
                if self.CLAUSE_ID_PATTERN.match(parts[0]):
                    clause_count += 1
                # 檢查是否有 Verdict
                if any(v in parts for v in ['P', 'N/A', 'NA', 'Fail', 'F']):
                    verdict_count += 1

        # 如果有足夠多的 clause + verdict，視為條款表頁
        if clause_count >= 3 and verdict_count >= 2:
            if page_number not in self.result.clause_checklist_pages:
                self.result.clause_checklist_pages.append(page_number)
                self.result.segments.append(PageSegment(
                    page_number=page_number,
                    segment_type="clause_checklist",
                    title="Clause Checklist (inferred)",
                    confidence=0.8
                ))

    def _detect_appended_tables(self, page_number: int, text: str):
        """偵測附表"""
        matches = self.TABLE_PATTERN.findall(text)

        for table_title in matches:
            table_title = table_title.strip()
            if table_title:
                self.result.detected_tables.append(table_title)

                if page_number not in self.result.appended_table_pages:
                    self.result.appended_table_pages.append(page_number)
                    self.result.segments.append(PageSegment(
                        page_number=page_number,
                        segment_type="appended_table",
                        title=f"TABLE: {table_title}"
                    ))


def segment_pdf(pdf_path: Path, output_path: Optional[Path] = None) -> SegmentResult:
    """
    便捷函數：分段 PDF

    Args:
        pdf_path: PDF 路徑
        output_path: 輸出 JSON 路徑 (可選)

    Returns:
        SegmentResult
    """
    segmenter = PDFSegmenter(pdf_path)
    result = segmenter.segment()

    if output_path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        logger.info(f"已儲存: {output_path}")

    return result


def main():
    import argparse

    parser = argparse.ArgumentParser(description='PDF 區塊偵測')
    parser.add_argument('pdf', help='PDF 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出 JSON 路徑')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    output_path = Path(args.output) if args.output else Path(f"output/{pdf_path.stem}_segments.json")

    result = segment_pdf(pdf_path, output_path)

    print(f"\n{'='*60}")
    print("分段結果")
    print(f"{'='*60}")
    print(f"總頁數: {result.total_pages}")
    print(f"Overview 頁: {result.overview_pages}")
    print(f"Energy Diagram 頁: {result.energy_diagram_pages}")
    print(f"Clause Checklist 頁: {result.clause_checklist_pages[:10]}...")
    print(f"Appended Table 頁: {result.appended_table_pages[:10]}...")
    print(f"\n偵測到的附表 ({len(result.detected_tables)} 個):")
    for t in result.detected_tables[:15]:
        print(f"  - {t}")


if __name__ == '__main__':
    main()
