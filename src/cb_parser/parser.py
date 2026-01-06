"""
CB PDF Parser - 解析 CB TRF PDF 並擷取結構化資料

主要擷取:
1. OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS (安全防護總攬表)
2. ENERGY SOURCE DIAGRAM (能量源圖)
3. Clause / Requirement + Test / Result - Remark / Verdict (條款表)
"""
from __future__ import annotations
import re
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple, Union
from dataclasses import dataclass, field, asdict

import pdfplumber

logger = logging.getLogger(__name__)


@dataclass
class OverviewItem:
    """Overview 表中的單一項目"""
    hazard_clause: str  # e.g., "5", "6", "7"
    description: str
    safeguards: str
    remarks: str = ""


@dataclass
class ClauseItem:
    """條款表中的單一條款"""
    clause_id: str  # e.g., "4.1.1", "6.4.5.2"
    requirement_test: str  # Requirement + Test 欄位內容
    result_remark: str  # Result + Remark 欄位內容
    verdict: str  # P, N/A, Fail, etc.
    page_number: int = 0


@dataclass
class ParseResult:
    """PDF 解析結果"""
    filename: str
    trf_no: str = ""
    test_report_no: str = ""
    overview_of_energy_sources: List[OverviewItem] = field(default_factory=list)
    energy_source_diagram_text: str = ""
    clauses: List[ClauseItem] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """轉換為字典格式"""
        return {
            "filename": self.filename,
            "trf_no": self.trf_no,
            "test_report_no": self.test_report_no,
            "overview_of_energy_sources": [asdict(item) for item in self.overview_of_energy_sources],
            "energy_source_diagram_text": self.energy_source_diagram_text,
            "clauses": [asdict(item) for item in self.clauses],
            "total_clauses": len(self.clauses),
            "errors": self.errors,
            "warnings": self.warnings
        }


class CBParser:
    """CB PDF 解析器"""

    # 關鍵段落標識
    OVERVIEW_MARKERS = [
        "OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS",
        "OVERVIEW OF SAFEGUARDS",
        "ENERGY SOURCES AND SAFEGUARDS"
    ]

    ENERGY_DIAGRAM_MARKERS = [
        "ENERGY SOURCE DIAGRAM",
        "ENERGY SOURCES DIAGRAM"
    ]

    CLAUSE_TABLE_MARKERS = [
        "Clause",
        "Requirement",
        "Verdict"
    ]

    # Verdict 識別模式
    VERDICT_PATTERN = re.compile(r'^(P|N/?A|NA|Fail|F|--|-)$', re.IGNORECASE)

    # Clause ID 模式 (e.g., 4, 4.1, 4.1.1, G.7.3.2.1)
    CLAUSE_ID_PATTERN = re.compile(r'^([A-Z]\.)?(\d+)(\.\d+)*$')

    def __init__(self, pdf_path: str | Path):
        """
        初始化解析器

        Args:
            pdf_path: PDF 檔案路徑
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        self.result = ParseResult(filename=self.pdf_path.name)

    def parse(self) -> ParseResult:
        """
        解析 PDF 並返回結構化結果

        Returns:
            ParseResult 包含所有擷取的資料
        """
        logger.info(f"Starting to parse: {self.pdf_path}")

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                # 1. 擷取基本資訊 (TRF No, Test Report No)
                self._extract_basic_info(pdf)

                # 2. 擷取 Overview 表
                self._extract_overview(pdf)

                # 3. 擷取 Energy Source Diagram
                self._extract_energy_diagram(pdf)

                # 4. 擷取 Clause 表
                self._extract_clauses(pdf)

        except Exception as e:
            logger.error(f"Error parsing PDF: {e}")
            self.result.errors.append(f"PDF parsing failed: {str(e)}")

        logger.info(f"Parsing complete. Found {len(self.result.clauses)} clauses")
        return self.result

    def _extract_basic_info(self, pdf: pdfplumber.PDF):
        """擷取基本資訊"""
        # 通常在前幾頁
        for page_num, page in enumerate(pdf.pages[:5]):
            text = page.extract_text() or ""

            # TRF No
            trf_match = re.search(r'TRF\s*No[\.:]?\s*([A-Z0-9_]+)', text, re.IGNORECASE)
            if trf_match and not self.result.trf_no:
                self.result.trf_no = trf_match.group(1)

            # Test Report No
            report_match = re.search(r'Test\s+Report\s+No[\.:]?\s*([A-Z0-9\-]+)', text, re.IGNORECASE)
            if report_match and not self.result.test_report_no:
                self.result.test_report_no = report_match.group(1)

    def _extract_overview(self, pdf: pdfplumber.PDF):
        """擷取 Overview of Energy Sources and Safeguards 表"""
        overview_page = None
        overview_page_num = -1

        # 尋找包含 Overview 標記的頁面
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            for marker in self.OVERVIEW_MARKERS:
                if marker in text.upper():
                    overview_page = page
                    overview_page_num = page_num
                    break
            if overview_page:
                break

        if not overview_page:
            self.result.warnings.append("Could not find OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS section")
            return

        logger.debug(f"Found Overview section on page {overview_page_num + 1}")

        # 嘗試擷取表格
        tables = overview_page.extract_tables()
        for table in tables:
            if not table or len(table) < 2:
                continue

            # 檢查是否為 Overview 表 (通常有 hazard clause 欄位)
            header = table[0] if table else []
            header_text = " ".join(str(cell) for cell in header if cell).upper()

            if "HAZARD" in header_text or "CLAUSE" in header_text or "SAFEGUARD" in header_text:
                for row in table[1:]:
                    if not row or all(not cell for cell in row):
                        continue

                    # 解析 Overview 項目
                    item = self._parse_overview_row(row)
                    if item:
                        self.result.overview_of_energy_sources.append(item)

    def _parse_overview_row(self, row: List) -> Optional[OverviewItem]:
        """解析 Overview 表的單一列"""
        if not row or len(row) < 2:
            return None

        # 清理並組合欄位
        cells = [str(cell).strip() if cell else "" for cell in row]

        # 嘗試識別 hazard clause (通常是數字如 5, 6, 7, 8, 9, 10)
        hazard_clause = ""
        description = ""
        safeguards = ""
        remarks = ""

        for i, cell in enumerate(cells):
            if re.match(r'^\d+$', cell):
                hazard_clause = cell
            elif i == 0:
                hazard_clause = cell
            elif i == 1:
                description = cell
            elif i == 2:
                safeguards = cell
            elif i >= 3:
                remarks = cell

        if hazard_clause or description:
            return OverviewItem(
                hazard_clause=hazard_clause,
                description=description,
                safeguards=safeguards,
                remarks=remarks
            )
        return None

    def _extract_energy_diagram(self, pdf: pdfplumber.PDF):
        """擷取 Energy Source Diagram 區段"""
        diagram_text_parts = []
        in_diagram_section = False

        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = text.split('\n')

            for line in lines:
                line_upper = line.upper().strip()

                # 檢查是否進入 Energy Diagram 區段
                for marker in self.ENERGY_DIAGRAM_MARKERS:
                    if marker in line_upper:
                        in_diagram_section = True
                        break

                # 檢查是否離開區段 (遇到 Clause 表開始)
                if in_diagram_section:
                    if any(marker in line_upper for marker in ["CLAUSE", "REQUIREMENT", "VERDICT"]):
                        if re.search(r'\bCLAUSE\b.*\bVERDICT\b', line_upper):
                            in_diagram_section = False
                            break

                    diagram_text_parts.append(line)

        self.result.energy_source_diagram_text = "\n".join(diagram_text_parts).strip()

        if not self.result.energy_source_diagram_text:
            self.result.warnings.append("Could not find ENERGY SOURCE DIAGRAM section")

    def _extract_clauses(self, pdf: pdfplumber.PDF):
        """擷取 Clause 條款表"""
        all_clauses = []
        in_clause_section = False

        for page_num, page in enumerate(pdf.pages):
            # 嘗試從表格擷取
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue

                # 檢查是否為 Clause 表
                header = table[0] if table else []
                header_text = " ".join(str(cell) for cell in header if cell).upper()

                if "CLAUSE" in header_text and ("VERDICT" in header_text or "REQUIREMENT" in header_text):
                    in_clause_section = True
                    # 解析表格內容
                    for row in table[1:]:
                        clause = self._parse_clause_row(row, page_num + 1)
                        if clause:
                            all_clauses.append(clause)
                elif in_clause_section:
                    # 繼續解析後續頁面的表格
                    for row in table:
                        clause = self._parse_clause_row(row, page_num + 1)
                        if clause:
                            all_clauses.append(clause)

        # 去重並排序
        seen_clauses = set()
        for clause in all_clauses:
            key = (clause.clause_id, clause.verdict)
            if key not in seen_clauses:
                seen_clauses.add(key)
                self.result.clauses.append(clause)

        if not self.result.clauses:
            self.result.errors.append("Could not extract any clauses from the PDF")

    def _parse_clause_row(self, row: List, page_number: int) -> Optional[ClauseItem]:
        """解析 Clause 表的單一列"""
        if not row or all(not cell for cell in row):
            return None

        cells = [str(cell).strip() if cell else "" for cell in row]

        # 識別 Clause ID (第一個符合模式的儲存格)
        clause_id = ""
        requirement_test = ""
        result_remark = ""
        verdict = ""

        for i, cell in enumerate(cells):
            # 檢查是否為 Clause ID
            if self.CLAUSE_ID_PATTERN.match(cell):
                clause_id = cell
            # 檢查是否為 Verdict
            elif self.VERDICT_PATTERN.match(cell):
                verdict = self._normalize_verdict(cell)
            # 其他內容
            elif cell and not clause_id:
                # 可能是 Clause ID 的變體
                if re.match(r'^[A-Z]?\d+', cell):
                    clause_id = cell
            elif cell and clause_id:
                if not requirement_test:
                    requirement_test = cell
                elif not result_remark:
                    result_remark = cell

        # 最後一個欄位通常是 Verdict
        if not verdict and cells:
            last_cell = cells[-1]
            if self.VERDICT_PATTERN.match(last_cell):
                verdict = self._normalize_verdict(last_cell)

        if clause_id:
            return ClauseItem(
                clause_id=clause_id,
                requirement_test=requirement_test,
                result_remark=result_remark,
                verdict=verdict,
                page_number=page_number
            )
        return None

    def _normalize_verdict(self, verdict: str) -> str:
        """標準化 Verdict 值"""
        verdict = verdict.upper().strip()
        if verdict in ["P", "PASS"]:
            return "P"
        elif verdict in ["N/A", "NA", "N.A."]:
            return "N/A"
        elif verdict in ["F", "FAIL"]:
            return "Fail"
        elif verdict in ["--", "-"]:
            return "--"
        return verdict


def parse_cb_pdf(pdf_path: Union[str, Path]) -> ParseResult:
    """
    便捷函數：解析 CB PDF

    Args:
        pdf_path: PDF 檔案路徑

    Returns:
        ParseResult 包含所有擷取的資料
    """
    parser = CBParser(pdf_path)
    return parser.parse()
