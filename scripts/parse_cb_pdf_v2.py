#!/usr/bin/env python3
"""
CB PDF 解析器 v2 - 改進版

針對 MC-601 等 CB TRF PDF 格式優化
"""
import sys
import re
import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass, field, asdict

import pdfplumber

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class ClauseItem:
    """條款項目"""
    clause_id: str
    requirement: str  # 要求/測試描述
    result_remark: str  # 結果/備註
    verdict: str  # P, N/A, Fail, --
    page_number: int = 0


@dataclass
class OverviewItem:
    """Overview 項目"""
    hazard_clause: str
    description: str
    safeguards: str
    remarks: str = ""


@dataclass
class ParseResultV2:
    """解析結果"""
    filename: str
    trf_no: str = ""
    test_report_no: str = ""
    model: str = ""

    overview_items: List[OverviewItem] = field(default_factory=list)
    clauses: List[ClauseItem] = field(default_factory=list)

    # 章節分群的條款表 (key: "4", "5", "6", ..., "B", "M", "T" 等)
    clause_tables: Dict[str, List[ClauseItem]] = field(default_factory=dict)

    # 附表 (key: "M.3", "M.4.2", "T.7", "X", "4.1.2" 等)
    appendix_tables: Dict[str, List[List[str]]] = field(default_factory=dict)

    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "filename": self.filename,
            "trf_no": self.trf_no,
            "test_report_no": self.test_report_no,
            "model": self.model,
            "overview_items": [asdict(item) for item in self.overview_items],
            "clauses": [asdict(item) for item in self.clauses],
            "clause_tables": {k: [asdict(c) for c in v] for k, v in self.clause_tables.items()},
            "appendix_tables": self.appendix_tables,
            "total_clauses": len(self.clauses),
            "chapter_count": len(self.clause_tables),
            "appendix_count": len(self.appendix_tables),
            "errors": self.errors,
            "warnings": self.warnings
        }

    def get_chapter_verdict(self, chapter_id: str) -> Optional[str]:
        """取得章節表頭的 verdict"""
        if not chapter_id:
            return None
        normalized = chapter_id.strip().rstrip('.')
        for clause in self.clauses:
            clause_key = clause.clause_id.strip().rstrip('.')
            if clause_key == normalized:
                return clause.verdict or None
        return None


class CBParserV2:
    """CB PDF 解析器 v2"""

    # Clause ID 模式
    # 包含章節表頭 (4, 5, 6, B) 以及子章節 (4.1, 5.2.3, B.1)
    CLAUSE_ID_PATTERN = re.compile(r'^([A-Z]\.?|[A-Z]\.\d+(\.\d+)*|\d+(\.\d+)*)$')
    CHAPTER_ID_PATTERN = re.compile(r'^([A-Z]\.?|\d+)$')

    # Verdict 模式
    # 包含 WingDings dash 符號 \uf0be
    VERDICT_VALUES = {'P', 'N/A', 'NA', 'N.A.', 'Fail', 'F', '--', '-', '⎯', '—', '\uf0be'}

    # 座標分欄參數
    LINE_MERGE_TOLERANCE = 3.0
    CLAUSE_COL_MAX_RATIO = 0.35
    VERDICT_COL_MIN_RATIO = 0.75

    def __init__(self, pdf_path: Union[str, Path]):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        self.result = ParseResultV2(filename=self.pdf_path.name)

    def parse(self) -> ParseResultV2:
        """解析 PDF"""
        logger.info(f"開始解析: {self.pdf_path.name}")

        with pdfplumber.open(self.pdf_path) as pdf:
            # 1. 提取基本資訊
            self._extract_basic_info(pdf)

            # 2. 提取 Overview
            self._extract_overview(pdf)

            # 3. 提取所有 Clause
            self._extract_all_clauses(pdf)

        logger.info(f"解析完成: {len(self.result.clauses)} 個條款")
        return self.result

    def _extract_basic_info(self, pdf: pdfplumber.PDF):
        """提取基本資訊"""
        for page in pdf.pages[:5]:
            text = page.extract_text() or ""

            # Test Report No
            match = re.search(r'Report\s+Number[.\s:]+([A-Z0-9\-]+)', text, re.IGNORECASE)
            if match and not self.result.test_report_no:
                self.result.test_report_no = match.group(1)

            # Model
            match = re.search(r'Model/Type\s+reference[.\s:]+([^\n]+)', text, re.IGNORECASE)
            if match and not self.result.model:
                self.result.model = match.group(1).strip()

    def _extract_overview(self, pdf: pdfplumber.PDF):
        """提取 Overview 表"""
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""

            if "OVERVIEW OF ENERGY SOURCES" in text.upper():
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue

                    for row in table:
                        if not row or len(row) < 2:
                            continue

                        # 檢查是否為數字開頭 (hazard clause)
                        first_cell = str(row[0]).strip() if row[0] else ""
                        if re.match(r'^\d+$', first_cell):
                            item = OverviewItem(
                                hazard_clause=first_cell,
                                description=str(row[1]).strip() if len(row) > 1 and row[1] else "",
                                safeguards=str(row[2]).strip() if len(row) > 2 and row[2] else "",
                                remarks=str(row[3]).strip() if len(row) > 3 and row[3] else ""
                            )
                            self.result.overview_items.append(item)

    def _is_verdict(self, value: str) -> bool:
        """檢查是否為 Verdict 值"""
        if not value:
            return False
        value = value.strip()
        value_upper = value.upper()
        # 包含 WingDings dash 符號 \uf0be
        return value_upper in {'P', 'N/A', 'NA', 'N.A.', 'FAIL', 'F', '--', '-', '⎯', '—', ''} or value == '\uf0be'

    def _normalize_verdict(self, value: str) -> str:
        """標準化 Verdict"""
        if not value:
            return ""
        original = value.strip()
        value = original.upper()
        if value == 'P':
            return 'P'
        elif value in {'N/A', 'NA', 'N.A.'}:
            return 'N/A'
        elif value in {'F', 'FAIL'}:
            return 'Fail'
        elif value in {'--', '-', '⎯', '—'} or original == '\uf0be':
            return '--'
        return value

    def _extract_all_clauses(self, pdf: pdfplumber.PDF):
        """提取所有條款"""
        clauses_by_id: Dict[str, ClauseItem] = {}

        # 1) 優先用座標分欄法抽取 clause_id + verdict
        self._extract_clauses_from_words(pdf, clauses_by_id)
        words_count = len(clauses_by_id)
        logger.info(f"座標抽取條款: {words_count} 個")

        # 2) 用表格抽取作為補充/回填 requirement/result
        self._extract_clauses_from_tables(pdf, clauses_by_id)

        self.result.clauses = list(clauses_by_id.values())

        # 依照 Clause ID 排序
        self.result.clauses.sort(key=lambda c: self._clause_sort_key(c.clause_id))

        # 按章節分群
        self._group_clauses_by_chapter()

    def _extract_clauses_from_words(
        self,
        pdf: pdfplumber.PDF,
        clauses_by_id: Dict[str, ClauseItem]
    ):
        """用座標分欄抽取 clause_id + verdict"""
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            if not words:
                continue

            page_width = page.width or 1
            lines = self._group_words_by_line(words, tolerance=self.LINE_MERGE_TOLERANCE)
            page_has_verdict = any(self._is_verdict(w.get("text", "")) for w in words)

            for line in lines:
                line_words = sorted(line, key=lambda w: w.get("x0", 0))
                verdict_raw = self._find_verdict_in_line(line_words, page_width)
                clause_id = self._find_clause_id_in_line(
                    line_words,
                    page_width,
                    allow_chapter=bool(verdict_raw)
                )
                if not clause_id:
                    continue

                if not verdict_raw and not page_has_verdict:
                    continue

                verdict = self._normalize_verdict(verdict_raw) if verdict_raw else ""
                existing = clauses_by_id.get(clause_id)
                if existing:
                    if verdict and existing.verdict in ("", "--"):
                        existing.verdict = verdict
                    continue

                clauses_by_id[clause_id] = ClauseItem(
                    clause_id=clause_id,
                    requirement="",
                    result_remark="",
                    verdict=verdict,
                    page_number=page_num + 1,
                )

    def _extract_clauses_from_tables(
        self,
        pdf: pdfplumber.PDF,
        clauses_by_id: Dict[str, ClauseItem]
    ):
        """用表格抽取補充 requirement/result"""
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()

            for table in tables:
                if not table:
                    continue

                for row in table:
                    if not row or len(row) < 2:
                        continue

                    # 檢查第一欄是否為 Clause ID
                    first_cell = str(row[0]).strip() if row[0] else ""

                    if self.CLAUSE_ID_PATTERN.match(first_cell):
                        last_cell = str(row[-1]).strip() if row[-1] else ""
                        is_chapter = bool(self.CHAPTER_ID_PATTERN.match(first_cell))

                        if is_chapter and not self._is_verdict(last_cell):
                            continue

                        if self._is_verdict(last_cell) or last_cell == '':
                            clause_id = first_cell
                            verdict = self._normalize_verdict(last_cell)

                            requirement = ""
                            result_remark = ""

                            if len(row) >= 4:
                                requirement = str(row[1]).strip() if row[1] else ""
                                result_remark = str(row[2]).strip() if row[2] else ""
                            elif len(row) >= 3:
                                requirement = str(row[1]).strip() if row[1] else ""
                                third = str(row[2]).strip() if row[2] else ""
                                if self._is_verdict(third):
                                    verdict = self._normalize_verdict(third)
                                else:
                                    result_remark = third
                            elif len(row) >= 2:
                                requirement = str(row[1]).strip() if row[1] else ""

                            existing = clauses_by_id.get(clause_id)
                            if existing:
                                if verdict and existing.verdict in ("", "--"):
                                    existing.verdict = verdict
                                if requirement and not existing.requirement:
                                    existing.requirement = requirement
                                if result_remark and not existing.result_remark:
                                    existing.result_remark = result_remark
                                continue

                            clauses_by_id[clause_id] = ClauseItem(
                                clause_id=clause_id,
                                requirement=requirement,
                                result_remark=result_remark,
                                verdict=verdict,
                                page_number=page_num + 1,
                            )

    def _group_words_by_line(self, words: List[Dict[str, Any]], tolerance: float) -> List[List[Dict[str, Any]]]:
        """將 words 依 y 座標分群為行"""
        if not words:
            return []

        sorted_words = sorted(words, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
        lines: List[List[Dict[str, Any]]] = []
        current_line: List[Dict[str, Any]] = []
        current_top: Optional[float] = None

        for word in sorted_words:
            top = word.get("top", 0)
            if current_top is None or abs(top - current_top) <= tolerance:
                current_line.append(word)
                if current_top is None:
                    current_top = top
            else:
                lines.append(current_line)
                current_line = [word]
                current_top = top

        if current_line:
            lines.append(current_line)

        return lines

    def _find_clause_id_in_line(
        self,
        line_words: List[Dict[str, Any]],
        page_width: float,
        allow_chapter: bool
    ) -> str:
        """從單行 words 中找 clause_id"""
        for word in line_words:
            text = self._clean_clause_token(word.get("text", ""))
            if not text:
                continue
            if word.get("x0", page_width) > page_width * self.CLAUSE_COL_MAX_RATIO:
                continue
            if not allow_chapter and self.CHAPTER_ID_PATTERN.match(text):
                continue
            if self.CLAUSE_ID_PATTERN.match(text):
                return text
        return ""

    def _find_verdict_in_line(self, line_words: List[Dict[str, Any]], page_width: float) -> str:
        """從單行 words 中找 verdict"""
        candidates = [w for w in line_words if w.get("text")]
        if not candidates:
            return ""

        right_zone = [
            w for w in candidates
            if w.get("x0", 0) >= page_width * self.VERDICT_COL_MIN_RATIO
        ]

        for word in sorted(right_zone, key=lambda w: w.get("x0", 0), reverse=True):
            text = word.get("text", "").strip()
            if self._is_verdict(text):
                return text

        for word in sorted(candidates, key=lambda w: w.get("x0", 0), reverse=True):
            text = word.get("text", "").strip()
            if self._is_verdict(text):
                return text

        return ""

    def _clean_clause_token(self, token: str) -> str:
        """清理 clause token"""
        if not token:
            return ""
        return token.strip().strip(":;").rstrip(".")

    def _group_clauses_by_chapter(self):
        """將條款按章節分群"""
        for clause in self.result.clauses:
            chapter = self._get_chapter(clause.clause_id)
            if chapter not in self.result.clause_tables:
                self.result.clause_tables[chapter] = []
            self.result.clause_tables[chapter].append(clause)

        logger.info(f"章節分群: {list(self.result.clause_tables.keys())}")

    def _get_chapter(self, clause_id: str) -> str:
        """取得條款的章節"""
        # 先檢查是否有字母前綴 (B, B.1, G.5, M.3, T.7 等)
        match = re.match(r'^([A-Z])(?:\.|$)', clause_id)
        if match:
            return match.group(1)

        # 純數字條款 (4.1.1, 5.2.3 等)
        match = re.match(r'^(\d+)', clause_id)
        if match:
            return match.group(1)

        return "other"

    def get_chapter_verdict(self, chapter_id: str) -> Optional[str]:
        """取得章節表頭的 verdict"""
        if not chapter_id:
            return None
        normalized = chapter_id.strip().rstrip('.')
        for clause in self.result.clauses:
            clause_key = clause.clause_id.strip().rstrip('.')
            if clause_key == normalized:
                return clause.verdict or None
        return None

    def _clause_sort_key(self, clause_id: str) -> Tuple:
        """產生排序用的 key"""
        # 分離字母前綴和數字部分
        match = re.match(r'^([A-Z])?\.?(.+)$', clause_id)
        if match:
            prefix = match.group(1) or ""
            numbers = match.group(2)

            # 將數字部分轉為 tuple
            parts = []
            for part in numbers.split('.'):
                try:
                    parts.append(int(part))
                except ValueError:
                    parts.append(0)

            return (prefix, tuple(parts))
        return ("", (0,))


def main():
    """主函數"""
    import argparse

    parser = argparse.ArgumentParser(description='解析 CB PDF')
    parser.add_argument('pdf', help='PDF 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出 JSON 路徑')
    parser.add_argument('--verbose', '-v', action='store_true', help='詳細輸出')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"錯誤: PDF 不存在: {pdf_path}")
        sys.exit(1)

    cb_parser = CBParserV2(pdf_path)
    result = cb_parser.parse()

    # 輸出摘要
    print(f"\n{'='*60}")
    print(f"解析結果摘要")
    print(f"{'='*60}")
    print(f"檔案: {result.filename}")
    print(f"報告編號: {result.test_report_no}")
    print(f"型號: {result.model}")
    print(f"Overview 項目: {len(result.overview_items)}")
    print(f"條款數: {len(result.clauses)}")

    # 統計 Verdict
    verdict_counts = {}
    for clause in result.clauses:
        v = clause.verdict or '(empty)'
        verdict_counts[v] = verdict_counts.get(v, 0) + 1

    print(f"\nVerdict 統計:")
    for v, count in sorted(verdict_counts.items()):
        print(f"  {v}: {count}")

    if args.verbose:
        print(f"\n前 10 個條款:")
        for clause in result.clauses[:10]:
            print(f"  {clause.clause_id}: {clause.requirement[:50]}... [{clause.verdict}]")

    # 儲存 JSON
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        print(f"\n已儲存: {output_path}")

    return result


if __name__ == '__main__':
    main()
