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


class CBParserV2:
    """CB PDF 解析器 v2"""

    # Clause ID 模式
    # 必須有子章節 (如 4.1, 5.2.3) 或字母前綴 (如 B.1, G.5)
    # 純數字 (4, 5, 6) 是章節標題，不是條款
    CLAUSE_ID_PATTERN = re.compile(r'^([A-Z]\.\d+(\.\d+)*|\d+\.\d+(\.\d+)*)$')

    # Verdict 模式
    # 包含 WingDings dash 符號 \uf0be
    VERDICT_VALUES = {'P', 'N/A', 'NA', 'N.A.', 'Fail', 'F', '--', '-', '⎯', '—', '\uf0be'}

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
        seen_clauses = set()

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

                    # Clause ID 格式: 4, 4.1, 4.1.1, B.1, G.5.3.4.2 等
                    if self.CLAUSE_ID_PATTERN.match(first_cell):
                        # 檢查最後一欄是否為 Verdict
                        last_cell = str(row[-1]).strip() if row[-1] else ""

                        # 如果最後一欄是 Verdict
                        if self._is_verdict(last_cell) or last_cell == '':
                            clause_id = first_cell
                            verdict = self._normalize_verdict(last_cell)

                            # 如果已存在且新 verdict 非空，更新 verdict
                            if clause_id in seen_clauses:
                                if verdict and verdict not in ('', '--'):
                                    # 找到現有條款並更新
                                    for c in self.result.clauses:
                                        if c.clause_id == clause_id and c.verdict in ('', '--'):
                                            c.verdict = verdict
                                            break
                                continue
                            seen_clauses.add(clause_id)

                            # 解析其他欄位
                            requirement = ""
                            result_remark = ""

                            if len(row) >= 4:
                                requirement = str(row[1]).strip() if row[1] else ""
                                result_remark = str(row[2]).strip() if row[2] else ""
                            elif len(row) >= 3:
                                requirement = str(row[1]).strip() if row[1] else ""
                                # 判斷第三欄是結果還是 verdict
                                third = str(row[2]).strip() if row[2] else ""
                                if self._is_verdict(third):
                                    verdict = self._normalize_verdict(third)
                                else:
                                    result_remark = third
                            elif len(row) >= 2:
                                requirement = str(row[1]).strip() if row[1] else ""

                            clause = ClauseItem(
                                clause_id=clause_id,
                                requirement=requirement,
                                result_remark=result_remark,
                                verdict=verdict,
                                page_number=page_num + 1
                            )
                            self.result.clauses.append(clause)

        # 依照 Clause ID 排序
        self.result.clauses.sort(key=lambda c: self._clause_sort_key(c.clause_id))

        # 按章節分群
        self._group_clauses_by_chapter()

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
        # 先檢查是否有字母前綴 (B.1, G.5, M.3, T.7 等)
        match = re.match(r'^([A-Z])\.', clause_id)
        if match:
            return match.group(1)

        # 純數字條款 (4.1.1, 5.2.3 等)
        match = re.match(r'^(\d+)', clause_id)
        if match:
            return match.group(1)

        return "other"

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
