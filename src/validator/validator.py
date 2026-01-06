"""
Validator Module - PDF vs Word 一致性驗證

功能:
1. 比較 PDF 擷取的 clauses 與 Word 內的 clauses
2. 檢查 verdict 一致性
3. 檢查術語違規
4. 生成驗證報告
"""
from __future__ import annotations
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Set
from dataclasses import dataclass, field

from docx import Document

from ..cb_parser import ParseResult, ClauseItem

logger = logging.getLogger(__name__)


@dataclass
class ClauseMismatch:
    """條款不一致"""
    clause_id: str
    issue_type: str  # missing, duplicate, order_mismatch, verdict_mismatch
    expected: str = ""
    actual: str = ""


@dataclass
class ValidationReport:
    """驗證報告"""
    # 摘要
    total_pdf_clauses: int = 0
    total_word_clauses: int = 0
    matching_clauses: int = 0

    # 問題
    missing_clauses: List[str] = field(default_factory=list)
    extra_clauses: List[str] = field(default_factory=list)
    duplicate_clauses: List[str] = field(default_factory=list)
    verdict_mismatches: List[ClauseMismatch] = field(default_factory=list)
    order_issues: List[str] = field(default_factory=list)
    term_violations: List[Dict[str, Any]] = field(default_factory=list)

    # 狀態
    status: str = "unknown"  # passed, failed, warning

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status,
            "summary": {
                "total_pdf_clauses": self.total_pdf_clauses,
                "total_word_clauses": self.total_word_clauses,
                "matching_clauses": self.matching_clauses,
                "coverage": self.matching_clauses / self.total_pdf_clauses if self.total_pdf_clauses > 0 else 0
            },
            "issues": {
                "missing_clauses": self.missing_clauses,
                "extra_clauses": self.extra_clauses,
                "duplicate_clauses": self.duplicate_clauses,
                "verdict_mismatches": [
                    {"clause_id": m.clause_id, "expected": m.expected, "actual": m.actual}
                    for m in self.verdict_mismatches
                ],
                "order_issues": self.order_issues,
                "term_violations": self.term_violations
            },
            "total_issues": self._count_issues()
        }

    def _count_issues(self) -> int:
        return (
            len(self.missing_clauses) +
            len(self.extra_clauses) +
            len(self.duplicate_clauses) +
            len(self.verdict_mismatches) +
            len(self.order_issues) +
            len(self.term_violations)
        )


class Validator:
    """PDF vs Word 驗證器"""

    def __init__(self):
        self.report = ValidationReport()

    def validate(
        self,
        pdf_result: ParseResult,
        word_path: Optional[Path] = None,
        word_clauses: Optional[List[ClauseItem]] = None
    ) -> ValidationReport:
        """
        執行驗證

        Args:
            pdf_result: PDF 解析結果
            word_path: Word 檔案路徑 (可選)
            word_clauses: Word 中的條款列表 (可選，如果沒有 word_path)

        Returns:
            ValidationReport
        """
        # 取得 PDF clauses
        pdf_clauses = pdf_result.clauses
        self.report.total_pdf_clauses = len(pdf_clauses)

        # 取得 Word clauses
        if word_path:
            word_clauses = self._extract_word_clauses(word_path)
        elif word_clauses is None:
            word_clauses = []

        self.report.total_word_clauses = len(word_clauses)

        # 執行比較
        self._compare_clauses(pdf_clauses, word_clauses)
        self._check_verdicts(pdf_clauses, word_clauses)
        self._check_order(pdf_clauses, word_clauses)

        # 決定狀態
        self._determine_status()

        return self.report

    def _extract_word_clauses(self, word_path: Path) -> List[ClauseItem]:
        """從 Word 檔案擷取條款"""
        clauses = []

        try:
            doc = Document(word_path)

            for table in doc.tables:
                if not table.rows:
                    continue

                # 檢查是否為 clause 表
                header = " ".join(cell.text for cell in table.rows[0].cells).upper()
                if "CLAUSE" not in header and "條款" not in header:
                    continue

                # 擷取條款
                for row in table.rows[1:]:
                    cells = row.cells
                    if len(cells) >= 2:
                        clause_id = cells[0].text.strip()
                        verdict = cells[-1].text.strip() if len(cells) >= 4 else ""

                        if clause_id:
                            clauses.append(ClauseItem(
                                clause_id=clause_id,
                                requirement_test="",
                                result_remark="",
                                verdict=verdict
                            ))

        except Exception as e:
            logger.error(f"Failed to extract Word clauses: {e}")

        return clauses

    def _compare_clauses(
        self,
        pdf_clauses: List[ClauseItem],
        word_clauses: List[ClauseItem]
    ):
        """比較條款覆蓋率"""
        pdf_ids = {c.clause_id for c in pdf_clauses}
        word_ids = {c.clause_id for c in word_clauses}

        # 缺失的條款
        self.report.missing_clauses = list(pdf_ids - word_ids)

        # 多餘的條款
        self.report.extra_clauses = list(word_ids - pdf_ids)

        # 匹配的條款
        self.report.matching_clauses = len(pdf_ids & word_ids)

        # 檢查重複
        seen = set()
        for clause in word_clauses:
            if clause.clause_id in seen:
                self.report.duplicate_clauses.append(clause.clause_id)
            seen.add(clause.clause_id)

    def _check_verdicts(
        self,
        pdf_clauses: List[ClauseItem],
        word_clauses: List[ClauseItem]
    ):
        """檢查 verdict 一致性"""
        pdf_verdicts = {c.clause_id: c.verdict for c in pdf_clauses}
        word_verdicts = {c.clause_id: c.verdict for c in word_clauses}

        for clause_id, pdf_verdict in pdf_verdicts.items():
            word_verdict = word_verdicts.get(clause_id)
            if word_verdict and pdf_verdict and pdf_verdict != word_verdict:
                self.report.verdict_mismatches.append(ClauseMismatch(
                    clause_id=clause_id,
                    issue_type="verdict_mismatch",
                    expected=pdf_verdict,
                    actual=word_verdict
                ))

    def _check_order(
        self,
        pdf_clauses: List[ClauseItem],
        word_clauses: List[ClauseItem]
    ):
        """檢查條款順序"""
        pdf_order = [c.clause_id for c in pdf_clauses]
        word_order = [c.clause_id for c in word_clauses]

        # 只比較共同的條款
        common = set(pdf_order) & set(word_order)
        pdf_filtered = [c for c in pdf_order if c in common]
        word_filtered = [c for c in word_order if c in common]

        if pdf_filtered != word_filtered:
            # 找出順序不一致的位置
            for i, (p, w) in enumerate(zip(pdf_filtered, word_filtered)):
                if p != w:
                    self.report.order_issues.append(
                        f"Position {i}: expected {p}, got {w}"
                    )
                    if len(self.report.order_issues) >= 5:  # 限制報告數量
                        break

    def _determine_status(self):
        """決定驗證狀態"""
        issues = self.report._count_issues()

        if issues == 0:
            self.report.status = "passed"
        elif self.report.missing_clauses or self.report.verdict_mismatches:
            self.report.status = "failed"
        else:
            self.report.status = "warning"


def validate_output(
    pdf_result: ParseResult,
    word_path: Optional[Path] = None
) -> ValidationReport:
    """
    便捷函數：驗證輸出一致性

    Args:
        pdf_result: PDF 解析結果
        word_path: Word 輸出路徑

    Returns:
        ValidationReport
    """
    validator = Validator()
    return validator.validate(pdf_result, word_path)
