#!/usr/bin/env python3
"""
PDF vs DOC 比對腳本 v2

功能：
1. 使用改進版 PDF 解析器
2. 提取人工翻譯 DOC 的內容
3. 比對差異並計算相似度
"""
import sys
import re
import json
import subprocess
import logging
from pathlib import Path
from difflib import SequenceMatcher, unified_diff
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

# 添加專案路徑
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.parse_cb_pdf_v2 import CBParserV2, ParseResultV2, ClauseItem

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class ClauseComparison:
    """條款比對結果"""
    clause_id: str
    similarity: float
    pdf_requirement: str
    pdf_result: str
    pdf_verdict: str
    doc_content: str
    differences: List[str] = field(default_factory=list)


@dataclass
class ComparisonReport:
    """比對報告"""
    pdf_file: str
    doc_file: str

    # 統計
    total_pdf_clauses: int = 0
    total_doc_clauses: int = 0
    matched_clauses: int = 0
    only_in_pdf: List[str] = field(default_factory=list)
    only_in_doc: List[str] = field(default_factory=list)

    # 相似度
    overall_similarity: float = 0.0
    clause_similarities: List[ClauseComparison] = field(default_factory=list)

    # 低相似度條款 (需要優化)
    low_similarity_clauses: List[ClauseComparison] = field(default_factory=list)


def extract_doc_text(doc_path: Path) -> str:
    """從 .doc 檔案提取文字 (使用 LibreOffice 確保中文正確)"""
    doc_path = Path(doc_path)

    # 優先使用 libreoffice 轉換 (支援中文)
    try:
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            result = subprocess.run(
                ['soffice', '--headless', '--convert-to', 'txt:Text',
                 '--outdir', tmpdir, str(doc_path)],
                capture_output=True,
                timeout=120
            )
            if result.returncode == 0:
                txt_file = Path(tmpdir) / (doc_path.stem + '.txt')
                if txt_file.exists():
                    return txt_file.read_text(encoding='utf-8', errors='ignore')
    except Exception as e:
        logger.warning(f"libreoffice 轉換失敗: {e}")

    # 備用: antiword
    try:
        result = subprocess.run(
            ['antiword', str(doc_path)],
            capture_output=True,
            text=True,
            timeout=60
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout
    except (subprocess.SubprocessError, FileNotFoundError):
        pass

    raise RuntimeError(f"無法讀取 DOC 檔案: {doc_path}")


def parse_doc_clauses(doc_text: str) -> Dict[str, Dict]:
    """從 DOC 文字解析條款"""
    clauses = {}

    # 分割成行
    lines = doc_text.split('\n')

    # Clause ID 模式: 4.1.1, B.2.3, G.5.3.4 等
    clause_pattern = re.compile(r'^([A-Z]?\.?\d+(?:\.\d+)*)\s+(.*)', re.MULTILINE)

    current_clause_id = None
    current_content = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        match = clause_pattern.match(line)
        if match:
            # 儲存前一個條款
            if current_clause_id:
                clauses[current_clause_id] = {
                    'content': '\n'.join(current_content).strip(),
                    'raw_lines': current_content.copy()
                }

            current_clause_id = match.group(1)
            remaining = match.group(2).strip()
            current_content = [remaining] if remaining else []
        elif current_clause_id:
            current_content.append(line)

    # 儲存最後一個
    if current_clause_id:
        clauses[current_clause_id] = {
            'content': '\n'.join(current_content).strip(),
            'raw_lines': current_content.copy()
        }

    return clauses


def normalize_text(text: str) -> str:
    """正規化文字用於比對"""
    if not text:
        return ""

    # 移除多餘空白
    text = ' '.join(text.split())

    # 統一標點符號
    text = text.replace('–', '-').replace('—', '-')
    text = text.replace('"', '"').replace('"', '"')
    text = text.replace(''', "'").replace(''', "'")

    # 轉小寫
    text = text.lower()

    return text


def calculate_similarity(text1: str, text2: str) -> float:
    """計算文字相似度"""
    t1 = normalize_text(text1)
    t2 = normalize_text(text2)

    if not t1 and not t2:
        return 1.0
    if not t1 or not t2:
        return 0.0

    return SequenceMatcher(None, t1, t2).ratio()


def compare_clauses(
    pdf_clauses: List[ClauseItem],
    doc_clauses: Dict[str, Dict]
) -> Tuple[List[ClauseComparison], List[str], List[str]]:
    """比對條款"""
    comparisons = []
    pdf_ids = {c.clause_id for c in pdf_clauses}
    doc_ids = set(doc_clauses.keys())

    # 找出共同和獨有的條款
    only_in_pdf = sorted(pdf_ids - doc_ids)
    only_in_doc = sorted(doc_ids - pdf_ids)
    common_ids = pdf_ids & doc_ids

    # 比對共同條款
    for pdf_clause in pdf_clauses:
        if pdf_clause.clause_id not in common_ids:
            continue

        doc_data = doc_clauses.get(pdf_clause.clause_id, {})
        doc_content = doc_data.get('content', '')

        # 合併 PDF 內容
        pdf_content = f"{pdf_clause.requirement} {pdf_clause.result_remark}".strip()

        # 計算相似度
        similarity = calculate_similarity(pdf_content, doc_content)

        comp = ClauseComparison(
            clause_id=pdf_clause.clause_id,
            similarity=similarity,
            pdf_requirement=pdf_clause.requirement,
            pdf_result=pdf_clause.result_remark,
            pdf_verdict=pdf_clause.verdict,
            doc_content=doc_content
        )
        comparisons.append(comp)

    return comparisons, only_in_pdf, only_in_doc


def run_comparison(pdf_path: Path, doc_path: Path, output_dir: Path = None) -> ComparisonReport:
    """執行比對"""
    report = ComparisonReport(
        pdf_file=str(pdf_path),
        doc_file=str(doc_path)
    )

    logger.info(f"\n{'='*70}")
    logger.info(f"PDF vs DOC 比對")
    logger.info(f"{'='*70}")
    logger.info(f"PDF: {pdf_path.name}")
    logger.info(f"DOC: {doc_path.name}")

    # 1. 解析 PDF
    logger.info(f"\n[步驟 1] 解析 PDF...")
    try:
        parser = CBParserV2(pdf_path)
        pdf_result = parser.parse()
        report.total_pdf_clauses = len(pdf_result.clauses)
        logger.info(f"  ✓ 擷取 {report.total_pdf_clauses} 個條款")
    except Exception as e:
        logger.error(f"  ✗ PDF 解析失敗: {e}")
        return report

    # 2. 提取 DOC
    logger.info(f"\n[步驟 2] 提取 DOC 內容...")
    try:
        doc_text = extract_doc_text(doc_path)
        logger.info(f"  ✓ DOC 文字長度: {len(doc_text)} 字元")
    except Exception as e:
        logger.error(f"  ✗ DOC 提取失敗: {e}")
        return report

    # 3. 解析 DOC 條款
    logger.info(f"\n[步驟 3] 解析 DOC 條款...")
    doc_clauses = parse_doc_clauses(doc_text)
    report.total_doc_clauses = len(doc_clauses)
    logger.info(f"  ✓ 識別 {report.total_doc_clauses} 個條款")

    # 4. 比對
    logger.info(f"\n[步驟 4] 比對條款...")
    comparisons, only_pdf, only_doc = compare_clauses(pdf_result.clauses, doc_clauses)

    report.clause_similarities = comparisons
    report.only_in_pdf = only_pdf
    report.only_in_doc = only_doc
    report.matched_clauses = len(comparisons)

    # 計算整體相似度
    if comparisons:
        report.overall_similarity = sum(c.similarity for c in comparisons) / len(comparisons)

    # 找出低相似度條款 (< 0.5)
    report.low_similarity_clauses = [c for c in comparisons if c.similarity < 0.5]

    # 5. 輸出報告
    logger.info(f"\n{'='*70}")
    logger.info(f"比對結果")
    logger.info(f"{'='*70}")
    logger.info(f"PDF 條款數: {report.total_pdf_clauses}")
    logger.info(f"DOC 條款數: {report.total_doc_clauses}")
    logger.info(f"共同條款數: {report.matched_clauses}")
    logger.info(f"PDF 獨有: {len(report.only_in_pdf)}")
    logger.info(f"DOC 獨有: {len(report.only_in_doc)}")
    logger.info(f"\n整體相似度: {report.overall_similarity:.1%}")

    # 相似度分布
    if comparisons:
        high = sum(1 for c in comparisons if c.similarity >= 0.8)
        medium = sum(1 for c in comparisons if 0.5 <= c.similarity < 0.8)
        low = sum(1 for c in comparisons if c.similarity < 0.5)

        logger.info(f"\n相似度分布:")
        logger.info(f"  高 (>=80%): {high} ({high/len(comparisons)*100:.1f}%)")
        logger.info(f"  中 (50-79%): {medium} ({medium/len(comparisons)*100:.1f}%)")
        logger.info(f"  低 (<50%): {low} ({low/len(comparisons)*100:.1f}%)")

    # 顯示低相似度條款範例
    if report.low_similarity_clauses:
        logger.info(f"\n低相似度條款範例 (前 5 個):")
        for comp in sorted(report.low_similarity_clauses, key=lambda x: x.similarity)[:5]:
            logger.info(f"\n  {comp.clause_id} (相似度: {comp.similarity:.1%})")
            logger.info(f"    PDF: {comp.pdf_requirement[:80]}...")
            logger.info(f"    DOC: {comp.doc_content[:80]}...")

    # 儲存報告
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        report_path = output_dir / f"{pdf_path.stem}_comparison_v2.json"
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump({
                'pdf_file': report.pdf_file,
                'doc_file': report.doc_file,
                'total_pdf_clauses': report.total_pdf_clauses,
                'total_doc_clauses': report.total_doc_clauses,
                'matched_clauses': report.matched_clauses,
                'overall_similarity': report.overall_similarity,
                'only_in_pdf': report.only_in_pdf[:50],
                'only_in_doc': report.only_in_doc[:50],
                'low_similarity_count': len(report.low_similarity_clauses),
                'clause_similarities': [
                    {
                        'clause_id': c.clause_id,
                        'similarity': c.similarity,
                        'pdf_requirement': c.pdf_requirement[:200],
                        'doc_content': c.doc_content[:200]
                    }
                    for c in sorted(comparisons, key=lambda x: x.similarity)[:50]
                ]
            }, f, ensure_ascii=False, indent=2)
        logger.info(f"\n報告已儲存: {report_path}")

    return report


def main():
    """主函數"""
    import argparse

    parser = argparse.ArgumentParser(description='比對 PDF 與 DOC')
    parser.add_argument('--pdf', required=True, help='PDF 檔案')
    parser.add_argument('--doc', required=True, help='DOC 檔案')
    parser.add_argument('--output', '-o', default='output/comparison', help='輸出目錄')
    parser.add_argument('--target', '-t', type=float, default=0.9, help='目標相似度 (預設 0.9)')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    doc_path = Path(args.doc)

    if not pdf_path.exists():
        print(f"錯誤: PDF 不存在: {pdf_path}")
        sys.exit(1)
    if not doc_path.exists():
        print(f"錯誤: DOC 不存在: {doc_path}")
        sys.exit(1)

    report = run_comparison(pdf_path, doc_path, Path(args.output))

    # 檢查是否達標
    if report.overall_similarity >= args.target:
        print(f"\n✅ 相似度達標: {report.overall_similarity:.1%} >= {args.target:.0%}")
        sys.exit(0)
    else:
        print(f"\n❌ 相似度未達標: {report.overall_similarity:.1%} < {args.target:.0%}")
        sys.exit(1)


if __name__ == '__main__':
    main()
