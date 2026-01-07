#!/usr/bin/env python3
"""
比對自動生成 Word 與人工翻譯 DOC

功能:
1. 提取兩份文件的表格內容
2. 比對條款數量、章節結構
3. 比對 Verdict 一致性
4. 計算內容相似度
5. 生成差異報告
"""
import sys
import re
import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional
from dataclasses import dataclass, field, asdict
from difflib import SequenceMatcher

from docx import Document

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class TableInfo:
    """表格資訊"""
    table_index: int
    first_cell: str
    row_count: int
    col_count: int
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class ClauseInfo:
    """條款資訊"""
    clause_id: str
    requirement: str
    result_remark: str
    verdict: str
    table_index: int


@dataclass
class CompareResult:
    """比對結果"""
    auto_file: str
    human_file: str

    # 統計
    auto_tables: int = 0
    human_tables: int = 0
    auto_clauses: int = 0
    human_clauses: int = 0

    # 匹配
    matched_clauses: int = 0
    verdict_matches: int = 0
    verdict_mismatches: int = 0

    # 相似度
    overall_similarity: float = 0.0
    structure_similarity: float = 0.0
    content_similarity: float = 0.0

    # 差異項目
    missing_in_auto: List[str] = field(default_factory=list)
    missing_in_human: List[str] = field(default_factory=list)
    verdict_diffs: List[Dict[str, str]] = field(default_factory=list)
    content_diffs: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def extract_tables_from_docx(doc_path: Path) -> List[TableInfo]:
    """從 DOCX 提取表格"""
    doc = Document(doc_path)
    tables = []

    for idx, table in enumerate(doc.tables):
        if not table.rows:
            continue

        # 取得第一個 cell
        first_cell = ""
        try:
            first_cell = table.cell(0, 0).text.strip()
            if '\n' in first_cell:
                first_cell = first_cell.split('\n')[0].strip()
        except:
            pass

        # 提取所有行
        rows = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                text = cell.text.strip()
                # 清理多餘空白
                text = ' '.join(text.split())
                row_data.append(text)
            rows.append(row_data)

        tables.append(TableInfo(
            table_index=idx,
            first_cell=first_cell,
            row_count=len(table.rows),
            col_count=len(table.rows[0].cells) if table.rows else 0,
            rows=rows
        ))

    return tables


def extract_clauses_from_tables(tables: List[TableInfo]) -> List[ClauseInfo]:
    """從表格提取條款"""
    clauses = []

    # Clause ID 模式
    clause_pattern = re.compile(r'^([A-Z]\.)?(\d+)(\.\d+)*$')

    # Verdict 值
    verdict_values = {'P', 'N/A', 'NA', 'N.A.', 'Fail', 'F', '--', '-', '符合', '不適用', '不符合'}

    for table in tables:
        for row in table.rows:
            if not row or len(row) < 2:
                continue

            first_cell = row[0].strip()

            # 檢查是否為 Clause ID
            if clause_pattern.match(first_cell):
                # 取得各欄位
                requirement = row[1] if len(row) > 1 else ""
                result_remark = row[2] if len(row) > 2 else ""
                verdict = row[-1] if len(row) > 1 else ""

                # 標準化 verdict
                verdict_upper = verdict.upper().strip()
                if verdict_upper in {'P', '符合'}:
                    verdict = 'P'
                elif verdict_upper in {'N/A', 'NA', 'N.A.', '不適用'}:
                    verdict = 'N/A'
                elif verdict_upper in {'F', 'FAIL', '不符合'}:
                    verdict = 'Fail'
                elif verdict_upper in {'--', '-', '⎯', '—'}:
                    verdict = '--'

                clauses.append(ClauseInfo(
                    clause_id=first_cell,
                    requirement=requirement,
                    result_remark=result_remark,
                    verdict=verdict,
                    table_index=table.table_index
                ))

    return clauses


def calculate_text_similarity(text1: str, text2: str) -> float:
    """計算文字相似度"""
    if not text1 and not text2:
        return 1.0
    if not text1 or not text2:
        return 0.0

    # 正規化文字
    text1 = ' '.join(text1.lower().split())
    text2 = ' '.join(text2.lower().split())

    return SequenceMatcher(None, text1, text2).ratio()


def compare_documents(auto_path: Path, human_path: Path) -> CompareResult:
    """比對兩份文件"""
    result = CompareResult(
        auto_file=str(auto_path),
        human_file=str(human_path)
    )

    # 提取表格
    logger.info("提取自動生成文件表格...")
    auto_tables = extract_tables_from_docx(auto_path)
    result.auto_tables = len(auto_tables)

    logger.info("提取人工翻譯文件表格...")
    human_tables = extract_tables_from_docx(human_path)
    result.human_tables = len(human_tables)

    # 提取條款
    logger.info("提取條款...")
    auto_clauses = extract_clauses_from_tables(auto_tables)
    human_clauses = extract_clauses_from_tables(human_tables)

    result.auto_clauses = len(auto_clauses)
    result.human_clauses = len(human_clauses)

    # 建立條款索引
    auto_clause_map = {c.clause_id: c for c in auto_clauses}
    human_clause_map = {c.clause_id: c for c in human_clauses}

    # 找出共同條款
    common_ids = set(auto_clause_map.keys()) & set(human_clause_map.keys())
    result.matched_clauses = len(common_ids)

    # 找出差異
    result.missing_in_auto = sorted(set(human_clause_map.keys()) - set(auto_clause_map.keys()))
    result.missing_in_human = sorted(set(auto_clause_map.keys()) - set(human_clause_map.keys()))

    # 比對共同條款
    total_content_similarity = 0.0

    for clause_id in sorted(common_ids):
        auto_clause = auto_clause_map[clause_id]
        human_clause = human_clause_map[clause_id]

        # 比對 Verdict
        if auto_clause.verdict == human_clause.verdict:
            result.verdict_matches += 1
        else:
            result.verdict_mismatches += 1
            result.verdict_diffs.append({
                "clause_id": clause_id,
                "auto_verdict": auto_clause.verdict,
                "human_verdict": human_clause.verdict
            })

        # 計算內容相似度
        req_sim = calculate_text_similarity(auto_clause.requirement, human_clause.requirement)
        total_content_similarity += req_sim

        # 如果相似度低於閾值，記錄差異
        if req_sim < 0.8:
            result.content_diffs.append({
                "clause_id": clause_id,
                "similarity": round(req_sim, 3),
                "auto": auto_clause.requirement[:100] + "..." if len(auto_clause.requirement) > 100 else auto_clause.requirement,
                "human": human_clause.requirement[:100] + "..." if len(human_clause.requirement) > 100 else human_clause.requirement
            })

    # 計算整體相似度
    if common_ids:
        result.content_similarity = total_content_similarity / len(common_ids)

    # 結構相似度 (基於表格數量和條款匹配率)
    if result.human_clauses > 0:
        result.structure_similarity = result.matched_clauses / result.human_clauses

    # Verdict 一致性
    if result.matched_clauses > 0:
        verdict_accuracy = result.verdict_matches / result.matched_clauses
    else:
        verdict_accuracy = 0.0

    # 整體相似度 (加權平均)
    result.overall_similarity = (
        result.structure_similarity * 0.3 +
        result.content_similarity * 0.4 +
        verdict_accuracy * 0.3
    )

    return result


def main():
    import argparse

    parser = argparse.ArgumentParser(description='比對自動生成與人工翻譯文件')
    parser.add_argument('--auto', required=True, help='自動生成的 DOCX')
    parser.add_argument('--human', required=True, help='人工翻譯的 DOC/DOCX')
    parser.add_argument('--output', '-o', help='輸出 JSON 報告')

    args = parser.parse_args()

    auto_path = Path(args.auto)
    human_path = Path(args.human)

    if not auto_path.exists():
        print(f"錯誤: 找不到自動生成文件: {auto_path}")
        sys.exit(1)

    if not human_path.exists():
        print(f"錯誤: 找不到人工翻譯文件: {human_path}")
        sys.exit(1)

    # 比對
    result = compare_documents(auto_path, human_path)

    # 輸出報告
    print(f"\n{'='*60}")
    print("文件比對報告")
    print(f"{'='*60}")
    print(f"自動生成: {result.auto_file}")
    print(f"人工翻譯: {result.human_file}")
    print()

    print("統計:")
    print(f"  自動生成表格數: {result.auto_tables}")
    print(f"  人工翻譯表格數: {result.human_tables}")
    print(f"  自動生成條款數: {result.auto_clauses}")
    print(f"  人工翻譯條款數: {result.human_clauses}")
    print(f"  匹配條款數: {result.matched_clauses}")
    print()

    print("相似度:")
    print(f"  整體相似度: {result.overall_similarity*100:.1f}%")
    print(f"  結構相似度: {result.structure_similarity*100:.1f}%")
    print(f"  內容相似度: {result.content_similarity*100:.1f}%")
    print()

    print("Verdict 比對:")
    print(f"  一致: {result.verdict_matches}")
    print(f"  不一致: {result.verdict_mismatches}")
    if result.verdict_matches + result.verdict_mismatches > 0:
        accuracy = result.verdict_matches / (result.verdict_matches + result.verdict_mismatches)
        print(f"  準確率: {accuracy*100:.1f}%")
    print()

    if result.missing_in_auto:
        print(f"自動生成缺少 ({len(result.missing_in_auto)} 項):")
        for cid in result.missing_in_auto[:10]:
            print(f"  - {cid}")
        if len(result.missing_in_auto) > 10:
            print(f"  ... 還有 {len(result.missing_in_auto) - 10} 項")
        print()

    if result.missing_in_human:
        print(f"人工翻譯缺少 ({len(result.missing_in_human)} 項):")
        for cid in result.missing_in_human[:10]:
            print(f"  - {cid}")
        if len(result.missing_in_human) > 10:
            print(f"  ... 還有 {len(result.missing_in_human) - 10} 項")
        print()

    if result.verdict_diffs:
        print(f"Verdict 差異 ({len(result.verdict_diffs)} 項):")
        for diff in result.verdict_diffs[:10]:
            print(f"  {diff['clause_id']}: 自動={diff['auto_verdict']} vs 人工={diff['human_verdict']}")
        if len(result.verdict_diffs) > 10:
            print(f"  ... 還有 {len(result.verdict_diffs) - 10} 項")
        print()

    if result.content_diffs:
        print(f"內容差異較大 ({len(result.content_diffs)} 項, 相似度<80%):")
        for diff in result.content_diffs[:5]:
            print(f"  {diff['clause_id']} (相似度: {diff['similarity']*100:.1f}%)")
            print(f"    自動: {diff['auto'][:60]}...")
            print(f"    人工: {diff['human'][:60]}...")
        if len(result.content_diffs) > 5:
            print(f"  ... 還有 {len(result.content_diffs) - 5} 項")

    # 儲存 JSON
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        print(f"\n已儲存報告: {output_path}")


if __name__ == '__main__':
    main()
