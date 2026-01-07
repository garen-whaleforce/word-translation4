#!/usr/bin/env python3
"""
附表提取器 (Appended Tables Extractor)

從 CB PDF 提取附表 (TABLE: ...) 資料
"""
import re
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field, asdict

import pdfplumber

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class AppendedTable:
    """附表資料"""
    table_id: str  # 例如 "5.2", "M.3", "B.2.5"
    title: str  # 完整標題
    page_number: int
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    verdict: str = ""  # 該表的整體判定 (P/N/A/Fail)


@dataclass
class AppendedTablesResult:
    """附表提取結果"""
    filename: str
    tables: Dict[str, AppendedTable] = field(default_factory=dict)
    errors: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "filename": self.filename,
            "table_count": len(self.tables),
            "tables": {k: asdict(v) for k, v in self.tables.items()},
            "table_ids": list(self.tables.keys()),
            "errors": self.errors,
        }


class AppendedTablesExtractor:
    """附表提取器"""

    # TABLE: 模式
    TABLE_PATTERN = re.compile(
        r"(?:TABLE|Table):\s*([^\n]+)",
        re.IGNORECASE
    )

    # 常見附表 ID 對映 (PDF 標題 -> Word 表格 ID)
    TABLE_ID_MAPPING = {
        # 5.x 系列
        "Classification of electrical energy sources": "5.2",
        "classification of electrical energy sources": "5.2",
        "Electric strength tests": "5.4.1.8",
        "Minimum clearances": "5.4.1.10.2",
        "Minimum creepage distances": "5.4.1.10.3",
        "Abnormal operating and fault condition tests": "5.4.2, 5.4.3",
        "Ball pressure test": "5.4.4.2",
        "Vicat softening temperature": "5.4.4.9",
        "Temperature measuring": "5.4.9",
        "Routing test": "5.5.2.2",
        "Capacitor discharge": "5.6.6",
        "Drop test": "5.7.4",
        "Impact test": "5.7.5",
        "Durability of markings": "5.8",

        # 6.x 系列
        "Determination of resistive PIS": "6.2.2",
        "Determination of Arcing PIS": "6.2.3.1",
        "Transformer and inductor PIS determination": "6.2.3.2",
        "Safeguards for flammable material": "6.4.8.3.3, 6.4.8.3.4, 6.4.8.3.5, P.2.2, P.2.3",

        # 8.x 系列
        "Equipment with electromechanical means for destruction": "8.5.5",

        # 9.x 系列
        "Temperature measurements": "9.6",
        "Temperature test results": "9.6",

        # B 系列
        "Critical components information": "5.4.1.4, 9.3, B.1.5, B.2.6",
        "Output voltage on USB ports": "B.2.5",
        "USB output voltage": "B.2.5",
        "Circuits intended for interconnection": "B.3, B.4",

        # M 系列
        "Charging safeguards for equipment containing": "M.3",
        "Battery charging tests": "M.4.2",

        # Q 系列
        "Backfeed safeguard": "Q.1",

        # T 系列
        "Voltage surge test": "T.2, T.3, T.4, T.5",
        "Mains transient test": "T.6, T.9",
        "Circuit design analysis": "T.7",
        "High pressure lamp": "T.8",

        # X 系列
        "Alternative method for determining minimum clearances": "X",

        # 4.1.2 系列
        "Equipment for direct insertion into mains": "4.1.2",
        "Direct plug-in": "4.1.2",
        "plug dimensions": "4.1.2",
    }

    # Verdict 模式
    VERDICT_PATTERN = re.compile(r'\b(P|N/?A|Fail|F)\b', re.IGNORECASE)

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        self.result = AppendedTablesResult(filename=self.pdf_path.name)

    def extract(self) -> AppendedTablesResult:
        """提取附表"""
        logger.info(f"提取附表: {self.pdf_path.name}")

        with pdfplumber.open(self.pdf_path) as pdf:
            current_table_info = None

            for page_num, page in enumerate(pdf.pages):
                page_number = page_num + 1
                text = page.extract_text() or ""

                # 找 TABLE: 標題
                matches = self.TABLE_PATTERN.findall(text)

                for title in matches:
                    title = title.strip()
                    if not title:
                        continue

                    # 提取 verdict
                    verdict_match = self.VERDICT_PATTERN.search(title)
                    verdict = verdict_match.group(1).upper() if verdict_match else ""
                    if verdict == "N/A" or verdict == "NA":
                        verdict = "N/A"

                    # 清理標題
                    clean_title = self.VERDICT_PATTERN.sub("", title).strip()
                    clean_title = re.sub(r'\s+', ' ', clean_title)

                    # 對映到 Word 表格 ID
                    table_id = self._match_table_id(clean_title)

                    if table_id:
                        # 提取表格資料
                        table_data = self._extract_table_near_title(page, title)

                        appended_table = AppendedTable(
                            table_id=table_id,
                            title=clean_title,
                            page_number=page_number,
                            verdict=verdict,
                            headers=table_data.get("headers", []),
                            rows=table_data.get("rows", [])
                        )

                        # 避免重複 (保留第一個)
                        if table_id not in self.result.tables:
                            self.result.tables[table_id] = appended_table

        logger.info(f"提取完成: {len(self.result.tables)} 個附表")
        return self.result

    def _match_table_id(self, title: str) -> Optional[str]:
        """根據標題對映 Word 表格 ID"""
        title_lower = title.lower()

        # 直接匹配
        for pattern, table_id in self.TABLE_ID_MAPPING.items():
            if pattern.lower() in title_lower:
                return table_id

        # 嘗試從標題提取數字編號
        # 例如 "5.4.1.10.3 Minimum creepage distances" -> "5.4.1.10.3"
        match = re.match(r'^([\d\.]+)\s', title)
        if match:
            return match.group(1)

        # 嘗試提取附錄字母編號
        # 例如 "M.3 Charging safeguards" -> "M.3"
        match = re.match(r'^([A-Z]\.\d+)', title)
        if match:
            return match.group(1)

        return None

    def _extract_table_near_title(self, page, title: str) -> Dict[str, Any]:
        """提取標題附近的表格資料"""
        result = {"headers": [], "rows": []}

        try:
            tables = page.extract_tables()
            if not tables:
                return result

            # 找最接近標題的表格
            # (簡化實作: 取頁面上的第一個表格)
            for table in tables:
                if not table or len(table) < 2:
                    continue

                # 跳過標題行是 "Clause" 開頭的 (那是條款表)
                first_row = table[0]
                if first_row and any(
                    str(c).strip().startswith("Clause") for c in first_row if c
                ):
                    continue

                # 取得 headers
                headers = [str(c).strip() if c else "" for c in table[0]]
                result["headers"] = headers

                # 取得 rows
                for row in table[1:]:
                    row_data = [str(c).strip() if c else "" for c in row]
                    if any(row_data):  # 跳過全空行
                        result["rows"].append(row_data)

                # 只取第一個合適的表格
                if result["rows"]:
                    break

        except Exception as e:
            logger.debug(f"提取表格失敗: {e}")

        return result


def extract_appended_tables(pdf_path: Path) -> AppendedTablesResult:
    """
    便捷函數：提取附表

    Args:
        pdf_path: PDF 路徑

    Returns:
        AppendedTablesResult
    """
    extractor = AppendedTablesExtractor(pdf_path)
    return extractor.extract()


def main():
    import argparse
    import json

    parser = argparse.ArgumentParser(description='提取 PDF 附表')
    parser.add_argument('pdf', help='PDF 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出 JSON 路徑')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    result = extract_appended_tables(pdf_path)

    print(f"\n{'='*60}")
    print("附表提取結果")
    print(f"{'='*60}")
    print(f"附表數: {len(result.tables)}")
    print(f"\n附表列表:")
    for table_id, table in sorted(result.tables.items()):
        print(f"  [{table_id}] {table.title[:40]}... (第{table.page_number}頁, {table.verdict})")
        if table.rows:
            print(f"       資料行: {len(table.rows)}")

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
        print(f"\n已儲存: {output_path}")


if __name__ == '__main__':
    main()
