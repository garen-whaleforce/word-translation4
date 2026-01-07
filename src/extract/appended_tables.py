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
    columns: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    verdict: str = ""  # 該表的整體判定 (P/N/A/Fail)
    supplementary_info: List[str] = field(default_factory=list)


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


@dataclass
class TableHeader:
    """附表標題行資訊"""
    table_id: Optional[str]
    title: str
    verdict: str
    page_number: int
    top: float
    bottom: float


@dataclass
class LineInfo:
    """文字行資訊"""
    text: str
    top: float
    bottom: float
    words: List[Dict[str, Any]]


class AppendedTablesExtractor:
    """附表提取器"""

    # TABLE: 標題行模式
    TABLE_HEADER_LINE_PATTERN = re.compile(r'\bTABLE\s*:', re.IGNORECASE)
    TABLE_HEADER_WITH_ID_PATTERN = re.compile(
        r'^(?P<table_id>(?:[A-Z]\.)?\d+(?:\.\d+)*|[A-Z])\s+TABLE\s*:\s*(?P<title>.+)$',
        re.IGNORECASE
    )

    # 座標分欄參數
    LINE_MERGE_TOLERANCE = 3.0
    HEADER_WORD_GAP = 6.0
    HEADER_MIN_WIDTH_RATIO = 0.3
    SUPPLEMENTARY_PATTERN = re.compile(
        r'^(Supplementary\s+information|Note|Remarks?)\b[:\s]',
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
            words_by_page = [
                page.extract_words(use_text_flow=True, keep_blank_chars=False)
                for page in pdf.pages
            ]
            headers = self._find_table_headers(pdf, words_by_page)

            for idx, header in enumerate(headers):
                next_header = headers[idx + 1] if idx + 1 < len(headers) else None
                table_id = header.table_id or self._match_table_id(header.title) or header.title

                columns, rows, supplementary_info = self._extract_table_segment(
                    pdf,
                    words_by_page,
                    header,
                    next_header
                )

                if not rows and not columns:
                    fallback = self._extract_table_near_title(
                        pdf.pages[header.page_number - 1],
                        header.title
                    )
                    columns = fallback.get("headers", [])
                    rows = fallback.get("rows", [])

                appended_table = AppendedTable(
                    table_id=table_id,
                    title=header.title,
                    page_number=header.page_number,
                    verdict=header.verdict,
                    headers=columns,
                    columns=columns,
                    rows=rows,
                    supplementary_info=supplementary_info,
                )

                if table_id not in self.result.tables:
                    self.result.tables[table_id] = appended_table

        logger.info(f"提取完成: {len(self.result.tables)} 個附表")
        return self.result

    def _find_table_headers(
        self,
        pdf: pdfplumber.PDF,
        words_by_page: List[List[Dict[str, Any]]]
    ) -> List[TableHeader]:
        headers: List[TableHeader] = []

        for page_num, page in enumerate(pdf.pages):
            words = words_by_page[page_num]
            lines = self._build_line_infos(words, tolerance=self.LINE_MERGE_TOLERANCE)
            for line in lines:
                if not self.TABLE_HEADER_LINE_PATTERN.search(line.text):
                    continue

                table_id, title, verdict = self._parse_table_header_line(line.text)
                if not title:
                    continue

                headers.append(TableHeader(
                    table_id=table_id,
                    title=title,
                    verdict=verdict,
                    page_number=page_num + 1,
                    top=line.top,
                    bottom=line.bottom,
                ))

        headers.sort(key=lambda h: (h.page_number, h.top))
        return headers

    def _parse_table_header_line(self, line_text: str) -> Tuple[Optional[str], str, str]:
        """解析 TABLE: 標題行"""
        verdict = ""
        verdict_match = self.VERDICT_PATTERN.search(line_text)
        if verdict_match:
            verdict = self._normalize_verdict(verdict_match.group(1))

        clean_text = self.VERDICT_PATTERN.sub("", line_text).strip()
        clean_text = re.sub(r'\s+', ' ', clean_text)

        table_id = None
        title = ""
        match = self.TABLE_HEADER_WITH_ID_PATTERN.match(clean_text)
        if match:
            table_id = match.group("table_id").strip()
            title = match.group("title").strip()
        else:
            parts = self.TABLE_HEADER_LINE_PATTERN.split(clean_text, 1)
            if len(parts) == 2:
                title = parts[1].strip()
            else:
                title = clean_text

        title = re.sub(r'\s+', ' ', title).strip()
        return table_id, title, verdict

    def _extract_table_segment(
        self,
        pdf: pdfplumber.PDF,
        words_by_page: List[List[Dict[str, Any]]],
        header: TableHeader,
        next_header: Optional[TableHeader],
    ) -> Tuple[List[str], List[List[str]], List[str]]:
        """以 TABLE 標題切段，抽取欄位與資料列"""
        start_page_idx = header.page_number - 1
        end_page_idx = next_header.page_number - 1 if next_header else len(pdf.pages) - 1

        segment_words: List[Dict[str, Any]] = []
        for page_idx in range(start_page_idx, end_page_idx + 1):
            for word in words_by_page[page_idx]:
                top = word.get("top", 0)
                if page_idx == start_page_idx and top <= header.bottom + 1:
                    continue
                if next_header and page_idx == end_page_idx and top >= next_header.top - 1:
                    continue
                segment_words.append(word)

        if not segment_words:
            return [], [], []

        page_width = pdf.pages[start_page_idx].width or 1
        lines = self._build_line_infos(segment_words, tolerance=self.LINE_MERGE_TOLERANCE)
        header_index = self._find_header_line_index(lines, page_width)
        if header_index is None:
            return [], [], []

        header_line = lines[header_index]
        columns, column_bounds = self._build_columns_from_header(header_line.words, page_width)
        if not columns:
            return [], [], []

        data_lines = lines[header_index + 1:]
        rows, supplementary_info = self._build_rows_from_lines(data_lines, column_bounds)

        return columns, rows, supplementary_info

    def _build_line_infos(self, words: List[Dict[str, Any]], tolerance: float) -> List[LineInfo]:
        lines: List[LineInfo] = []
        if not words:
            return lines

        sorted_words = sorted(words, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
        current: List[Dict[str, Any]] = []
        current_top: Optional[float] = None

        for word in sorted_words:
            top = word.get("top", 0)
            if current_top is None or abs(top - current_top) <= tolerance:
                current.append(word)
                if current_top is None:
                    current_top = top
            else:
                lines.append(self._line_from_words(current))
                current = [word]
                current_top = top

        if current:
            lines.append(self._line_from_words(current))

        return lines

    def _line_from_words(self, words: List[Dict[str, Any]]) -> LineInfo:
        sorted_words = sorted(words, key=lambda w: w.get("x0", 0))
        text = " ".join(w.get("text", "") for w in sorted_words).strip()
        top = min(w.get("top", 0) for w in words)
        bottom = max(w.get("bottom", 0) for w in words)
        return LineInfo(text=text, top=top, bottom=bottom, words=sorted_words)

    def _find_header_line_index(self, lines: List[LineInfo], page_width: float) -> Optional[int]:
        for idx, line in enumerate(lines):
            if not line.text:
                continue
            if self.TABLE_HEADER_LINE_PATTERN.search(line.text):
                continue
            if len(line.words) < 2:
                continue
            min_x = min(w.get("x0", 0) for w in line.words)
            max_x = max(w.get("x1", 0) for w in line.words)
            if (max_x - min_x) < page_width * self.HEADER_MIN_WIDTH_RATIO:
                continue
            return idx
        return None

    def _build_columns_from_header(
        self,
        header_words: List[Dict[str, Any]],
        page_width: float
    ) -> Tuple[List[str], List[Tuple[float, float]]]:
        if not header_words:
            return [], []

        sorted_words = sorted(header_words, key=lambda w: w.get("x0", 0))
        grouped = []
        current = None

        for word in sorted_words:
            text = word.get("text", "").strip()
            if not text:
                continue
            if current is None:
                current = {"text": text, "x0": word.get("x0", 0), "x1": word.get("x1", 0)}
                continue
            if word.get("x0", 0) - current["x1"] <= self.HEADER_WORD_GAP:
                current["text"] = f"{current['text']} {text}".strip()
                current["x1"] = max(current["x1"], word.get("x1", 0))
            else:
                grouped.append(current)
                current = {"text": text, "x0": word.get("x0", 0), "x1": word.get("x1", 0)}

        if current:
            grouped.append(current)

        columns = [g["text"] for g in grouped if g["text"]]
        bounds = []
        for i, col in enumerate(grouped):
            if i == 0:
                start = 0
            else:
                start = (grouped[i - 1]["x1"] + col["x0"]) / 2
            if i == len(grouped) - 1:
                end = page_width
            else:
                end = (col["x1"] + grouped[i + 1]["x0"]) / 2
            bounds.append((start, end))

        return columns, bounds

    def _build_rows_from_lines(
        self,
        lines: List[LineInfo],
        column_bounds: List[Tuple[float, float]]
    ) -> Tuple[List[List[str]], List[str]]:
        rows: List[List[str]] = []
        supplementary: List[str] = []

        for line in lines:
            text = line.text.strip()
            if not text:
                continue
            if self.SUPPLEMENTARY_PATTERN.search(text):
                supplementary.append(text)
                continue
            if self.TABLE_HEADER_LINE_PATTERN.search(text):
                continue

            row = [""] * len(column_bounds)
            for word in line.words:
                word_text = word.get("text", "").strip()
                if not word_text:
                    continue
                center = (word.get("x0", 0) + word.get("x1", 0)) / 2
                for idx, (start, end) in enumerate(column_bounds):
                    if start <= center < end:
                        row[idx] = f"{row[idx]} {word_text}".strip()
                        break

            if any(cell for cell in row):
                rows.append(row)

        return rows, supplementary

    def _normalize_verdict(self, value: str) -> str:
        if not value:
            return ""
        value = value.strip().upper()
        if value in {"P"}:
            return "P"
        if value in {"N/A", "NA", "N.A."}:
            return "N/A"
        if value in {"F", "FAIL"}:
            return "Fail"
        return value

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
