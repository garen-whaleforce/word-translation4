#!/usr/bin/env python3
"""
能量源圖提取器 (Energy Source Diagram Extractor)

從 CB PDF 提取能量源圖:
- 圖像區塊提取
- ES/PS/MS/TS/RS 勾選狀態解析
"""
import re
import io
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field, asdict

import pdfplumber
from PIL import Image

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class EnergySourceCheckbox:
    """能量源勾選項"""
    category: str  # ES, PS, MS, TS, RS
    label: str
    checked: bool = False
    confidence: float = 0.0


@dataclass
class EnergyDiagramResult:
    """能量源圖提取結果"""
    found: bool = False
    page_number: int = 0

    # 圖像
    diagram_image_path: Optional[str] = None
    diagram_width: int = 0
    diagram_height: int = 0

    # 勾選狀態
    checkboxes: List[EnergySourceCheckbox] = field(default_factory=list)

    # 原始文字
    raw_text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "found": self.found,
            "page_number": self.page_number,
            "diagram_image_path": self.diagram_image_path,
            "diagram_width": self.diagram_width,
            "diagram_height": self.diagram_height,
            "checkboxes": [asdict(cb) for cb in self.checkboxes],
            "checkbox_summary": self._get_checkbox_summary(),
        }

    def _get_checkbox_summary(self) -> Dict[str, List[str]]:
        """取得勾選摘要"""
        summary = {"ES": [], "PS": [], "MS": [], "TS": [], "RS": []}
        for cb in self.checkboxes:
            if cb.checked and cb.category in summary:
                summary[cb.category].append(cb.label)
        return summary


class EnergyDiagramExtractor:
    """能量源圖提取器"""

    # 能量源類別
    ENERGY_SOURCE_CATEGORIES = {
        "ES": "Electrical energy source",
        "PS": "Potential ignition source",
        "MS": "Mechanical energy source",
        "TS": "Thermal energy source",
        "RS": "Radiation energy source",
    }

    # 常見的能量源項目
    ES_ITEMS = [
        "mains", "external", "internal", "battery", "capacitor",
        "ES1", "ES2", "ES3"
    ]

    PS_ITEMS = [
        "PS1", "PS2", "PS3", "arcing", "resistive", "lamp"
    ]

    MS_ITEMS = [
        "MS1", "MS2", "MS3", "moving parts", "high pressure",
        "implosion", "explosion"
    ]

    TS_ITEMS = [
        "TS1", "TS2", "TS3", "surface", "internal"
    ]

    RS_ITEMS = [
        "RS1", "RS2", "RS3", "laser", "LED", "lamp", "X-ray",
        "ionizing", "acoustic"
    ]

    # 勾選符號模式
    CHECKBOX_CHECKED_PATTERNS = [
        r'☑', r'☒', r'\[x\]', r'\[X\]', r'✓', r'✔', r'√',
        r'\byes\b', r'\bYES\b', r'\bY\b'
    ]

    CHECKBOX_UNCHECKED_PATTERNS = [
        r'☐', r'□', r'\[\s*\]', r'\bno\b', r'\bNO\b', r'\bN\b'
    ]

    def __init__(self, pdf_path: Path):
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF not found: {pdf_path}")

        self.result = EnergyDiagramResult()

    def extract(self, output_dir: Optional[Path] = None) -> EnergyDiagramResult:
        """
        提取能量源圖

        Args:
            output_dir: 圖像輸出目錄

        Returns:
            EnergyDiagramResult
        """
        logger.info(f"提取能量源圖: {self.pdf_path.name}")

        with pdfplumber.open(self.pdf_path) as pdf:
            # 找到能量源圖頁面
            diagram_page = self._find_diagram_page(pdf)

            if diagram_page:
                page_num, page = diagram_page
                self.result.found = True
                self.result.page_number = page_num

                # 提取文字
                self.result.raw_text = page.extract_text() or ""

                # 解析勾選狀態
                self._parse_checkboxes(page)

                # 提取圖像
                if output_dir:
                    self._extract_diagram_image(page, page_num, output_dir)

        if self.result.found:
            logger.info(f"  找到能量源圖: 第 {self.result.page_number} 頁")
            logger.info(f"  勾選項目: {len([cb for cb in self.result.checkboxes if cb.checked])} 個")
        else:
            logger.warning("  未找到能量源圖")

        return self.result

    def _find_diagram_page(self, pdf: pdfplumber.PDF) -> Optional[Tuple[int, Any]]:
        """找到能量源圖頁面"""
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""

            # 檢查是否包含能量源圖標題
            if re.search(r"ENERGY\s+SOURCE\s+DIAGRAM", text, re.IGNORECASE):
                return (page_num + 1, page)

            # 備選: 檢查是否有 ES/PS/MS/TS/RS 組合
            if all(cat in text for cat in ["ES", "PS", "MS"]):
                if re.search(r"diagram|Diagram|DIAGRAM", text):
                    return (page_num + 1, page)

        return None

    def _parse_checkboxes(self, page):
        """解析勾選狀態"""
        text = page.extract_text() or ""
        lines = text.split('\n')

        current_category = None

        for line in lines:
            line_upper = line.upper().strip()

            # 偵測類別
            for cat, desc in self.ENERGY_SOURCE_CATEGORIES.items():
                if cat in line_upper or desc.upper() in line_upper:
                    current_category = cat
                    break

            # 偵測勾選狀態
            if current_category:
                # 檢查是否有勾選符號
                is_checked = any(
                    re.search(pattern, line)
                    for pattern in self.CHECKBOX_CHECKED_PATTERNS
                )

                # 提取項目名稱
                items = self._get_items_for_category(current_category)
                for item in items:
                    if item.lower() in line.lower():
                        self.result.checkboxes.append(EnergySourceCheckbox(
                            category=current_category,
                            label=item,
                            checked=is_checked,
                            confidence=0.7 if is_checked else 0.5
                        ))

        # 如果沒有偵測到具體勾選，從表格嘗試
        if not self.result.checkboxes:
            self._parse_from_tables(page)

    def _get_items_for_category(self, category: str) -> List[str]:
        """取得類別的項目列表"""
        items_map = {
            "ES": self.ES_ITEMS,
            "PS": self.PS_ITEMS,
            "MS": self.MS_ITEMS,
            "TS": self.TS_ITEMS,
            "RS": self.RS_ITEMS,
        }
        return items_map.get(category, [])

    def _parse_from_tables(self, page):
        """從表格解析勾選狀態"""
        tables = page.extract_tables()

        for table in tables:
            if not table:
                continue

            for row in table:
                if not row or len(row) < 2:
                    continue

                # 檢查是否有能量源類別
                first_cell = str(row[0]).strip().upper() if row[0] else ""

                for cat in self.ENERGY_SOURCE_CATEGORIES.keys():
                    if cat in first_cell:
                        # 檢查其他欄位是否有勾選
                        for i, cell in enumerate(row[1:], 1):
                            cell_text = str(cell).strip() if cell else ""

                            is_checked = any(
                                re.search(pattern, cell_text)
                                for pattern in self.CHECKBOX_CHECKED_PATTERNS
                            )

                            if is_checked or cell_text:
                                self.result.checkboxes.append(EnergySourceCheckbox(
                                    category=cat,
                                    label=f"Item {i}",
                                    checked=is_checked,
                                    confidence=0.6
                                ))

    def _extract_diagram_image(self, page, page_num: int, output_dir: Path):
        """提取圖像"""
        try:
            output_dir.mkdir(parents=True, exist_ok=True)

            # 方法1: 提取頁面內嵌圖像
            images = page.images
            if images:
                for i, img in enumerate(images):
                    # 取得圖像區域
                    x0, y0, x1, y1 = img['x0'], img['top'], img['x1'], img['bottom']

                    # 裁剪頁面
                    cropped = page.within_bbox((x0, y0, x1, y1))
                    if cropped:
                        # 轉換為 PIL Image
                        pil_img = cropped.to_image(resolution=150).original

                        # 儲存
                        img_path = output_dir / f"energy_diagram_p{page_num}_{i}.png"
                        pil_img.save(img_path)

                        self.result.diagram_image_path = str(img_path)
                        self.result.diagram_width = pil_img.width
                        self.result.diagram_height = pil_img.height

                        logger.info(f"  已提取圖像: {img_path}")
                        return

            # 方法2: 將整頁轉為圖像
            pil_img = page.to_image(resolution=150).original

            img_path = output_dir / f"energy_diagram_page_{page_num}.png"
            pil_img.save(img_path)

            self.result.diagram_image_path = str(img_path)
            self.result.diagram_width = pil_img.width
            self.result.diagram_height = pil_img.height

            logger.info(f"  已提取整頁圖像: {img_path}")

        except Exception as e:
            logger.warning(f"  提取圖像失敗: {e}")


def extract_energy_diagram(
    pdf_path: Path,
    output_dir: Optional[Path] = None
) -> EnergyDiagramResult:
    """
    便捷函數：提取能量源圖

    Args:
        pdf_path: PDF 路徑
        output_dir: 圖像輸出目錄

    Returns:
        EnergyDiagramResult
    """
    extractor = EnergyDiagramExtractor(pdf_path)
    return extractor.extract(output_dir)


def main():
    import argparse
    import json

    parser = argparse.ArgumentParser(description='提取能量源圖')
    parser.add_argument('pdf', help='PDF 檔案路徑')
    parser.add_argument('--output', '-o', help='輸出目錄')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    output_dir = Path(args.output) if args.output else Path(f"output/{pdf_path.stem}")

    result = extract_energy_diagram(pdf_path, output_dir)

    print(f"\n{'='*60}")
    print("能量源圖提取結果")
    print(f"{'='*60}")
    print(f"找到: {result.found}")
    print(f"頁碼: {result.page_number}")
    print(f"圖像: {result.diagram_image_path}")

    if result.checkboxes:
        print(f"\n勾選項目:")
        for cb in result.checkboxes:
            status = "✓" if cb.checked else "☐"
            print(f"  {status} [{cb.category}] {cb.label}")

    # 儲存 JSON
    json_path = output_dir / "energy_diagram.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result.to_dict(), f, ensure_ascii=False, indent=2)
    print(f"\n已儲存: {json_path}")


if __name__ == '__main__':
    main()
