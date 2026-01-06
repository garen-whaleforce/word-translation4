#!/usr/bin/env python3
"""
從 PDF 生成翻譯後的 Word 檔案

流程:
1. 使用 parse_cb_pdf_v2 解析 PDF
2. 使用 translate_and_compare 中的翻譯器翻譯
3. 填充 Word 模板輸出
"""
import sys
import json
import logging
from pathlib import Path
from typing import List, Dict

from docx import Document
from docx.shared import Pt

# 添加專案路徑
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.parse_cb_pdf_v2 import CBParserV2, ParseResultV2, ClauseItem
from scripts.translate_and_compare import PDFTranslator

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


def generate_word_from_pdf(
    pdf_path: Path,
    template_path: Path,
    output_path: Path
) -> Dict:
    """
    從 PDF 生成翻譯後的 Word 檔案
    """
    result = {
        'pdf': str(pdf_path),
        'template': str(template_path),
        'output': str(output_path),
        'clauses_parsed': 0,
        'clauses_translated': 0,
        'clauses_filled': 0,
    }

    logger.info(f"\n{'='*60}")
    logger.info("PDF → Word 翻譯輸出")
    logger.info(f"{'='*60}")
    logger.info(f"PDF: {pdf_path.name}")
    logger.info(f"模板: {template_path.name}")

    # 1. 解析 PDF
    logger.info(f"\n[步驟 1] 解析 PDF...")
    parser = CBParserV2(pdf_path)
    parse_result = parser.parse()
    result['clauses_parsed'] = len(parse_result.clauses)
    logger.info(f"  ✓ 擷取 {result['clauses_parsed']} 個條款")

    # 2. 翻譯
    logger.info(f"\n[步驟 2] 翻譯條款...")
    translator = PDFTranslator()

    # 收集需要翻譯的文字
    texts_to_translate = []
    for clause in parse_result.clauses:
        combined = f"{clause.requirement} {clause.result_remark}".strip()
        texts_to_translate.append(combined)

    # 批量翻譯
    translated_texts = translator.translate_batch(texts_to_translate)
    result['clauses_translated'] = sum(1 for t in translated_texts if t.strip())
    logger.info(f"  ✓ 翻譯完成 ({result['clauses_translated']} 個非空)")

    # 3. 填充 Word 模板
    logger.info(f"\n[步驟 3] 填充 Word 模板...")
    doc = Document(template_path)

    # 找到 Clause 表格
    clause_table = None
    for table in doc.tables:
        if table.rows:
            header = ' '.join(c.text for c in table.rows[0].cells).upper()
            if 'CLAUSE' in header and 'VERDICT' in header:
                clause_table = table
                break

    if not clause_table:
        logger.error("  ✗ 找不到 Clause 表格")
        result['error'] = "找不到 Clause 表格"
        return result

    # 清空現有數據行 (保留表頭)
    while len(clause_table.rows) > 1:
        clause_table._tbl.remove(clause_table.rows[-1]._tr)

    # 填充數據
    for i, clause in enumerate(parse_result.clauses):
        translated = translated_texts[i] if i < len(translated_texts) else ""

        row = clause_table.add_row()
        cells = row.cells

        # 填充欄位
        cells[0].text = clause.clause_id
        cells[1].text = translated or clause.requirement  # 翻譯後的 requirement
        cells[2].text = clause.result_remark or ""
        cells[3].text = clause.verdict or ""

        result['clauses_filled'] += 1

    logger.info(f"  ✓ 填充 {result['clauses_filled']} 列")

    # 4. 保存
    logger.info(f"\n[步驟 4] 保存輸出...")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    logger.info(f"  ✓ 輸出: {output_path}")

    return result


def main():
    import argparse

    parser = argparse.ArgumentParser(description='從 PDF 生成翻譯後的 Word')
    parser.add_argument('--pdf', required=True, help='PDF 檔案路徑')
    parser.add_argument('--template', default='templates/Generic-CB.docx', help='Word 模板')
    parser.add_argument('--output', '-o', help='輸出路徑 (預設: output/{pdf_stem}_translated.docx)')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    template_path = Path(args.template)

    if not pdf_path.exists():
        print(f"錯誤: PDF 不存在: {pdf_path}")
        sys.exit(1)
    if not template_path.exists():
        print(f"錯誤: 模板不存在: {template_path}")
        sys.exit(1)

    output_path = Path(args.output) if args.output else Path(f"output/{pdf_path.stem}_translated.docx")

    result = generate_word_from_pdf(pdf_path, template_path, output_path)

    print(f"\n{'='*60}")
    print("完成!")
    print(f"{'='*60}")
    print(f"條款解析: {result['clauses_parsed']}")
    print(f"條款翻譯: {result['clauses_translated']}")
    print(f"條款填充: {result['clauses_filled']}")
    print(f"輸出檔案: {result['output']}")


if __name__ == '__main__':
    main()
