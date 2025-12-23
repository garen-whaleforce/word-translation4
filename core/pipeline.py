# core/pipeline.py
"""
CB PDF → CNS DOCX 轉換 Pipeline (v2)

使用新的 PDF 範圍翻譯方法：
- 保留 page 1-4 封面資訊抓取
- page 5 開始使用 PDF 範圍直接翻譯
"""
import os
import sys
import json
import tempfile
import shutil
from pathlib import Path
from typing import Tuple, Optional
from datetime import datetime

# 添加專案根目錄到 path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.models import Job, JobStatus
from core.storage import get_storage, StorageClient


def process_job(job: Job, storage: Optional[StorageClient] = None, redis_client=None) -> Job:
    """
    處理一個 CB PDF 轉換任務 (v2 - 使用新的 PDF 範圍翻譯方法)

    流程:
    1. 下載 PDF 到暫存目錄
    2. 抽取封面資訊 (meta data for page 1-4)
    3. PDF 範圍翻譯 (page 5 開始，直接翻譯並插入模板)
    4. 上傳結果到 Storage
    5. 更新 Job 狀態

    Args:
        job: Job 物件
        storage: StorageClient (預設使用全局實例)
        redis_client: Redis client (用於即時更新進度)

    Returns:
        更新後的 Job 物件
    """
    if storage is None:
        storage = get_storage()

    def update_progress(gate_name: str, status: str, message: str):
        """更新進度到 Redis"""
        job.add_qa_result(gate_name, status, message)
        if redis_client:
            try:
                redis_client.set(f"job:{job.job_id}", job.to_json())
            except:
                pass

    # 建立暫存目錄
    work_dir = Path(tempfile.mkdtemp(prefix=f"cns_{job.job_id}_"))
    out_dir = work_dir / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        job.update_status(JobStatus.RUNNING)
        update_progress("start", "PASS", "開始處理任務")

        # ===== Step 1: 下載 PDF =====
        update_progress("download_pdf", "PASS", "正在下載 PDF...")
        pdf_path = work_dir / job.pdf_filename
        storage.download_file(job.original_pdf_key, str(pdf_path))
        update_progress("download_pdf_done", "PASS", "PDF 下載完成")

        # ===== Step 2: 抽取封面資訊 (page 1-4) =====
        update_progress("extract_meta", "PASS", "正在抽取封面資訊...")
        meta = _extract_cover_meta(str(pdf_path), job.pdf_filename)
        update_progress("extract_meta_done", "PASS", "封面資訊抽取完成")

        # ===== Step 3: PDF 範圍翻譯 (page 5 開始) =====
        update_progress("translate_pdf", "PASS", "正在翻譯 PDF 內容...")

        from tools.translate_pdf_range import (
            find_translation_range,
            extract_tables_from_range,
            translate_tables
        )

        # 3a. 找出翻譯範圍
        start_page, end_page = find_translation_range(str(pdf_path))
        update_progress("find_range", "PASS", f"翻譯範圍: Page {start_page + 1} ~ {end_page}")

        # 3b. 抽取表格
        tables = extract_tables_from_range(str(pdf_path), start_page, end_page)
        update_progress("extract_tables", "PASS", f"抽取 {len(tables)} 個表格")

        # 3c. 翻譯表格
        update_progress("llm_translate", "PASS", "正在進行 LLM 翻譯...")
        translated_tables = translate_tables(tables)
        update_progress("llm_translate_done", "PASS", "LLM 翻譯完成")

        # ===== Step 4: 渲染 Word =====
        update_progress("render_word", "PASS", "正在渲染 Word 報告...")
        template_path = PROJECT_ROOT / "templates" / "CNS_15598_1_109_template_clean.docx"
        docx_path = out_dir / "cns_report.docx"

        # 準備封面欄位
        cover_fields = {
            'report_no': job.cover_report_no,
            'applicant_name': job.cover_applicant_name,
            'applicant_address': job.cover_applicant_address,
        }

        _render_word_v2(
            template_path=str(template_path),
            translated_tables=translated_tables,
            meta=meta,
            cover_fields=cover_fields,
            output_path=str(docx_path)
        )
        update_progress("render_word_done", "PASS", "Word 報告渲染完成")

        # ===== Step 5: 設定狀態並上傳 =====
        update_progress("upload_results", "PASS", "正在上傳結果...")
        job.update_status(JobStatus.PASS)
        job.docx_type = "FINAL"

        # 上傳結果
        job_prefix = f"jobs/{job.job_id}"

        # 上傳 DOCX
        docx_key = f"{job_prefix}/cns_report.docx"
        storage.upload_file(str(docx_path), docx_key)
        job.docx_key = docx_key

        # 儲存翻譯後的表格資料 (for debugging)
        tables_json_path = out_dir / "translated_tables.json"
        with open(tables_json_path, 'w', encoding='utf-8') as f:
            json.dump(translated_tables, f, ensure_ascii=False, indent=2)

        json_key = f"{job_prefix}/translated_tables.json"
        storage.upload_file(str(tables_json_path), json_key)
        job.json_data_key = json_key

        # 讀取 LLM 統計
        from core.llm_translator import get_translator
        translator = get_translator()
        job.llm_stats = translator.get_cost_estimate()

        update_progress("final_qa", "PASS", "任務完成")

    except Exception as e:
        job.update_status(JobStatus.ERROR)
        job.error_message = str(e)
        import traceback
        print(f"Pipeline error: {traceback.format_exc()}")

    finally:
        # 清理暫存目錄
        try:
            shutil.rmtree(work_dir)
        except:
            pass

    return job


def _extract_cover_meta(pdf_path: str, pdf_filename: str) -> dict:
    """
    抽取封面資訊 (page 1-4)

    Returns:
        dict: {
            'model_type_references': [...],
            'model_type_references_str': '...',
            'report_reference': '...',
            'manufacturer_name': '...',
            ...
        }
    """
    import pdfplumber
    import re

    meta = {
        'model_type_references': [],
        'model_type_references_str': '',
        'report_reference': '',
        'manufacturer_name': '',
        'manufacturer_address': '',
        'test_item': '',
        'ratings': '',
    }

    with pdfplumber.open(pdf_path) as pdf:
        # 從前幾頁抽取文字
        text = ""
        for i in range(min(10, len(pdf.pages))):
            text += (pdf.pages[i].extract_text() or "") + "\n"

        # 抽取型號
        model_patterns = [
            r'Model[:\s]+([A-Z0-9\-]+)',
            r'Type[:\s]+([A-Z0-9\-]+)',
            r'Model/Type[:\s]+([A-Z0-9\-]+)',
        ]
        for pattern in model_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                meta['model_type_references'] = list(set(matches))
                break

        meta['model_type_references_str'] = ', '.join(meta['model_type_references'])

        # 抽取報告編號
        report_patterns = [
            r'Report\s+(?:Reference|No\.?)[:\s]+([A-Z0-9\-/]+)',
            r'Certificate\s+No\.?[:\s]+([A-Z0-9\-/]+)',
        ]
        for pattern in report_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                meta['report_reference'] = match.group(1)
                break

        # 抽取製造商名稱
        mfr_patterns = [
            r'Manufacturer[:\s]+(.+?)(?:\n|Address)',
            r'Applicant[:\s]+(.+?)(?:\n|Address)',
        ]
        for pattern in mfr_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                meta['manufacturer_name'] = match.group(1).strip()
                break

    return meta


def _render_word_v2(
    template_path: str,
    translated_tables: list,
    meta: dict,
    cover_fields: dict,
    output_path: str
):
    """
    渲染 Word 文件 (v2 - 使用翻譯後的表格)

    1. 載入模板
    2. 填充封面資訊 (page 1-4)
    3. 插入翻譯後的表格 (page 5 開始)
    4. 儲存

    表格格式：
    - 保留 PDF 原有的欄位結構（不強制 4 欄）
    - 判定欄: P→符合, N/A→不適用, F→不符合
    - 章節標題行（第一欄為數字如 4, 5, 6）加灰色背景 #D9D9D9
    """
    from docx import Document
    from docx.shared import Pt, Twips, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import re

    doc = Document(template_path)

    # ===== 填充封面資訊 =====
    _fill_cover_fields(doc, meta, cover_fields)

    # ===== 插入翻譯後的表格 =====
    # 找到插入位置（表格 3 之後）
    insert_after_table_idx = 3

    if insert_after_table_idx < len(doc.tables):
        last_table = doc.tables[insert_after_table_idx]
        insert_element = last_table._tbl
    else:
        insert_element = doc.element.body[-1]

    # 總寬度 (twips) - A4 紙張內容區域約 9589 twips
    TOTAL_WIDTH = 9589

    # 判定值映射
    VERDICT_MAP = {
        'P': '符合',
        'PASS': '符合',
        'N/A': '不適用',
        'NA': '不適用',
        'F': '不符合',
        'FAIL': '不符合',
    }

    # 需要灰色背景的行類型判斷
    def should_have_gray_background(row: list, row_idx: int, all_rows: list) -> bool:
        """
        判斷該行是否需要灰色背景

        PDF 原始格式中，以下行需要反灰：
        1. 表格標題行（如 OVERVIEW OF ENERGY SOURCES...）
        2. 章節標題行（第一欄為純數字 4, 5, 6 或字母 B, G, M）
        3. 子標題行（如 Class and Energy Source, Safeguards 等）
        4. 欄位標題行（如 B, S, R 或 1st S, 2nd S）
        """
        if not row:
            return False

        first_col = (row[0] or "").strip()
        row_text = ' '.join([str(c or '') for c in row]).strip()
        row_text_upper = row_text.upper()

        # 1. 表格標題行（通常是第一行，且內容較長或包含特定關鍵字）
        if row_idx == 0:
            title_keywords = [
                'OVERVIEW', 'ENERGY SOURCES', 'SAFEGUARDS', '能源', '防護', '總覽',
                '安全防護', 'ENERGY SOURCE'
            ]
            for kw in title_keywords:
                if kw.upper() in row_text_upper or kw in row_text:
                    return True

        # 2. 章節標題行（第一欄為純數字或單一大寫字母）
        # 主章節：純數字（4, 5, 6...10）
        if re.match(r'^[4-9]$|^10$|^[1-9][0-9]$', first_col):
            return True
        # 附錄章節：單一大寫字母（B, G, M, etc.）
        if re.match(r'^[A-Z]$', first_col):
            return True

        # 3. 子標題行（包含特定標題關鍵字）
        subtitle_keywords = [
            # 英文
            'Class and Energy Source', 'Body Part', 'Safeguards',
            'Material part', 'Possible Hazard', 'Clause',
            # 中文
            '等級與能源來源', '類別和能量來源', '身體部位', '防護措施',
            '材料部位', '可能的危險', '條款', '安全防護'
        ]
        for kw in subtitle_keywords:
            if kw.lower() in row_text.lower() or kw in row_text:
                # 但排除純數據行（如 ES3: xxx）
                if not first_col.startswith('ES') and not first_col.startswith('PS'):
                    return True

        # 4. 欄位標題行（短標題如 B, S, R）
        # 檢查是否所有非空欄位都是短標題（1-3 個字元）
        non_empty_cells = [c for c in row if c and str(c).strip()]
        if len(non_empty_cells) >= 2:
            short_headers = ['B', 'S', 'R', '1st S', '2nd S', '1st', '2nd', '基本', '補充', '強化']
            if any(str(c).strip() in short_headers for c in non_empty_cells):
                # 如果大部分欄位都是短標題，則為標題行
                short_count = sum(1 for c in non_empty_cells if str(c).strip() in short_headers or len(str(c).strip()) <= 3)
                if short_count >= len(non_empty_cells) // 2:
                    return True

        return False

    # 判定欄判斷：檢查內容是否為判定值
    def is_verdict_cell(cell_text: str) -> bool:
        if not cell_text:
            return False
        text = cell_text.strip().upper()
        return text in VERDICT_MAP

    # 逐個插入表格
    for t_idx, table_data in enumerate(translated_tables):
        rows = table_data['rows']
        col_count = table_data['col_count']
        merge_info = table_data.get('merge_info', [])

        if not rows:
            continue

        # 保留原始欄位數
        actual_cols = col_count

        # 建立新表格（使用 PDF 原有的欄位數）
        new_table = doc.add_table(rows=len(rows), cols=actual_cols)

        # 計算平均欄寬
        col_widths = [TOTAL_WIDTH // actual_cols] * actual_cols

        # 設定表格寬度和欄寬
        _set_table_width(new_table, TOTAL_WIDTH)
        _set_column_widths(new_table, col_widths)

        # 設定表格框線
        _set_table_borders(new_table)

        # 建立合併資訊的快速查詢表
        merge_lookup = {}
        for m in merge_info:
            key = (m['row'], m['col'])
            merge_lookup[key] = m['colspan']

        # 先處理合併儲存格
        for m in merge_info:
            r_idx = m['row']
            c_idx = m['col']
            colspan = m['colspan']

            if r_idx < len(new_table.rows) and c_idx < len(new_table.rows[r_idx].cells):
                try:
                    # 合併儲存格：從 c_idx 到 c_idx + colspan - 1
                    start_cell = new_table.rows[r_idx].cells[c_idx]
                    end_col = min(c_idx + colspan - 1, len(new_table.rows[r_idx].cells) - 1)
                    if end_col > c_idx:
                        end_cell = new_table.rows[r_idx].cells[end_col]
                        start_cell.merge(end_cell)
                except Exception as e:
                    # 合併失敗時繼續（可能是範圍超出）
                    pass

        # 填入資料
        for r_idx, row in enumerate(rows):
            # 判斷是否需要灰色背景
            needs_gray_bg = should_have_gray_background(row, r_idx, rows)

            for c_idx, cell_text in enumerate(row):
                if c_idx < len(new_table.rows[r_idx].cells):
                    cell = new_table.rows[r_idx].cells[c_idx]

                    # 如果是被合併的儲存格（不是起始儲存格），跳過
                    # 檢查此 cell 是否為某個合併範圍的起始點
                    is_merge_start = (r_idx, c_idx) in merge_lookup

                    # 判定欄（最後一欄或內容為判定值）轉換
                    if cell_text and is_verdict_cell(cell_text):
                        cell_text = VERDICT_MAP.get(cell_text.strip().upper(), cell_text)

                    # 只在起始儲存格或非合併儲存格填入內容
                    if cell_text:
                        cell.text = cell_text

                    # 設定字型
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(11)
                            run.font.name = '標楷體'
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

                    # 需要灰色背景的行：除了最後一欄（判定欄）外都加灰色背景
                    if needs_gray_bg and c_idx < actual_cols - 1:
                        _set_cell_shading(cell, "D9D9D9")

        # 移動表格到正確位置
        insert_element.addnext(new_table._tbl)
        insert_element = new_table._tbl

        if (t_idx + 1) % 20 == 0:
            print(f"  已插入 {t_idx + 1}/{len(translated_tables)} 個表格...")

    # 儲存
    doc.save(output_path)
    print(f"[完成] 輸出: {output_path}")


def _normalize_table_rows(rows: list, original_cols: int, target_cols: int) -> list:
    """
    標準化表格行到目標欄數

    策略：
    - 如果原本 < 4 欄：合併到前幾欄，最後一欄留給判定
    - 如果原本 > 4 欄：合併中間欄位
    - 如果原本 = 4 欄：直接使用
    """
    if original_cols == target_cols:
        return rows

    normalized = []
    for row in rows:
        if len(row) == target_cols:
            normalized.append(row)
        elif len(row) < target_cols:
            # 補空欄
            new_row = row + [''] * (target_cols - len(row))
            normalized.append(new_row)
        else:
            # 合併多餘欄位（保留第一欄和最後一欄，中間合併）
            first_col = row[0]
            last_col = row[-1]
            middle_cols = row[1:-1]

            # 根據內容分配到 3 個中間位置
            if len(middle_cols) >= 2:
                col1 = middle_cols[0]
                col2 = ' '.join(middle_cols[1:])
            else:
                col1 = ' '.join(middle_cols)
                col2 = ''

            normalized.append([first_col, col1, col2, last_col])

    return normalized


def _set_table_width(table, width_twips: int):
    """設定表格總寬度"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')

    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(width_twips))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def _set_column_widths(table, widths: list):
    """設定各欄寬度"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    tbl = table._tbl

    # 建立或取得 tblGrid
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is None:
        tblGrid = OxmlElement('w:tblGrid')
        tbl.insert(0, tblGrid)
    else:
        # 清除現有的 gridCol
        for child in list(tblGrid):
            tblGrid.remove(child)

    # 加入欄寬定義
    for width in widths:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(width))
        tblGrid.append(gridCol)

    # 設定每個 cell 的寬度
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths):
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = OxmlElement('w:tcW')
                tcW.set(qn('w:w'), str(widths[idx]))
                tcW.set(qn('w:type'), 'dxa')
                tcPr.append(tcW)


def _fill_cover_fields(doc, meta: dict, cover_fields: dict):
    """填充封面欄位"""
    # 遍歷所有表格找封面表格
    for table in doc.tables[:4]:  # 只看前 4 個表格
        for row in table.rows:
            for cell in row.cells:
                text = cell.text

                # 替換封面欄位佔位符
                if '{{ meta.model_type_references_str }}' in text:
                    cell.text = text.replace('{{ meta.model_type_references_str }}',
                                            meta.get('model_type_references_str', ''))

                if '{{ cover_report_no }}' in text or '報告編號' in text:
                    if cover_fields.get('report_no'):
                        # 找到報告編號欄位並填入
                        pass

                if '{{ cover_applicant_name }}' in text:
                    cell.text = text.replace('{{ cover_applicant_name }}',
                                            cover_fields.get('applicant_name', ''))

                if '{{ cover_applicant_address }}' in text:
                    cell.text = text.replace('{{ cover_applicant_address }}',
                                            cover_fields.get('applicant_address', ''))


def _set_table_borders(table):
    """設定表格框線"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')

    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)

    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def _set_cell_shading(cell, color: str):
    """
    設定儲存格背景色

    Args:
        cell: Word 儲存格物件
        color: 16 進位顏色碼 (如 "D9D9D9")
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # 移除現有的 shading
    existing_shd = tcPr.find(qn('w:shd'))
    if existing_shd is not None:
        tcPr.remove(existing_shd)

    # 建立新的 shading
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


# ============================================================
# 舊版 pipeline 函數 (保留供向後兼容)
# ============================================================

def process_job_legacy(job: Job, storage: Optional[StorageClient] = None, redis_client=None) -> Job:
    """
    舊版 pipeline (保留供向後兼容)
    """
    # 舊版邏輯...
    pass
