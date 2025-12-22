# core/pipeline.py
"""
CB PDF → CNS DOCX 轉換 Pipeline
封裝所有 tools/ 下的處理邏輯為單一 process_job() 函數
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


def process_job(job: Job, storage: Optional[StorageClient] = None) -> Job:
    """
    處理一個 CB PDF 轉換任務

    流程:
    1. 下載 PDF 到暫存目錄
    2. 抽取 PDF 資料 (extract_cb_pdf, extract_pdf_clause_rows, extract_special_tables)
    3. 生成 CNS JSON (generate_cns_json)
    4. 渲染 Word 文件 (render_word)
    5. 上傳結果到 Storage
    6. 更新 Job 狀態

    Args:
        job: Job 物件
        storage: StorageClient (預設使用全局實例)

    Returns:
        更新後的 Job 物件
    """
    if storage is None:
        storage = get_storage()

    # 建立暫存目錄
    work_dir = Path(tempfile.mkdtemp(prefix=f"cns_{job.job_id}_"))
    out_dir = work_dir / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        job.update_status(JobStatus.RUNNING)

        # ===== Step 1: 下載 PDF =====
        pdf_path = work_dir / job.pdf_filename
        storage.download_file(job.original_pdf_key, str(pdf_path))

        # ===== Step 2: 抽取 PDF 資料 =====
        from tools.extract_cb_pdf import main as extract_cb_pdf_main, extract_annex_model_rows
        from tools.extract_pdf_clause_rows import extract_clause_rows, find_clause_start_page
        from tools.extract_special_tables import (
            extract_overview_energy_sources,
            extract_table_5522,
            extract_table_b25,
            extract_table_52
        )
        import pdfplumber

        # 2a. 基本抽取
        _run_extract_cb_pdf(str(pdf_path), str(out_dir))

        # 2b. 條款列抽取
        with pdfplumber.open(str(pdf_path)) as pdf:
            start_idx = find_clause_start_page(pdf)
            clause_rows = extract_clause_rows(pdf, start_idx)

        clause_rows_path = out_dir / "pdf_clause_rows.json"
        with open(clause_rows_path, 'w', encoding='utf-8') as f:
            json.dump(clause_rows, f, ensure_ascii=False, indent=2)

        # 2c. 特殊表格抽取
        with pdfplumber.open(str(pdf_path)) as pdf:
            try:
                overview = extract_overview_energy_sources(pdf)
            except ValueError as e:
                overview = {'error': str(e), 'rows': []}

            try:
                table_5522 = extract_table_5522(pdf)
            except ValueError as e:
                table_5522 = {'error': str(e)}

            try:
                table_b25 = extract_table_b25(pdf)
            except ValueError as e:
                table_b25 = {'error': str(e)}

            try:
                table_52 = extract_table_52(pdf)
            except ValueError as e:
                table_52 = {'error': str(e)}

        special_tables = {
            'overview': overview,
            'table_5522': table_5522,
            'table_b25': table_b25,
            'table_52': table_52
        }
        special_tables_path = out_dir / "cb_special_tables.json"
        with open(special_tables_path, 'w', encoding='utf-8') as f:
            json.dump(special_tables, f, ensure_ascii=False, indent=2, default=list)

        # 2d. 附表 Model 行抽取
        cb_tables_path = out_dir / "cb_tables_text.json"
        if cb_tables_path.exists():
            with open(cb_tables_path, 'r', encoding='utf-8') as f:
                cb_tables = json.load(f)
            annex_model_rows = extract_annex_model_rows(cb_tables)
        else:
            annex_model_rows = []

        annex_model_rows_path = out_dir / "cb_annex_model_rows.json"
        with open(annex_model_rows_path, 'w', encoding='utf-8') as f:
            json.dump(annex_model_rows, f, ensure_ascii=False, indent=2)

        # ===== Step 3: 生成 CNS JSON =====
        from tools.generate_cns_json import (
            load_json,
            extract_meta_from_chunks,
            convert_overview_to_cns,
            dedupe_clauses
        )

        chunks = load_json(out_dir / "cb_text_chunks.json")
        overview_raw = load_json(out_dir / "cb_overview_raw.json")
        clauses_raw = load_json(out_dir / "cb_clauses_raw.json")

        meta = extract_meta_from_chunks(chunks, job.pdf_filename)
        overview_cns = convert_overview_to_cns(overview_raw)
        clauses = dedupe_clauses(clauses_raw)

        # overview_cb_p12_rows 保留完整列
        overview_cb_p12_rows = overview.get('rows', []) if 'error' not in overview else []

        cns_data = {
            'meta': meta,
            'overview_energy_sources_and_safeguards': overview_cns,
            'overview_cb_p12_rows': overview_cb_p12_rows,
            'clauses': clauses,
            'attachments_or_annex': []
        }

        cns_json_path = out_dir / "cns_report_data.json"
        with open(cns_json_path, 'w', encoding='utf-8') as f:
            json.dump(cns_data, f, ensure_ascii=False, indent=2)

        # ===== Step 4: 渲染 Word =====
        template_path = PROJECT_ROOT / "templates" / "CNS_15598_1_109_template_clean.docx"
        docx_path = out_dir / "cns_report.docx"

        # 準備封面欄位
        cover_fields = {
            'report_no': job.cover_report_no,
            'applicant_name': job.cover_applicant_name,
            'applicant_address': job.cover_applicant_address,
        }

        _run_render_word(
            json_path=str(cns_json_path),
            template_path=str(template_path),
            pdf_clause_rows_path=str(clause_rows_path),
            special_tables_path=str(special_tables_path),
            output_path=str(docx_path),
            cover_fields=cover_fields,
            annex_model_rows_path=str(annex_model_rows_path)
        )

        # ===== Step 5: 設定狀態並上傳 =====
        job.update_status(JobStatus.PASS)
        job.docx_type = "FINAL"

        # 上傳結果
        job_prefix = f"jobs/{job.job_id}"

        # 上傳 DOCX
        docx_key = f"{job_prefix}/cns_report.docx"
        storage.upload_file(str(docx_path), docx_key)
        job.docx_key = docx_key

        # 上傳 JSON 資料
        json_key = f"{job_prefix}/cns_report_data.json"
        storage.upload_file(str(cns_json_path), json_key)
        job.json_data_key = json_key

        # 上傳 clause rows
        clause_rows_key = f"{job_prefix}/pdf_clause_rows.json"
        storage.upload_file(str(clause_rows_path), clause_rows_key)
        job.pdf_clause_rows_key = clause_rows_key

        # 讀取 LLM 統計並存入 Job
        llm_stats_path = out_dir / "llm_stats.json"
        if llm_stats_path.exists():
            with open(llm_stats_path, 'r', encoding='utf-8') as f:
                job.llm_stats = json.load(f)

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


def _run_extract_cb_pdf(pdf_path: str, out_dir: str):
    """執行 extract_cb_pdf"""
    import pdfplumber
    import re
    from tools.extract_cb_pdf import (
        norm, find_overview_page, extract_overview_table,
        find_clause_pages, extract_clauses_from_pages, extract_table_412
    )

    chunks = []
    tables = []
    overview_data = []
    clauses_data = []

    with pdfplumber.open(pdf_path) as pdf:
        # 抽取所有頁面文字
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            chunks.append({
                "page": i,
                "text": norm(text)
            })

            try:
                tbs = page.extract_tables() or []
            except Exception:
                tbs = []

            for t in tbs:
                rows = []
                for row in t:
                    rows.append([norm(c) if c else "" for c in row])
                tables.append({
                    "page": i,
                    "rows": rows
                })

        # 抽取 Overview 表
        overview_page_idx = find_overview_page(pdf)
        if overview_page_idx >= 0:
            overview_page = pdf.pages[overview_page_idx]
            overview_data = extract_overview_table(overview_page)

        # 抽取 Clause 表
        clause_start_idx = find_clause_pages(pdf)
        if clause_start_idx >= 0:
            clauses_data = extract_clauses_from_pages(pdf, clause_start_idx)

    # 抽取 4.1.2 表格
    table_412_data = extract_table_412(tables)

    # 輸出
    out_path = Path(out_dir)
    with open(out_path / "cb_text_chunks.json", "w", encoding="utf-8") as f:
        json.dump(chunks, f, ensure_ascii=False, indent=2)

    with open(out_path / "cb_tables_text.json", "w", encoding="utf-8") as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)

    with open(out_path / "cb_overview_raw.json", "w", encoding="utf-8") as f:
        json.dump(overview_data, f, ensure_ascii=False, indent=2)

    with open(out_path / "cb_clauses_raw.json", "w", encoding="utf-8") as f:
        json.dump(clauses_data, f, ensure_ascii=False, indent=2)

    with open(out_path / "cb_table_412.json", "w", encoding="utf-8") as f:
        json.dump(table_412_data, f, ensure_ascii=False, indent=2)


def _run_render_word(json_path: str, template_path: str, pdf_clause_rows_path: str,
                     special_tables_path: str, output_path: str,
                     cover_fields: dict = None, annex_model_rows_path: str = None):
    """執行 render_word"""
    import subprocess

    cmd = [
        sys.executable,
        str(PROJECT_ROOT / "tools" / "render_word.py"),
        "--json", json_path,
        "--template", template_path,
        "--pdf_clause_rows", pdf_clause_rows_path,
        "--special_tables", special_tables_path,
        "--out", output_path
    ]

    # 添加附表 Model 行參數
    if annex_model_rows_path:
        cmd.extend(["--annex_model_rows", annex_model_rows_path])

    # 添加封面欄位參數
    if cover_fields:
        if cover_fields.get('report_no'):
            cmd.extend(["--cover_report_no", cover_fields['report_no']])
        if cover_fields.get('applicant_name'):
            cmd.extend(["--cover_applicant_name", cover_fields['applicant_name']])
        if cover_fields.get('applicant_address'):
            cmd.extend(["--cover_applicant_address", cover_fields['applicant_address']])

    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(PROJECT_ROOT))

    if result.returncode != 0:
        raise RuntimeError(f"render_word failed: {result.stderr}")
