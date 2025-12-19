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
    4. QA Gate 1: Overview match
    5. QA Gate 2: Clause table match
    6. 渲染 Word 文件 (render_word)
    7. QA Gate 3: Final QA
    8. 上傳結果到 Storage
    9. 更新 Job 狀態

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
        from tools.extract_cb_pdf import main as extract_cb_pdf_main
        from tools.extract_pdf_clause_rows import extract_clause_rows, find_clause_start_page
        from tools.extract_special_tables import (
            extract_overview_energy_sources,
            extract_table_5522,
            extract_table_b25
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

        job.add_qa_result("extract_clause_rows", "PASS", f"抽取 {len(clause_rows)} 條款列")

        # 2c. 特殊表格抽取
        with pdfplumber.open(str(pdf_path)) as pdf:
            try:
                overview = extract_overview_energy_sources(pdf)
                job.add_qa_result("extract_overview", "PASS", f"Overview {overview['total_rows']} 列")
            except ValueError as e:
                job.add_qa_result("extract_overview", "FAIL", str(e))
                overview = {'error': str(e), 'rows': []}

            try:
                table_5522 = extract_table_5522(pdf)
                job.add_qa_result("extract_5522", "PASS", f"5.5.2.2 verdict={table_5522.get('verdict', 'N/A')}")
            except ValueError as e:
                job.add_qa_result("extract_5522", "WARN", str(e))
                table_5522 = {'error': str(e)}

            try:
                table_b25 = extract_table_b25(pdf)
                job.add_qa_result("extract_b25", "PASS", f"B.2.5 I_rated={table_b25.get('i_rated_values', [])}")
            except ValueError as e:
                job.add_qa_result("extract_b25", "WARN", str(e))
                table_b25 = {'error': str(e)}

        special_tables = {
            'overview': overview,
            'table_5522': table_5522,
            'table_b25': table_b25
        }
        special_tables_path = out_dir / "cb_special_tables.json"
        with open(special_tables_path, 'w', encoding='utf-8') as f:
            json.dump(special_tables, f, ensure_ascii=False, indent=2, default=list)

        # ===== Step 3: 生成 CNS JSON =====
        from tools.generate_cns_json import (
            load_json,
            extract_meta_from_chunks,
            convert_overview_to_cns,
            dedupe_clauses,
            generate_qa
        )

        chunks = load_json(out_dir / "cb_text_chunks.json")
        overview_raw = load_json(out_dir / "cb_overview_raw.json")
        clauses_raw = load_json(out_dir / "cb_clauses_raw.json")

        meta = extract_meta_from_chunks(chunks, job.pdf_filename)
        overview_cns = convert_overview_to_cns(overview_raw)
        clauses = dedupe_clauses(clauses_raw)

        # overview_cb_p12_rows 保留完整 10 列
        overview_cb_p12_rows = overview.get('rows', []) if 'error' not in overview else []

        qa = generate_qa(meta, overview_cns, clauses, overview_raw)

        cns_data = {
            'meta': meta,
            'overview_energy_sources_and_safeguards': overview_cns,
            'overview_cb_p12_rows': overview_cb_p12_rows,
            'clauses': clauses,
            'attachments_or_annex': [],
            'qa': qa
        }

        cns_json_path = out_dir / "cns_report_data.json"
        with open(cns_json_path, 'w', encoding='utf-8') as f:
            json.dump(cns_data, f, ensure_ascii=False, indent=2)

        job.add_qa_result("generate_cns_json", qa['summary']['status'],
                          f"overview={len(overview_cns)}, clauses={len(clauses)}")

        # ===== Step 4: 渲染 Word =====
        template_path = PROJECT_ROOT / "templates" / "CNS_15598_1_109_template_clean.docx"
        docx_path = out_dir / "cns_report.docx"

        _run_render_word(
            json_path=str(cns_json_path),
            template_path=str(template_path),
            pdf_clause_rows_path=str(clause_rows_path),
            special_tables_path=str(special_tables_path),
            output_path=str(docx_path)
        )

        job.add_qa_result("render_word", "PASS", "Word 文件生成完成")

        # ===== Step 5: QA Gates =====
        all_pass = True

        # Gate 1: Overview match
        from tools.qa_overview_match import extract_word_overview_rows, compare_rows
        word_overview_rows = extract_word_overview_rows(docx_path)
        overview_match = compare_rows(overview_cb_p12_rows, word_overview_rows)

        overview_match_path = out_dir / "qa_overview_match.json"
        with open(overview_match_path, 'w', encoding='utf-8') as f:
            json.dump(overview_match, f, ensure_ascii=False, indent=2)

        if overview_match.get('match') and len(overview_match.get('differences', [])) == 0:
            job.add_qa_result("qa_overview_match", "PASS", "Overview 表格一致")
        else:
            job.add_qa_result("qa_overview_match", "WARN",
                              f"差異: {len(overview_match.get('differences', []))}",
                              {"differences": overview_match.get('differences', [])})
            # Overview mismatch 目前設為 WARN，不影響最終結果

        # Gate 2: Clause table match
        from tools.qa_clause_table_match import extract_word_clause_rows, compare_clause_tables
        word_clause_rows = extract_word_clause_rows(docx_path)
        clause_match = compare_clause_tables(clause_rows, word_clause_rows)

        clause_match_path = out_dir / "qa_clause_table_match.json"
        with open(clause_match_path, 'w', encoding='utf-8') as f:
            json.dump(clause_match, f, ensure_ascii=False, indent=2)

        if clause_match.get('status') == 'PASS':
            job.add_qa_result("qa_clause_match", "PASS", "條款表格一致")
        else:
            job.add_qa_result("qa_clause_match", "WARN",
                              f"問題: {len(clause_match.get('issues', []))}",
                              {"issues": clause_match.get('issues', [])})
            # Clause mismatch 目前設為 WARN

        # Gate 3: Final QA
        from tools.final_qa import (
            is_missing, check_docx_overview, check_docx_5522, check_docx_b25
        )

        final_issues = []

        # 基本檢查
        for k in ['cb_report_no', 'applicant', 'model_type_references', 'ratings_input']:
            if is_missing(meta.get(k)):
                final_issues.append({"type": "missing_meta", "field": k})

        if len(overview_cns) < 1:
            final_issues.append({"type": "overview_empty"})

        if len(clauses) < 20:
            final_issues.append({"type": "clauses_too_few", "count": len(clauses)})

        # Word 內容檢查
        docx_overview_check = check_docx_overview(docx_path)
        if not docx_overview_check['has_capacitor_row']:
            final_issues.append({"type": "docx_missing_capacitor_row"})

        if not docx_overview_check['has_5_5_2_in_es3']:
            final_issues.append({"type": "docx_es3_missing_5_5_2"})

        final_qa_report = {
            "status": "PASS" if len(final_issues) == 0 else "FAIL",
            "issues": final_issues,
            "stats": {
                "overview_rows": len(overview_cns),
                "clauses": len(clauses),
                "word_overview_rows": docx_overview_check['data_row_count']
            }
        }

        final_qa_path = out_dir / "final_qa_report.json"
        with open(final_qa_path, 'w', encoding='utf-8') as f:
            json.dump(final_qa_report, f, ensure_ascii=False, indent=2)

        if final_qa_report['status'] == 'PASS':
            job.add_qa_result("final_qa", "PASS", "所有檢查通過")
            all_pass = True
        else:
            job.add_qa_result("final_qa", "FAIL", f"{len(final_issues)} 個問題",
                              {"issues": final_issues})
            all_pass = False

        # ===== Step 6: 決定最終狀態並上傳 =====
        if all_pass:
            job.update_status(JobStatus.PASS)
            job.docx_type = "FINAL"
        else:
            job.update_status(JobStatus.FAIL)
            job.docx_type = "DRAFT"

        # 上傳結果
        job_prefix = f"jobs/{job.job_id}"

        # 上傳 DOCX
        docx_key = f"{job_prefix}/{job.docx_type.lower()}_cns_report.docx"
        storage.upload_file(str(docx_path), docx_key)
        job.docx_key = docx_key

        # 上傳 JSON 資料
        json_key = f"{job_prefix}/cns_report_data.json"
        storage.upload_file(str(cns_json_path), json_key)
        job.json_data_key = json_key

        # 上傳 QA 報告
        qa_report_key = f"{job_prefix}/qa_report.json"
        combined_qa = {
            "job_id": job.job_id,
            "status": job.status.value,
            "qa_results": [r.to_dict() for r in job.qa_results],
            "final_qa": final_qa_report,
            "overview_match": overview_match,
            "clause_match": clause_match,
            "timestamp": datetime.utcnow().isoformat()
        }
        storage.upload_json(combined_qa, qa_report_key)
        job.qa_report_key = qa_report_key

        # 上傳 clause rows
        clause_rows_key = f"{job_prefix}/pdf_clause_rows.json"
        storage.upload_file(str(clause_rows_path), clause_rows_key)
        job.pdf_clause_rows_key = clause_rows_key

    except Exception as e:
        job.update_status(JobStatus.ERROR)
        job.error_message = str(e)
        import traceback
        job.add_qa_result("pipeline_error", "ERROR", str(e),
                          {"traceback": traceback.format_exc()})

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
                     special_tables_path: str, output_path: str):
    """執行 render_word"""
    import subprocess
    result = subprocess.run([
        sys.executable,
        str(PROJECT_ROOT / "tools" / "render_word.py"),
        "--json", json_path,
        "--template", template_path,
        "--pdf_clause_rows", pdf_clause_rows_path,
        "--special_tables", special_tables_path,
        "--out", output_path
    ], capture_output=True, text=True, cwd=str(PROJECT_ROOT))

    if result.returncode != 0:
        raise RuntimeError(f"render_word failed: {result.stderr}")
