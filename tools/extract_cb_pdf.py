# tools/extract_cb_pdf.py
import json, re, argparse
from pathlib import Path
import pdfplumber

def norm(s: str) -> str:
    s = s or ""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def find_overview_page(pdf) -> int:
    """找出 OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS 所在頁 index"""
    for i, page in enumerate(pdf.pages[:30]):
        text = (page.extract_text() or '').upper()
        if 'OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS' in text:
            return i
    return -1

def extract_overview_table(page) -> list:
    """使用 lines-based 策略抽取 Overview 表格，確保不漏列"""
    tables = page.extract_tables({
        'vertical_strategy': 'lines',
        'horizontal_strategy': 'lines',
        'intersection_tolerance': 5,
        'snap_tolerance': 5,
        'join_tolerance': 5,
    })

    if not tables:
        return []

    # 找最大的表格（Overview 表通常是該頁最大的）
    main_table = max(tables, key=lambda t: len(t))

    overview_rows = []
    # 識別資料列：包含 ES/PS/MS/TS/RS/N/A 開頭的列
    energy_pattern = re.compile(r'^(ES[123]|PS[123]|MS[123]|TS[123]|RS[123]|N/A)\s*:', re.IGNORECASE)
    na_only_pattern = re.compile(r'^N/A$', re.IGNORECASE)

    current_clause = ""
    for row in main_table:
        if not row or not row[0]:
            continue

        first_cell = norm(row[0])

        # 記錄當前 clause (5, 6, 7, 8, 9, 10)
        if re.match(r'^[5-9]$|^10$', first_cell):
            current_clause = first_cell
            continue

        # 跳過表頭列
        if 'Class and Energy Source' in first_cell or 'Clause' in first_cell:
            continue
        if '(e.g.' in first_cell:
            continue
        if first_cell == 'OVERVIEW':
            continue

        # 資料列：ES3, PS2, MS1 等開頭
        if energy_pattern.match(first_cell):
            # 正常資料列
            overview_rows.append({
                'clause': current_clause,
                'row': [norm(c) if c else '' for c in row]
            })
        elif na_only_pattern.match(first_cell) and len(row) >= 5:
            # N/A 開頭的列（如 Clause 7, 10）
            all_na = all(norm(c or '') in ['N/A', ''] for c in row)
            if all_na:
                overview_rows.append({
                    'clause': current_clause,
                    'row': [norm(c) if c else '' for c in row]
                })

    return overview_rows

def find_clause_pages(pdf) -> tuple:
    """找出 Clause 表開始和結束頁"""
    start_idx = -1
    for i, page in enumerate(pdf.pages):
        text = (page.extract_text() or '')
        if 'Clause Requirement + Test Result - Remark Verdict' in text:
            start_idx = i
            break
        if re.search(r'Clause\s+Requirement.*Verdict', text, re.IGNORECASE):
            start_idx = i
            break
    return start_idx

def find_table_412_pages(tables: list) -> tuple:
    """找出 4.1.2 Critical components 表格的頁碼範圍"""
    start_page = None
    end_page = None

    for tbl in tables:
        page = tbl.get('page')
        rows = tbl.get('rows', [])

        if not rows:
            continue

        first_row_text = str(rows[0])

        # 找起始頁：含有 "4.1.2" 和 "Critical components" 或 "TABLE:"
        if '4.1.2' in first_row_text and ('Critical' in first_row_text or 'TABLE:' in first_row_text):
            start_page = page

        # 找結束頁：在 4.1.2 表格之後，出現新的條款編號表格
        if start_page and page > start_page:
            # 檢查是否是其他條款的 TABLE (如 5.2, 5.4.1.4 等)
            for row in rows[:2]:
                row_text = str(row)
                # 如果出現其他條款編號的 TABLE，表示 4.1.2 結束
                if 'TABLE:' in row_text and '4.1.2' not in row_text:
                    if not end_page:
                        end_page = page - 1
                    break

    # 如果沒找到結束頁，假設延續到 start_page + 10
    if start_page and not end_page:
        end_page = start_page + 10

    return start_page, end_page


def extract_table_412(tables: list) -> list:
    """
    從 cb_tables_text.json 提取 4.1.2 Critical components information 表格

    Returns:
        list of dict: 每個 dict 包含 part, manufacturer, model, spec, standard, mark
    """
    start_page, end_page = find_table_412_pages(tables)

    if not start_page:
        return []

    component_rows = []

    for tbl in tables:
        page = tbl.get('page')
        rows = tbl.get('rows', [])

        if not page or page < start_page or page > end_page:
            continue

        # 跳過頁眉表格 (IEC 62368-1, Clause...)
        if rows and rows[0] and 'IEC 62368' in str(rows[0]):
            continue

        for row in rows:
            if not row:
                continue

            # 跳過表頭行
            first_cell = norm(row[0] or '')
            if first_cell in ['4.1.2', 'Object / part No.', 'Clause', ''] or 'TABLE:' in first_cell:
                continue
            if 'trademark' in first_cell.lower() or 'Manufacturer' in first_cell:
                continue

            # 根據欄位數量調整索引
            if len(row) == 8:  # Page 70 格式 (有空欄)
                component_rows.append({
                    'part': norm(row[0] or ''),
                    'manufacturer': norm(row[2] or ''),
                    'model': norm(row[3] or ''),
                    'spec': norm(row[4] or ''),
                    'standard': norm(row[5] or ''),
                    'mark': norm(row[6] or ''),
                })
            elif len(row) >= 6:  # Page 71+ 格式
                component_rows.append({
                    'part': norm(row[0] or ''),
                    'manufacturer': norm(row[1] or ''),
                    'model': norm(row[2] or ''),
                    'spec': norm(row[3] or ''),
                    'standard': norm(row[4] or ''),
                    'mark': norm(row[5] or ''),
                })

    return component_rows


def extract_clauses_from_pages(pdf, start_idx: int) -> list:
    """從指定頁開始抽取所有條款"""
    if start_idx < 0:
        return []

    clauses = []
    # clause_id pattern: 4, 4.1.1, 5.2.2.2, B.2.5, Q.1, T.2 等
    clause_id_pattern = re.compile(r'^([4-9]|10|[A-Z])(\.[0-9]+)*$')

    for page in pdf.pages[start_idx:]:
        tables = page.extract_tables({
            'vertical_strategy': 'lines',
            'horizontal_strategy': 'lines',
            'intersection_tolerance': 3,
            'snap_tolerance': 3,
            'join_tolerance': 3,
        })

        for tbl in tables:
            for row in tbl:
                if not row or len(row) < 2:
                    continue

                first_cell = norm(row[0] or '')

                # 檢查是否為 clause_id
                if clause_id_pattern.match(first_cell):
                    clause_id = first_cell

                    # 取得 title, remark, verdict
                    title = norm(row[1] or '') if len(row) > 1 else ''
                    remark = norm(row[2] or '') if len(row) > 2 else ''
                    verdict_raw = norm(row[-1] or '') if row else ''

                    # verdict 標準化
                    verdict = ''
                    if verdict_raw.upper() in ['P', 'PASS', '符合']:
                        verdict = 'P'
                    elif verdict_raw.upper() in ['F', 'FAIL', '不符合']:
                        verdict = 'F'
                    elif verdict_raw.upper() in ['N/A', 'NA', '不適用']:
                        verdict = 'N/A'
                    elif verdict_raw.upper() in ['N', 'N.A.']:
                        verdict = 'N/A'
                    else:
                        verdict = verdict_raw

                    clauses.append({
                        'clause_id': clause_id,
                        'clause_title': title,
                        'test_result_or_remark': remark,
                        'verdict': verdict,
                        'evidence_quote': ' '.join([norm(c or '') for c in row])
                    })

    return clauses

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out_dir", required=True)
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    chunks = []
    tables = []
    overview_data = []
    clauses_data = []

    with pdfplumber.open(args.pdf) as pdf:
        # 1. 抽取所有頁面文字
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

        # 2. 專門抽取 Overview 表（使用 lines-based 策略）
        overview_page_idx = find_overview_page(pdf)
        if overview_page_idx >= 0:
            overview_page = pdf.pages[overview_page_idx]
            overview_data = extract_overview_table(overview_page)
            print(f"Overview 頁: {overview_page_idx + 1}, 抽取 {len(overview_data)} 列")

        # 3. 抽取 Clause 表
        clause_start_idx = find_clause_pages(pdf)
        if clause_start_idx >= 0:
            clauses_data = extract_clauses_from_pages(pdf, clause_start_idx)
            print(f"Clause 起始頁: {clause_start_idx + 1}, 抽取 {len(clauses_data)} 條")

    # 4. 抽取 4.1.2 Critical components 表格
    table_412_data = extract_table_412(tables)
    if table_412_data:
        print(f"4.1.2 表格: 抽取 {len(table_412_data)} 列零件資料")

    # 輸出檔案
    out_chunks = out_dir / "cb_text_chunks.json"
    out_tables = out_dir / "cb_tables_text.json"
    out_overview = out_dir / "cb_overview_raw.json"
    out_clauses = out_dir / "cb_clauses_raw.json"
    out_table_412 = out_dir / "cb_table_412.json"

    with open(out_chunks, "w", encoding="utf-8") as f:
        json.dump(chunks, f, ensure_ascii=False, indent=2)

    with open(out_tables, "w", encoding="utf-8") as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)

    with open(out_overview, "w", encoding="utf-8") as f:
        json.dump(overview_data, f, ensure_ascii=False, indent=2)

    with open(out_clauses, "w", encoding="utf-8") as f:
        json.dump(clauses_data, f, ensure_ascii=False, indent=2)

    with open(out_table_412, "w", encoding="utf-8") as f:
        json.dump(table_412_data, f, ensure_ascii=False, indent=2)

    print("OK")
    print("cb_text_chunks:", out_chunks)
    print("cb_tables_text:", out_tables)
    print("cb_overview_raw:", out_overview)
    print("cb_clauses_raw:", out_clauses)
    print("cb_table_412:", out_table_412)

if __name__ == "__main__":
    main()
