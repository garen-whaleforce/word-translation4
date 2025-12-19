# tools/generate_cns_json.py
"""
從 CB PDF 抽取的原始資料生成 CNS 15598-1 JSON
使用 cb_overview_raw.json 和 cb_clauses_raw.json
"""
import json
import re
import argparse
from pathlib import Path

def load_json(path: Path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def extract_meta_from_chunks(chunks: list, pdf_name: str) -> dict:
    """從文字 chunks 中抽取 meta 資訊"""
    meta = {
        "source_pdf_name": pdf_name,
        "standard": "IEC 62368-1:2018",
        "target_report": "CNS 15598-1 (109年版)",
        "cb_report_no": "",
        "report_date": "",
        "applicant": "",
        "manufacturer": "",
        "factory_locations": [],
        "trade_mark": "",
        "model_type_references": [],
        "ratings_input": "",
        "ratings_output": "",
        "notes": []
    }

    # 合併前幾頁文字
    first_pages_text = '\n'.join([c['text'] for c in chunks[:15]])

    # Report No - 多種格式
    m = re.search(r'Report\s*Number[.\s]*:\s*([A-Z0-9]+\s*\d+)', first_pages_text, re.IGNORECASE)
    if m:
        meta['cb_report_no'] = m.group(1).strip()
    else:
        m = re.search(r'Report\s*No\.?\s*[:\s]*([A-Z0-9]{2,}\s*\d+)', first_pages_text, re.IGNORECASE)
        if m:
            meta['cb_report_no'] = m.group(1).strip()

    # Date
    m = re.search(r'Date\s*of\s*issue\s*[.\s]*:\s*(\d{4}[-/]\d{2}[-/]\d{2})', first_pages_text, re.IGNORECASE)
    if m:
        meta['report_date'] = m.group(1)

    # Applicant - 多種格式
    m = re.search(r"Applicant.*?:\s*([A-Z][^\n]+)", first_pages_text, re.IGNORECASE)
    if m:
        meta['applicant'] = m.group(1).strip()

    # Manufacturer - 多種格式
    m = re.search(r'Manufacturer\s*[.\s]*:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        mfr = m.group(1).strip()
        if 'same as' in mfr.lower() or 'see above' in mfr.lower():
            meta['manufacturer'] = 'Same as applicant'
        else:
            meta['manufacturer'] = mfr

    # Model - 多種格式
    m = re.search(r'Model/Type\s*reference\s*[.\s]*:\s*([A-Z0-9][\w\-]+(?:\s*,\s*[A-Z0-9][\w\-]+)?)', first_pages_text, re.IGNORECASE)
    if m:
        models = m.group(1).strip()
        # 分割多個型號（只用逗號分隔）
        model_list = re.split(r'\s*,\s*', models)
        meta['model_type_references'] = [m.strip() for m in model_list if m.strip()]
    else:
        m = re.search(r'Model[/\s]*Type\s*Ref[:\s]*([^\n]+)', first_pages_text, re.IGNORECASE)
        if m:
            models = m.group(1).strip()
            model_list = re.split(r'\s*,\s*', models)
            meta['model_type_references'] = [m.strip() for m in model_list if m.strip()]

    # Ratings - 多種格式
    m = re.search(r'Ratings\s*[.\s]*:\s*Input:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        meta['ratings_input'] = m.group(1).strip()
    else:
        m = re.search(r'Rated\s*input[:\s]*([^\n]+)', first_pages_text, re.IGNORECASE)
        if m:
            meta['ratings_input'] = m.group(1).strip()
        else:
            m = re.search(r'Input:\s*([0-9\-]+V[^\n]+)', first_pages_text, re.IGNORECASE)
            if m:
                meta['ratings_input'] = m.group(1).strip()

    # Ratings Output
    m = re.search(r'Output:\s*([^\n]+)', first_pages_text, re.IGNORECASE)
    if m:
        meta['ratings_output'] = m.group(1).strip()
    else:
        m = re.search(r'Rated\s*output[:\s]*([^\n]+)', first_pages_text, re.IGNORECASE)
        if m:
            meta['ratings_output'] = m.group(1).strip()

    return meta

def convert_overview_to_cns(overview_raw: list) -> list:
    """將原始 overview 資料轉換為 CNS 格式"""
    result = []

    clause_map = {
        '5': 'Clause 5 Electrically-caused injury',
        '6': 'Clause 6 Electrically-caused fire',
        '7': 'Clause 7 Injury caused by hazardous substances',
        '8': 'Clause 8 Mechanically-caused injury',
        '9': 'Clause 9 Thermal burn',
        '10': 'Clause 10 Radiation'
    }

    for item in overview_raw:
        clause = item.get('clause', '')
        row = item.get('row', [])

        if len(row) < 5:
            continue

        energy_source = row[0].replace('\n', ' ').strip()
        parts_involved = row[1].replace('\n', ' ').strip()

        # 組合 safeguards (B, S, R 或 B, 1st S, 2nd S)
        safeguards_parts = []
        for i, label in enumerate(['B', 'S', 'R']):
            if i + 2 < len(row) and row[i + 2] and row[i + 2] != 'N/A':
                val = row[i + 2].replace('\n', ' ').strip()
                if val and val != 'N/A':
                    safeguards_parts.append(f"{label}: {val}")

        safeguards = ', '.join(safeguards_parts) if safeguards_parts else 'N/A'

        # 抽取 energy source class (ES3, PS2, MS1 等)
        energy_class = ''
        m = re.match(r'^(ES[123]|PS[123]|MS[123]|TS[123]|RS[123]|N/A)', energy_source)
        if m:
            energy_class = m.group(1)

        result.append({
            'energy_source_class': energy_class,
            'parts_involved': energy_source,
            'safeguards': safeguards,
            'remarks_or_clause_ref': clause_map.get(clause, f'Clause {clause}'),
            'evidence_quote': ' '.join([c.replace('\n', ' ') for c in row])
        })

    return result

def dedupe_clauses(clauses_raw: list) -> list:
    """去重 clause（保留第一個出現的）"""
    seen = set()
    result = []
    for c in clauses_raw:
        cid = c.get('clause_id', '')
        if cid and cid not in seen:
            seen.add(cid)
            result.append(c)
    return result

def generate_qa(meta: dict, overview: list, clauses: list, pdf_overview_raw: list) -> dict:
    """生成 QA 報告"""
    issues = []

    # 必填欄位檢查
    required_meta = ['cb_report_no', 'applicant', 'model_type_references', 'ratings_input']
    for field in required_meta:
        val = meta.get(field)
        if not val or val == 'MISSING' or (isinstance(val, list) and len(val) == 0):
            issues.append(f"meta.{field} is missing or empty")

    # Overview 列數檢查
    if len(overview) < 1:
        issues.append("overview_energy_sources_and_safeguards has no rows")

    # Clause 數量檢查
    if len(clauses) < 20:
        issues.append(f"clauses count is {len(clauses)}, expected >= 20")

    # 特殊檢查：Capacitor row
    has_capacitor_in_pdf = any('Capacitor connected between L and N' in str(r) for r in pdf_overview_raw)
    has_capacitor_in_json = any('Capacitor connected between L and N' in str(o) for o in overview)

    if has_capacitor_in_pdf and not has_capacitor_in_json:
        issues.append("PDF contains 'Capacitor connected between L and N' but JSON overview is missing it")

    # ES3 計數檢查
    pdf_es3_count = sum(1 for r in pdf_overview_raw if 'ES3:' in str(r.get('row', [])))
    json_es3_count = sum(1 for o in overview if o.get('energy_source_class') == 'ES3')
    if pdf_es3_count > json_es3_count:
        issues.append(f"PDF has {pdf_es3_count} ES3 rows but JSON has only {json_es3_count}")

    # 第一列 ES3 的 safeguards 必須包含 5.5.2
    first_es3_safeguards = ''
    for r in pdf_overview_raw:
        row = r.get('row', [])
        if row and 'ES3:' in str(row[0]) and 'Primary' in str(row[0]):
            first_es3_safeguards = str(row[-1]) if row else ''
            break

    if '5.5.2' in first_es3_safeguards:
        first_json_es3 = next((o for o in overview if o.get('energy_source_class') == 'ES3'), None)
        if first_json_es3 and '5.5.2' not in str(first_json_es3):
            issues.append("First ES3 row safeguards should contain '5.5.2' but it's missing in JSON")

    status = 'PASS' if not issues else 'FAIL'

    return {
        'summary': {
            'status': status,
            'overview_rows': len(overview),
            'clauses_count': len(clauses),
            'issues_count': len(issues)
        },
        'issues': issues
    }

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input_dir", required=True, help="包含 cb_*.json 的目錄")
    ap.add_argument("--pdf_name", required=True, help="PDF 檔名")
    ap.add_argument("--out", required=True, help="輸出 JSON 路徑")
    ap.add_argument("--special_tables", default=None, help="特殊表格 JSON 路徑 (cb_special_tables.json)")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)

    # 讀取原始資料
    chunks = load_json(input_dir / "cb_text_chunks.json")
    overview_raw = load_json(input_dir / "cb_overview_raw.json")
    clauses_raw = load_json(input_dir / "cb_clauses_raw.json")

    # 讀取特殊表格（如果有）
    special_tables = {}
    special_tables_path = Path(args.special_tables) if args.special_tables else input_dir / "cb_special_tables.json"
    if special_tables_path.exists():
        special_tables = load_json(special_tables_path)

    # 生成各區塊
    meta = extract_meta_from_chunks(chunks, args.pdf_name)
    overview = convert_overview_to_cns(overview_raw)
    clauses = dedupe_clauses(clauses_raw)

    # 從特殊表格抽取 overview_cb_p12_rows（保留完整 10 列）
    overview_cb_p12_rows = []
    if 'overview' in special_tables and 'rows' in special_tables['overview']:
        overview_cb_p12_rows = special_tables['overview']['rows']

    # QA
    qa = generate_qa(meta, overview, clauses, overview_raw)

    # 更新 QA：overview_cb_p12_rows 列數檢查
    if len(overview_cb_p12_rows) != 10:
        qa['issues'].append(f"overview_cb_p12_rows should have 10 rows but has {len(overview_cb_p12_rows)}")
        qa['summary']['status'] = 'FAIL'
        qa['summary']['issues_count'] = len(qa['issues'])

    # 組合最終 JSON
    result = {
        'meta': meta,
        'overview_energy_sources_and_safeguards': overview,
        'overview_cb_p12_rows': overview_cb_p12_rows,
        'clauses': clauses,
        'attachments_or_annex': [],
        'qa': qa
    }

    # 輸出
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Generated: {out_path}")
    print(f"QA Status: {qa['summary']['status']}")
    print(f"overview_cb_p12_rows: {len(overview_cb_p12_rows)} rows")
    if qa['issues']:
        for issue in qa['issues']:
            print(f"  - {issue}")

if __name__ == "__main__":
    main()
