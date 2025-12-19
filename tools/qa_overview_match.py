# tools/qa_overview_match.py
"""
QA 工具：比較 PDF Overview 與 Word 安全防護總攬表
輸出詳細的對比報告 overview_match_report.json
"""
import json
import argparse
from pathlib import Path
from docx import Document


def load_json(p: Path) -> dict:
    with p.open("r", encoding="utf-8") as f:
        return json.load(f)


def extract_word_overview_rows(docx_path: Path) -> list:
    """從 Word 安全防護總攬表抽取資料列"""
    rows = []

    if not docx_path.exists():
        return rows

    doc = Document(str(docx_path))

    for tbl in doc.tables:
        if not tbl.rows:
            continue
        first_cell = tbl.rows[0].cells[0].text if tbl.rows[0].cells else ''

        if '安全防護總攬表' in first_cell:
            current_clause = None
            for row in tbl.rows:
                first_cell_text = row.cells[0].text.strip() if row.cells else ''

                # 追蹤章節
                if first_cell_text in ['5.1', '6.1', '7.1', '8.1', '9.1', '10.1']:
                    current_clause = first_cell_text
                    continue

                # 偵測資料列
                if (first_cell_text.startswith('ES') or
                    first_cell_text.startswith('PS') or
                    first_cell_text.startswith('MS') or
                    first_cell_text.startswith('TS') or
                    first_cell_text in ['N/A', '無']):

                    row_data = {
                        'cns_clause': current_clause,
                        'energy_source': first_cell_text,
                        'body_part': row.cells[1].text.strip() if len(row.cells) > 1 else '',
                        'safeguard_b': row.cells[2].text.strip() if len(row.cells) > 2 else '',
                        'safeguard_s': row.cells[3].text.strip() if len(row.cells) > 3 else '',
                        'safeguard_r': row.cells[4].text.strip() if len(row.cells) > 4 else '',
                    }
                    rows.append(row_data)

    return rows


def compare_rows(pdf_rows: list, word_rows: list) -> dict:
    """比較 PDF 和 Word 的列"""
    result = {
        'pdf_row_count': len(pdf_rows),
        'word_row_count': len(word_rows),
        'match': len(pdf_rows) == len(word_rows) == 10,
        'differences': [],
        'pdf_rows': [],
        'word_rows': []
    }

    # CB Clause 到 CNS 章節的映射
    cb_to_cns = {5: '5.1', 6: '6.1', 7: '7.1', 8: '8.1', 9: '9.1', 10: '10.1'}

    # 格式化 PDF rows
    for i, row in enumerate(pdf_rows):
        cb_clause = row.get('cb_clause', 0)
        result['pdf_rows'].append({
            'index': i,
            'cb_clause': cb_clause,
            'cns_clause': cb_to_cns.get(cb_clause, '?'),
            'energy_source': row.get('class_energy_source', '').replace('\n', ' ')[:50],
            'body_part': row.get('body_or_material', '').replace('\n', ' ')[:30],
            'safeguards': f"B:{row.get('basic', 'N/A')} S:{row.get('supp1', 'N/A')} R:{row.get('supp2', 'N/A')}"[:60]
        })

    # 格式化 Word rows
    for i, row in enumerate(word_rows):
        result['word_rows'].append({
            'index': i,
            'cns_clause': row.get('cns_clause', '?'),
            'energy_source': row.get('energy_source', '')[:50],
            'body_part': row.get('body_part', '')[:30],
            'safeguards': f"B:{row.get('safeguard_b', '')} S:{row.get('safeguard_s', '')} R:{row.get('safeguard_r', '')}"[:60]
        })

    # 找差異
    if len(pdf_rows) != len(word_rows):
        result['differences'].append({
            'type': 'row_count_mismatch',
            'detail': f'PDF has {len(pdf_rows)} rows, Word has {len(word_rows)} rows'
        })
        result['match'] = False

    # Clause 5 特別檢查
    pdf_clause5 = [r for r in pdf_rows if r.get('cb_clause') == 5]
    word_clause5 = [r for r in word_rows if r.get('cns_clause') == '5.1']

    if len(pdf_clause5) != len(word_clause5):
        result['differences'].append({
            'type': 'clause_5_mismatch',
            'detail': f'PDF Clause 5 has {len(pdf_clause5)} rows, Word has {len(word_clause5)} rows'
        })
        result['match'] = False

    # 檢查必要列
    # Capacitor 列
    pdf_has_cap = any('Capacitor' in str(r.get('class_energy_source', '')) for r in pdf_rows)
    word_has_cap = any('X電容' in r.get('energy_source', '') or 'Capacitor' in r.get('energy_source', '') for r in word_rows)

    if pdf_has_cap and not word_has_cap:
        result['differences'].append({
            'type': 'missing_capacitor_in_word',
            'detail': 'PDF has Capacitor row but Word does not'
        })
        result['match'] = False

    # ES1 輸出列
    pdf_has_es1 = any('ES1:' in str(r.get('class_energy_source', '')) and 'Secondary' in str(r.get('class_energy_source', '')) for r in pdf_rows)
    word_has_es1 = any('ES1' in r.get('energy_source', '') and ('輸出' in r.get('energy_source', '') or 'output' in r.get('energy_source', '').lower()) for r in word_rows)

    if pdf_has_es1 and not word_has_es1:
        result['differences'].append({
            'type': 'missing_es1_output_in_word',
            'detail': 'PDF has ES1 Secondary output row but Word does not'
        })
        result['match'] = False

    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True, help="cns_report_data.json 路徑")
    ap.add_argument("--docx", required=True, help="輸出的 Word 檔案")
    ap.add_argument("--out", required=True, help="輸出報告路徑")
    args = ap.parse_args()

    # 讀取資料
    data = load_json(Path(args.json))
    pdf_rows = data.get('overview_cb_p12_rows', [])

    # 抽取 Word 資料
    word_rows = extract_word_overview_rows(Path(args.docx))

    # 比較
    result = compare_rows(pdf_rows, word_rows)

    # 判定
    status = 'PASS' if result['match'] and len(result['differences']) == 0 else 'FAIL'
    result['status'] = status

    # 輸出
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print("=" * 50)
    print(f"Overview Match QA: {status}")
    print("=" * 50)
    print(f"\nPDF rows: {result['pdf_row_count']}")
    print(f"Word rows: {result['word_row_count']}")

    if result['differences']:
        print(f"\n發現 {len(result['differences'])} 個差異:")
        for d in result['differences']:
            print(f"  - [{d['type']}] {d['detail']}")
    else:
        print("\nPDF 與 Word 安全防護總攬表一致!")

    print(f"\n詳細報告: {out_path}")

    if status != 'PASS':
        raise SystemExit(2)


if __name__ == "__main__":
    main()
