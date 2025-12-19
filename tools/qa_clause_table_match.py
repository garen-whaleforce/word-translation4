# tools/qa_clause_table_match.py
"""
QA 工具：驗證 PDF 主幹條款表與 Word 條款表一致性
確保：
1. 行數一致
2. 順序一致
3. clause_id 集合一致
4. PDF remark 為空時，Word remark 也必須為空
"""
import json
import argparse
from pathlib import Path
from docx import Document
import re


def load_json(p: Path) -> list:
    with p.open("r", encoding="utf-8") as f:
        return json.load(f)


def is_valid_clause_id(clause_id: str) -> bool:
    """檢查是否為有效的條款 ID"""
    if not clause_id:
        return False
    # 數字章節：4, 4.1, 4.1.1, 5, 5.1, ..., 10, 10.1, etc.
    # 字母章節：B, B.1, B.1.1, C, C.1, ..., M, M.1, etc.
    numeric_pattern = re.compile(r'^([4-9]|10)(\.[0-9]+)*$')
    letter_pattern = re.compile(r'^[BCDEFGHJKLM](\.[0-9]+)*$')
    return bool(numeric_pattern.match(clause_id) or letter_pattern.match(clause_id))


def extract_word_clause_rows(docx_path: Path) -> list:
    """從 Word 抽取主幹條款表資料"""
    rows = []

    if not docx_path.exists():
        return rows

    doc = Document(str(docx_path))

    # 主條款表格識別：
    # - 數字章節表格 (4-10): 第一個 cell 必須剛好是章節號 "4", "5", ..., "10"
    # - 字母章節表格: 第一個 cell 以章節字母開頭 (B.1, C.1, etc.)，
    #   但排除詳細表格（如 B.2.5, B.3, B.4 等多層次號碼）
    numeric_sections = ['4', '5', '6', '7', '8', '9', '10']
    letter_sections = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M']

    main_tables_found = set()  # 追蹤已找到的章節表格

    for tbl_idx, tbl in enumerate(doc.tables):
        if not tbl.rows:
            continue

        first_cell = tbl.rows[0].cells[0].text.strip() if tbl.rows[0].cells else ''

        # 檢查是否為主條款表格
        is_main_table = False
        section = None

        # 直接匹配數字章節開頭（必須剛好是單一數字）
        if first_cell in numeric_sections:
            section = first_cell
            if section not in main_tables_found:
                is_main_table = True
                main_tables_found.add(section)
        else:
            # 字母章節：只接受 "X.1" 格式作為 B-M 合併表格的開頭
            # B.1 表示 B~M 合併表格，我們要識別這個並將所有字母章節視為已找到
            # 排除 "B.2.5", "M.3" 等詳細表格
            if re.match(r'^B\.[0-9]+$', first_cell) and 'B' not in main_tables_found:
                # 這是 B~M 合併表格，標記所有字母章節為已找到
                is_main_table = True
                for sec in letter_sections:
                    main_tables_found.add(sec)

        if not is_main_table:
            continue

        for row in tbl.rows:
            if len(row.cells) < 4:
                continue

            clause_id = row.cells[0].text.strip()
            title = row.cells[1].text.strip()
            remark = row.cells[2].text.strip()
            verdict = row.cells[3].text.strip()

            # 跳過表頭（第一欄為標題文字，非條款編號）
            if clause_id in ['Clause', '條款', '章節']:
                continue

            rows.append({
                'clause_id': clause_id,
                'title': title[:50],
                'remark': remark,
                'verdict': verdict
            })

    return rows


def compare_clause_tables(pdf_rows: list, word_rows: list) -> dict:
    """比較 PDF 和 Word 的條款表"""
    result = {
        'pdf_row_count': len(pdf_rows),
        'word_row_count': len(word_rows),
        'issues': [],
        'status': 'PASS'
    }

    # 1. 行數檢查
    if len(pdf_rows) != len(word_rows):
        result['issues'].append({
            'type': 'row_count_mismatch',
            'detail': f'PDF 有 {len(pdf_rows)} 列，Word 有 {len(word_rows)} 列'
        })
        result['status'] = 'FAIL'

    # 2. clause_id 集合檢查
    pdf_clause_ids = set(r['clause_id'] for r in pdf_rows if r['clause_id'])
    word_clause_ids = set(r['clause_id'] for r in word_rows if r['clause_id'])

    pdf_only = pdf_clause_ids - word_clause_ids
    word_only = word_clause_ids - pdf_clause_ids

    if pdf_only:
        result['issues'].append({
            'type': 'pdf_only_clauses',
            'detail': f'{len(pdf_only)} 個條款只在 PDF 中',
            'samples': sorted(list(pdf_only))[:10]
        })
        result['status'] = 'FAIL'

    if word_only:
        result['issues'].append({
            'type': 'word_only_clauses',
            'detail': f'{len(word_only)} 個條款只在 Word 中（模板殘留）',
            'samples': sorted(list(word_only))[:10]
        })
        result['status'] = 'FAIL'

    # 3. 順序檢查（前 N 個條款）
    min_len = min(len(pdf_rows), len(word_rows), 50)
    order_mismatches = []
    for i in range(min_len):
        pdf_cid = pdf_rows[i].get('clause_id', '')
        word_cid = word_rows[i].get('clause_id', '')
        if pdf_cid != word_cid:
            order_mismatches.append({
                'index': i,
                'pdf': pdf_cid,
                'word': word_cid
            })

    if order_mismatches:
        result['issues'].append({
            'type': 'order_mismatch',
            'detail': f'{len(order_mismatches)} 個位置順序不一致',
            'samples': order_mismatches[:5]
        })
        result['status'] = 'FAIL'

    # 4. remark 一致性檢查（PDF 空則 Word 也須空）
    remark_violations = []
    pdf_by_cid = {r['clause_id']: r for r in pdf_rows if r['clause_id']}
    word_by_cid = {r['clause_id']: r for r in word_rows if r['clause_id']}

    for cid in pdf_by_cid:
        if cid in word_by_cid:
            pdf_remark = pdf_by_cid[cid].get('remark', '')
            word_remark = word_by_cid[cid].get('remark', '')

            # PDF 空但 Word 不空 => 模板殘留
            if not pdf_remark and word_remark and word_remark not in ['', '—', '-']:
                remark_violations.append({
                    'clause_id': cid,
                    'pdf_remark': '(空)',
                    'word_remark': word_remark[:30]
                })

    if remark_violations:
        result['issues'].append({
            'type': 'remark_template_residue',
            'detail': f'{len(remark_violations)} 個條款有模板殘留的 remark',
            'samples': remark_violations[:10]
        })
        result['status'] = 'FAIL'

    # 統計
    result['pdf_clause_ids'] = sorted(list(pdf_clause_ids))[:20]
    result['word_clause_ids'] = sorted(list(word_clause_ids))[:20]

    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf_rows", required=True, help="PDF 條款列 JSON")
    ap.add_argument("--docx", required=True, help="Word 輸出檔案")
    ap.add_argument("--out", required=True, help="輸出報告路徑")
    args = ap.parse_args()

    # 讀取資料
    pdf_rows = load_json(Path(args.pdf_rows))
    word_rows = extract_word_clause_rows(Path(args.docx))

    # 比較
    result = compare_clause_tables(pdf_rows, word_rows)

    # 輸出
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print("=" * 50)
    print(f"Clause Table Match QA: {result['status']}")
    print("=" * 50)
    print(f"\nPDF 條款列數: {result['pdf_row_count']}")
    print(f"Word 條款列數: {result['word_row_count']}")

    if result['issues']:
        print(f"\n發現 {len(result['issues'])} 個問題:")
        for issue in result['issues']:
            print(f"  - [{issue['type']}] {issue['detail']}")
            if 'samples' in issue:
                for s in issue['samples'][:3]:
                    print(f"      {s}")
    else:
        print("\nPDF 與 Word 條款表一致!")

    print(f"\n詳細報告: {out_path}")

    if result['status'] != 'PASS':
        raise SystemExit(2)


if __name__ == "__main__":
    main()
