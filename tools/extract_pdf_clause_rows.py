# tools/extract_pdf_clause_rows.py
"""
從 CB PDF 提取完整的條款列（包含子項目如 5.3.1 a), 5.3.2.2 b) 等）
輸出格式：
[
  {"clause_id": "4", "req": "GENERAL REQUIREMENTS", "remark": "", "verdict": "P", "pdf_page": 12},
  {"clause_id": "5.3.1 a)", "req": "Accessible ES1/ES2...", ...},
  {"clause_id": "", "req": "...", ..., "parent_clause": "5.3.2.2"},  # 無 clause_id 的子項
  ...
]
"""
import json
import re
import argparse
from pathlib import Path
import pdfplumber


def norm(s: str) -> str:
    """正規化字串"""
    s = s or ""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def is_valid_clause_id(cid: str) -> bool:
    """
    檢查是否為有效的條款 ID
    支援：4, 4.1.1, 5.3.1 a), B.2.5, M.3, N.1, O.2, P.3, Q.1, R.2, S.1, T.2, U.1, V.3, W.1, X.2, Y.3 等
    """
    if not cid:
        return False

    # 標準格式: 4, 4.1, 4.1.1, 5.3.1, 10.2.3, B.1, M.3, N.1, O.2, P.3, Q.1, R.2, S.1, T.2, U.1, V.3, W.1, X.2, Y.3 等
    # 附錄字母：B~M (主條款), N~Y (額外附錄，排除 I 和 O 因為容易與數字混淆，但 O 在某些報告中使用)
    standard_pattern = re.compile(r'^([4-9]|10|[BCDEFGHJKLMNOPQRSTUVWXY])(\.[0-9]+)*$')

    # 帶有子項目的格式: 5.3.1 a), 5.3.2.2 b), G.4.5 c) 等
    subitem_pattern = re.compile(r'^([4-9]|10|[BCDEFGHJKLMNOPQRSTUVWXY])(\.[0-9]+)*\s+[a-z]\)$')

    return bool(standard_pattern.match(cid) or subitem_pattern.match(cid))


def normalize_verdict(verdict_raw: str) -> str:
    """標準化 verdict 欄位"""
    v = (verdict_raw or '').strip().upper()
    if v in ['P', 'PASS', '符合']:
        return 'P'
    elif v in ['F', 'FAIL', '不符合']:
        return 'F'
    elif v in ['N/A', 'NA', '不適用', 'N.A.']:
        return 'N/A'
    elif v in ['—', '-', '⎯', '']:
        return '⎯'
    return verdict_raw or ''


def extract_clause_rows(pdf, start_page: int = 0) -> list:
    """
    從 PDF 提取所有條款列

    Args:
        pdf: pdfplumber PDF 物件
        start_page: 開始頁碼 (0-indexed)

    Returns:
        list of dict: 每個 dict 包含 clause_id, req, remark, verdict, pdf_page, parent_clause
    """
    rows = []
    current_parent = ""

    # 條款表格的識別模式（第一列必須是 Clause 或條款編號）
    header_pattern = re.compile(r'(Clause|IEC\s+\d+|條款)', re.IGNORECASE)

    for page_idx in range(start_page, len(pdf.pages)):
        page = pdf.pages[page_idx]
        page_num = page_idx + 1  # 1-indexed

        try:
            tables = page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
                'intersection_tolerance': 3,
                'snap_tolerance': 3,
                'join_tolerance': 3,
            })
        except Exception:
            continue

        for tbl in tables:
            if not tbl:
                continue

            for row in tbl:
                if not row or len(row) < 2:
                    continue

                first_cell = norm(row[0] or '')

                # 跳過表頭列
                if header_pattern.match(first_cell):
                    continue
                if first_cell in ['Clause', 'Requirement + Test', 'Result - Remark', 'Verdict']:
                    continue

                # 提取各欄位
                req = norm(row[1] or '') if len(row) > 1 else ''
                remark = norm(row[2] or '') if len(row) > 2 else ''
                verdict_raw = norm(row[-1] or '') if row else ''
                verdict = normalize_verdict(verdict_raw)

                # 判斷 clause_id
                if is_valid_clause_id(first_cell):
                    clause_id = first_cell
                    current_parent = clause_id.split()[0]  # 去掉 a), b) 後綴作為 parent

                    rows.append({
                        'clause_id': clause_id,
                        'req': req,
                        'remark': remark,
                        'verdict': verdict,
                        'pdf_page': page_num
                    })
                elif first_cell == '' and not req and remark:
                    # 情況：remark 跨行延續（如 "9.3, B.1.5, B.2.6)"）
                    # 前一行可能有未完成的 remark，需要合併
                    if rows and rows[-1].get('remark', '').rstrip().endswith(','):
                        # 前一行 remark 以逗號結尾，合併當前行的 remark
                        rows[-1]['remark'] = rows[-1]['remark'] + ' ' + remark
                    elif rows and '(' in rows[-1].get('remark', '') and ')' not in rows[-1].get('remark', ''):
                        # 前一行 remark 有未閉合的括號，合併當前行
                        rows[-1]['remark'] = rows[-1]['remark'] + ' ' + remark
                elif first_cell == '' and req:
                    req_clean = req.replace('\n', ' ').strip()

                    # 檢查是否是需要合併到前一行的延續行
                    # 情況 1: 標題斷行（全大寫英文且 verdict 是 ⎯）
                    is_title_continuation = (
                        rows and
                        verdict == '⎯' and
                        req_clean.isupper() and
                        re.match(r'^[A-Z\s,]+$', req_clean)
                    )

                    # 情況 2: 描述性延續行（verdict 是 ⎯ 且以冒號結尾或包含 "...." 省略號）
                    # 這些通常是表格欄位標題的延續，如 "for contact gaps (mm) ...:"
                    is_description_continuation = (
                        rows and
                        verdict == '⎯' and
                        (req_clean.endswith(':') or
                         re.search(r'\.{3,}', req_clean) or  # 省略號 "....."
                         req_clean.lower() in ['circuit elements', 'parts'])  # 常見的單獨標題詞
                    )

                    # 情況 3: 前一行的 verdict 有值（如 N/A, P），當前行 verdict 為 ⎯
                    # 且當前行看起來是前一行描述的一部分
                    prev_has_verdict = (
                        rows and
                        rows[-1].get('verdict') in ['N/A', 'P', 'F'] and
                        verdict == '⎯'
                    )

                    if is_title_continuation or is_description_continuation:
                        # 合併到前一行的 req
                        rows[-1]['req'] = rows[-1]['req'] + ' ' + req_clean
                    elif prev_has_verdict and len(req_clean) < 50:
                        # 短描述且前一行有明確 verdict，合併
                        rows[-1]['req'] = rows[-1]['req'] + ' ' + req_clean
                    else:
                        # 無 clause_id 的子項目
                        rows.append({
                            'clause_id': '',
                            'req': req,
                            'remark': remark,
                            'verdict': verdict,
                            'pdf_page': page_num,
                            'parent_clause': current_parent
                        })

    return rows


def find_clause_start_page(pdf) -> int:
    """找出條款表格開始的頁碼 (0-indexed)"""
    for i, page in enumerate(pdf.pages[:30]):
        text = (page.extract_text() or '')
        # 找 "Clause Requirement + Test Result - Remark Verdict" 或類似表頭
        if 'Clause' in text and 'Requirement' in text and 'Verdict' in text:
            # 確認這頁有 "4" 或 "4.1" 這樣的條款編號
            if re.search(r'\n4\s+', text) or re.search(r'\n4\.1', text):
                return i
    return 10  # 預設從第 11 頁開始


def main():
    ap = argparse.ArgumentParser(description='從 CB PDF 提取完整條款列')
    ap.add_argument("--pdf", required=True, help="CB PDF 路徑")
    ap.add_argument("--out", required=True, help="輸出 JSON 路徑")
    ap.add_argument("--start_page", type=int, default=None, help="起始頁碼 (1-indexed)")
    args = ap.parse_args()

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with pdfplumber.open(args.pdf) as pdf:
        if args.start_page:
            start_idx = args.start_page - 1  # 轉為 0-indexed
        else:
            start_idx = find_clause_start_page(pdf)

        print(f"從第 {start_idx + 1} 頁開始提取...")
        rows = extract_clause_rows(pdf, start_idx)
        print(f"共提取 {len(rows)} 列")

        # 統計
        with_cid = sum(1 for r in rows if r['clause_id'])
        without_cid = len(rows) - with_cid
        print(f"  - 有 clause_id: {with_cid} 列")
        print(f"  - 無 clause_id (子項目): {without_cid} 列")

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)

    print(f"輸出: {out_path}")


if __name__ == "__main__":
    main()
