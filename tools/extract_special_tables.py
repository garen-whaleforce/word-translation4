# tools/extract_special_tables.py
"""
專用抽取器：抽取 CB PDF 中的特殊表格
1. Overview of Energy Sources and Safeguards (p.12)
2. 5.5.2.2 Stored discharge on capacitors
3. B.2.5 Input tests
"""
import json
import re
import argparse
from pathlib import Path
import pdfplumber

def norm(s: str) -> str:
    s = s or ""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def find_page_by_content(pdf, search_text: str, max_pages: int = 84) -> int:
    """找出包含特定文字的頁面 index"""
    for i, page in enumerate(pdf.pages[:max_pages]):
        text = (page.extract_text() or '').upper()
        if search_text.upper() in text:
            return i
    return -1

def extract_overview_energy_sources(pdf) -> dict:
    """
    抽取 OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS 表格
    確保：
    1. 不漏列、不去重、保留 N/A 列
    2. 包含 ES3 Capacitor 列和 5.5.2 引用
    3. 總共 10 列（Clause 5: 3列, Clause 6: 2列, Clause 7: 1列, Clause 8: 2列, Clause 9: 1列, Clause 10: 1列）
    """
    result = {
        'page': -1,
        'rows': [],
        'has_capacitor_row': False,
        'first_es3_has_5_5_2': False,
        'has_es1_output': False,
        'raw_table': [],
        'total_rows': 0
    }

    # 找 Overview 頁
    page_idx = find_page_by_content(pdf, 'OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS')
    if page_idx < 0:
        raise ValueError("無法找到 OVERVIEW OF ENERGY SOURCES AND SAFEGUARDS 頁")

    result['page'] = page_idx + 1
    page = pdf.pages[page_idx]

    # 使用 lines-based 策略抽取表格
    tables = page.extract_tables({
        'vertical_strategy': 'lines',
        'horizontal_strategy': 'lines',
        'intersection_tolerance': 5,
        'snap_tolerance': 5,
        'join_tolerance': 5,
    })

    if not tables:
        raise ValueError("Overview 頁無法抽取表格")

    # 找最大表格
    main_table = max(tables, key=lambda t: len(t))
    result['raw_table'] = [[norm(c) if c else '' for c in row] for row in main_table]

    # 解析資料列 - 支援 ES/PS/MS/TS/RS 或純 N/A 列
    energy_pattern = re.compile(r'^(ES[123]|PS[123]|MS[123]|TS[123]|RS[123])\s*:', re.IGNORECASE)
    current_clause = ""
    current_hazard = ""
    es3_count = 0

    # Clause 對應的 hazard
    clause_hazards = {
        '5': 'Electrically-caused injury',
        '6': 'Electrically-caused fire',
        '7': 'Injury caused by hazardous substances',
        '8': 'Mechanically-caused injury',
        '9': 'Thermal burn',
        '10': 'Radiation'
    }

    for row in main_table:
        if not row or not row[0]:
            continue

        first_cell = norm(row[0])

        # 記錄 clause 和 hazard
        if re.match(r'^[5-9]$|^10$', first_cell):
            current_clause = first_cell
            current_hazard = clause_hazards.get(current_clause, '')
            continue

        # 跳過表頭
        if 'Class and Energy Source' in first_cell or 'Clause' in first_cell:
            continue
        if '(e.g.' in first_cell:
            continue
        if first_cell.upper() == 'OVERVIEW':
            continue
        if first_cell == 'Safeguards':
            continue

        # 資料列：ES/PS/MS/TS/RS 開頭 或 純 N/A 列
        is_energy_row = energy_pattern.match(first_cell)
        is_na_row = first_cell == 'N/A' and current_clause != ''

        if is_energy_row or is_na_row:
            row_data = {
                'cb_clause': int(current_clause) if current_clause else 0,
                'possible_hazard': current_hazard,
                'class_energy_source': first_cell,
                'body_or_material': norm(row[1]) if len(row) > 1 else '',
                'basic': norm(row[2]) if len(row) > 2 else '',
                'supp1': norm(row[3]) if len(row) > 3 else '',
                'supp2': norm(row[4]) if len(row) > 4 else '',
                'source_pdf_page': result['page']
            }

            # 檢查 Capacitor 列（處理換行）
            first_cell_oneline = first_cell.replace('\n', ' ')
            if 'Capacitor connected between L and N' in first_cell_oneline:
                result['has_capacitor_row'] = True
                row_data['is_capacitor_row'] = True

            # 檢查 ES1 Secondary output 列（處理換行）
            if 'ES1:' in first_cell and 'Secondary output' in first_cell_oneline:
                result['has_es1_output'] = True
                row_data['is_es1_output'] = True

            # 檢查第一個 ES3 的 5.5.2
            if first_cell.startswith('ES3:'):
                es3_count += 1
                if es3_count == 1:
                    safeguards_text = ' '.join([row_data['basic'], row_data['supp1'], row_data['supp2']])
                    if '5.5.2' in safeguards_text:
                        result['first_es3_has_5_5_2'] = True

            result['rows'].append(row_data)

    result['total_rows'] = len(result['rows'])

    # 驗證
    if not result['has_capacitor_row']:
        raise ValueError("Overview 表格缺少 'Capacitor connected between L and N' 列")

    if not result['first_es3_has_5_5_2']:
        raise ValueError("Overview 表格第一個 ES3 列的 safeguards 缺少 '5.5.2'")

    if not result['has_es1_output']:
        raise ValueError("Overview 表格缺少 'ES1: Secondary output connector' 列")

    if result['total_rows'] != 10:
        raise ValueError(f"Overview 表格應有 10 列，但只抽到 {result['total_rows']} 列")

    return result

def extract_table_5522(pdf) -> dict:
    """
    抽取 5.5.2.2 TABLE: Stored discharge on capacitors
    包含量測值和 Supplementary information
    """
    result = {
        'page': -1,
        'verdict': '',
        'rows': [],
        'supplementary_info': '',
        'x_capacitors': '',
        'bleeding_resistor': ''
    }

    # 找 5.5.2.2 頁
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ''
        if '5.5.2.2' in text and 'TABLE' in text and 'capacitor' in text.lower():
            result['page'] = i + 1

            tables = page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
            })

            for tbl in tables:
                for row in tbl:
                    if not row:
                        continue
                    first_cell = norm(row[0] or '')

                    # 找 5.5.2.2 表格開始
                    if first_cell == '5.5.2.2':
                        # 取得 verdict
                        verdict_cell = norm(row[-1] or '') if row else ''
                        if verdict_cell.upper() in ['P', 'PASS']:
                            result['verdict'] = 'P'
                        elif verdict_cell.upper() in ['N/A', 'NA']:
                            result['verdict'] = 'N/A'
                        else:
                            result['verdict'] = verdict_cell

                    # 資料列（Phase to...）
                    if first_cell.startswith('Phase to'):
                        data_row = {
                            'location': first_cell,
                            'supply_voltage': norm(row[2]) if len(row) > 2 else '',
                            'condition': norm(row[3]) if len(row) > 3 else '',
                            'switch_position': norm(row[4]) if len(row) > 4 else '',
                            'measured_voltage': norm(row[5]) if len(row) > 5 else '',
                            'es_class': norm(row[6]) if len(row) > 6 else ''
                        }
                        result['rows'].append(data_row)

                    # Supplementary information
                    if 'Supplementary information' in first_cell or 'X-capacitor' in first_cell:
                        supp_text = ' '.join([norm(c or '') for c in row])
                        result['supplementary_info'] = supp_text

                        # 抽取 X-capacitors 值
                        m = re.search(r'X-capacitors.*?:\s*([^;]+)', supp_text, re.IGNORECASE)
                        if m:
                            result['x_capacitors'] = m.group(1).strip()

                        # 抽取 bleeding resistor
                        m = re.search(r'bleeding resistor.*?:\s*([^;]+)', supp_text, re.IGNORECASE)
                        if m:
                            result['bleeding_resistor'] = m.group(1).strip()

            break

    # 驗證：若 verdict=P 必須有資料列
    if result['verdict'] == 'P' and len(result['rows']) == 0:
        raise ValueError("5.5.2.2 verdict=P 但無量測資料列")

    return result

def extract_table_b25(pdf) -> dict:
    """
    抽取 B.2.5 TABLE: Input test
    包含額定電流 I rated
    """
    result = {
        'page': -1,
        'verdict': '',
        'rows': [],
        'i_rated_values': set()  # 收集所有 I rated 值
    }

    # 找 B.2.5 TABLE 頁
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ''
        if 'B.2.5' in text and 'TABLE' in text and 'Input' in text:
            result['page'] = i + 1

            tables = page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
            })

            for tbl in tables:
                for row in tbl:
                    if not row:
                        continue
                    first_cell = norm(row[0] or '')

                    # 找 B.2.5 表頭
                    if first_cell == 'B.2.5':
                        verdict_cell = norm(row[-1] or '') if row else ''
                        if verdict_cell.upper() in ['P', 'PASS']:
                            result['verdict'] = 'P'

                    # 資料列（電壓數字開頭）
                    if re.match(r'^\d+$', first_cell):
                        voltage = first_cell
                        freq = norm(row[1]) if len(row) > 1 else ''
                        i_actual = norm(row[3]) if len(row) > 3 else ''
                        i_rated = norm(row[4]) if len(row) > 4 else ''
                        power = norm(row[5]) if len(row) > 5 else ''

                        data_row = {
                            'voltage': voltage,
                            'frequency': freq,
                            'i_actual': i_actual,
                            'i_rated': i_rated,
                            'power': power
                        }
                        result['rows'].append(data_row)

                        # 收集 I rated 值
                        if i_rated and i_rated != '--':
                            result['i_rated_values'].add(i_rated)

                    # Model 列
                    if first_cell.startswith('Model:'):
                        result['rows'].append({'model': first_cell})

            break

    # 轉換 set 為 list
    result['i_rated_values'] = list(result['i_rated_values'])

    return result

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True, help="CB PDF 路徑")
    ap.add_argument("--out_dir", required=True, help="輸出目錄")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with pdfplumber.open(args.pdf) as pdf:
        # 1. Overview 表
        print("抽取 Overview 表...")
        try:
            overview = extract_overview_energy_sources(pdf)
            print(f"  頁數: {overview['page']}")
            print(f"  資料列: {len(overview['rows'])}")
            print(f"  有 Capacitor 列: {overview['has_capacitor_row']}")
            print(f"  第一 ES3 有 5.5.2: {overview['first_es3_has_5_5_2']}")
        except ValueError as e:
            print(f"  錯誤: {e}")
            overview = {'error': str(e)}

        # 2. 5.5.2.2 表
        print("\n抽取 5.5.2.2 表...")
        try:
            table_5522 = extract_table_5522(pdf)
            print(f"  頁數: {table_5522['page']}")
            print(f"  Verdict: {table_5522['verdict']}")
            print(f"  資料列: {len(table_5522['rows'])}")
            print(f"  X-capacitors: {table_5522['x_capacitors']}")
        except ValueError as e:
            print(f"  錯誤: {e}")
            table_5522 = {'error': str(e)}

        # 3. B.2.5 表
        print("\n抽取 B.2.5 表...")
        try:
            table_b25 = extract_table_b25(pdf)
            print(f"  頁數: {table_b25['page']}")
            print(f"  Verdict: {table_b25['verdict']}")
            print(f"  資料列: {len(table_b25['rows'])}")
            print(f"  I rated 值: {table_b25['i_rated_values']}")
        except ValueError as e:
            print(f"  錯誤: {e}")
            table_b25 = {'error': str(e)}

    # 輸出
    output = {
        'overview': overview,
        'table_5522': table_5522,
        'table_b25': table_b25
    }

    out_file = out_dir / 'cb_special_tables.json'
    with open(out_file, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2, default=list)

    print(f"\n輸出: {out_file}")

if __name__ == "__main__":
    main()
