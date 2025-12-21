import json
import argparse
import re
import sys
from pathlib import Path
from copy import deepcopy
from docxtpl import DocxTemplate
from docx import Document

# 添加 core 模組路徑
sys.path.insert(0, str(Path(__file__).parent.parent))

# 導入 LLM 翻譯器
try:
    from core.llm_translator import llm_translate, get_translator
    HAS_LLM = True
except ImportError:
    HAS_LLM = False
    def llm_translate(text: str) -> str:
        return text
    def get_translator():
        return None

# 載入條款翻譯字典（若存在）
CLAUSE_TRANSLATIONS = {}
TRANSLATIONS_PATH = Path(__file__).parent / 'clause_translations.json'
if TRANSLATIONS_PATH.exists():
    with open(TRANSLATIONS_PATH, 'r', encoding='utf-8') as f:
        CLAUSE_TRANSLATIONS = json.load(f)


def clean_field_text(text: str) -> str:
    """清理欄位文字，移除填充符號和末尾冒號，用於字典匹配"""
    if not text:
        return ''
    # 統一撇號格式（彎撇號 → 直撇號）
    # \u2018 = '（左單引號），\u2019 = '（右單引號）
    cleaned = text.replace('\u2018', "'").replace('\u2019', "'")
    # 處理 Wingdings/Symbol 字體的特殊字元
    # \uf0b0 是度數符號 °
    # \uf044 是 Delta 符號 Δ（在 ΔU 中，移除 Δ 以匹配字典鍵）
    # \uf0be 是大於等於符號 ≥
    cleaned = re.sub(r'\(\s*\uf0b0\s*C\s*\)', '(C)', cleaned)
    cleaned = cleaned.replace('\uf0b0', '°')  # 其他位置的度數符號
    cleaned = cleaned.replace('\uf044', '')   # 移除 Delta 符號
    cleaned = cleaned.replace('\uf0be', '≥')  # 大於等於符號
    # 移除填充用的點號序列及其後的值（如 " ..... : R" 或 " ..... : ini, b"）
    cleaned = re.sub(r'\.?\s*\.{3,}\s*:\s*.*$', '', cleaned)
    # 移除末尾的冒號和空格
    cleaned = re.sub(r'\s*:\s*$', '', cleaned)
    # 移除末尾的單個句點（如果有）
    cleaned = re.sub(r'\.\s*$', '', cleaned)
    # 移除多餘空格
    cleaned = ' '.join(cleaned.split())
    return cleaned


def translate_field_title(text: str) -> str:
    """
    智能翻譯欄位標題：
    1. 先嘗試完全匹配原始文字
    2. 再嘗試清理後的文字（移除填充符號）
    3. 查詢翻譯字典
    4. 移除填充符號，只保留末尾冒號（符合人工輸出格式）
    """
    if not text or not re.search(r'[a-zA-Z]{2,}', text):
        return text

    # 檢查是否有冒號（末尾冒號或填充點號後的冒號）
    has_colon = bool(re.search(r':\s*$', text) or re.search(r'\.{2,}\s*:', text))

    # 1. 先嘗試完全匹配原始文字
    lookup_key = text.strip()
    if lookup_key in CLAUSE_TRANSLATIONS:
        entry = CLAUSE_TRANSLATIONS[lookup_key]
        if isinstance(entry, dict):
            title_cn = entry.get('title_cn', '')
        else:
            title_cn = str(entry) if entry else ''
        if title_cn:
            return title_cn

    # 2. 清理後的文字用於字典查詢
    cleaned = clean_field_text(text)

    # 查詢翻譯字典
    if cleaned in CLAUSE_TRANSLATIONS:
        entry = CLAUSE_TRANSLATIONS[cleaned]
        # 處理不同格式：可能是字典或字串
        if isinstance(entry, dict):
            title_cn = entry.get('title_cn', '')
        else:
            title_cn = str(entry) if entry else ''

        if title_cn:
            # 移除填充符號，只保留末尾冒號（符合人工輸出格式）
            if has_colon:
                return title_cn + ':'
            return title_cn

    return text


def normalize_text_format(text: str) -> str:
    """
    正規化文字格式 - 暫時禁用，因為人工輸出格式不一致
    直接返回原文字，避免引入新的差異
    """
    return text


def load_json(p: Path) -> dict:
    with p.open("r", encoding="utf-8") as f:
        return json.load(f)

def normalize_context(data: dict) -> dict:
    """
    讓模板端好用：提供一些常用衍生欄位（不改原資料語意）
    """
    meta = data.get("meta", {})
    meta.setdefault("model_type_references", [])

    # 拆分主型號和系列型號
    # 第一個型號為主型號，其餘為系列型號
    all_models = meta.get("model_type_references") or []
    if all_models:
        meta["main_model"] = all_models[0]  # 主型號
        meta["series_models"] = all_models[1:] if len(all_models) > 1 else []  # 系列型號
    else:
        meta["main_model"] = ""
        meta["series_models"] = []

    # 主型號欄位使用第一個型號
    meta["model_type_references_str"] = meta["main_model"]
    # 系列型號欄位使用其餘型號（逗號分隔）
    meta["series_models_str"] = ", ".join(meta["series_models"])

    # overview / clauses / attachments 若不存在，保證為 list
    data.setdefault("overview_energy_sources_and_safeguards", [])
    data.setdefault("clauses", [])
    data.setdefault("attachments_or_annex", [])

    # 方便模板顯示 QA
    qa = data.get("qa", {})
    data["qa_status"] = (qa.get("summary", {}) or {}).get("status", "MISSING")

    return data

def translate_verdict(verdict: str) -> str:
    """將英文 verdict 轉換為中文"""
    verdict_map = {
        'P': '符合',
        'PASS': '符合',
        'N/A': '不適用',
        'NA': '不適用',
        'N.A.': '不適用',
        'F': '不符合',
        'FAIL': '不符合',
    }
    return verdict_map.get(verdict.upper().strip(), verdict)


def translate_energy_source(energy_source: str, clause: int) -> str:
    """將英文 energy source 轉換為中文"""
    energy_source_oneline = energy_source.replace('\n', ' ').strip()

    translations = {
        # Clause 5 - Electrically-caused injury
        'ES3: Primary circuits supplied by a.c. mains supply': 'ES3: 所有連接到AC主電源的線路',
        'ES3: The circuit connected to AC mains (Except output circuits)': 'ES3: 所有連接到AC主電源的線路(輸出電路除外)',
        'ES3: Capacitor connected between L and N': 'ES3: X電容(於L與N之間)',
        'ES1: Secondary output connector': 'ES1: 輸出電路(輸出連接器)',
        'ES1: Output circuits': 'ES1: 輸出電路',
        # Clause 6 - Electrically-caused fire
        'PS3: All primary circuits inside the equipment enclosure': 'PS3: 設備外殼內所有的主線路',
        'PS3: All circuits except for output circuits': 'PS3: 所有電路(輸出電路除外)',
        'PS2: Secondary output connector': 'PS2: 輸出電路(輸出連接器)',
        'PS2: secondary part circuits': 'PS2: 二次側電路',
        # Clause 8 - Mechanically-caused injury
        'MS1: Mass of the unit': 'MS1: 設備質量',
        'MS1: Edges and corners': 'MS1: 邊與角',
        'MS1: Edges and corners of enclosure': 'MS1: 外殼的邊與角',
        # Clause 9 - Thermal burn
        'TS1: Plastic enclosure': 'TS1: 塑膠外殼',
        'TS1: External surface': 'TS1: 外部表面',
        'TS3: Internal parts/circuits': 'TS3: 內部零件/電路',
        # N/A
        'N/A': '無',
    }

    return translations.get(energy_source_oneline, energy_source_oneline)

def translate_body_part(body_part: str, clause: int) -> str:
    """將英文 body part / material 轉換為中文"""
    body_part_oneline = body_part.replace('\n', ' ').strip()

    translations = {
        'Ordinary': '普通人員',
        'Instructed': '受指導人員',
        'Skilled': '技術人員',
        'Ordinary, Instructed, Skilled': '普通人員、受指導人員、技術人員',
        'N/A': '無',
        'All combustible materials within equipment fire enclosure': '設備外殼內所有易燃材料',
        'Connections of secondary equipment': '二次設備的接線處',
        # Clause 6 materials
        'PCB': '印刷電路板(PCB)',
        'Enclosure': '外殼',
        'Plastic materials not part of PS3 circuit': '非 PS3 電路的塑膠材料',
    }

    return translations.get(body_part_oneline, body_part_oneline)

def translate_safeguard(safeguard: str, clause: int) -> str:
    """將英文 safeguard 轉換為中文"""
    if not safeguard or safeguard.strip() == 'N/A':
        return '無'

    safeguard_oneline = safeguard.replace('\n', ' ').strip()

    # 精確匹配翻譯
    translations = {
        'N/A': '無',
        'Enclosure': '外殼',
        'See 6.3': '見條款6.3',
        'V-0': 'V-0',
        'V-1 or better': 'V-1或更佳',
        'V-2 or better': 'V-2或更佳',
    }

    if safeguard_oneline in translations:
        return translations[safeguard_oneline]

    # 模式匹配處理
    if 'Enclosure See' in safeguard_oneline or 'Enclosure, see' in safeguard_oneline:
        # 提取條款號碼
        import re
        clauses = re.findall(r'[0-9]+\.[0-9]+(?:\.[0-9]+)?', safeguard_oneline)
        if clauses:
            return f"外殼, 見條款{', '.join(clauses)}"
        return '外殼'
    if 'See 5.5.2.2' in safeguard_oneline:
        return '見條款5.5.2.2'
    if 'Equipment safeguard' in safeguard_oneline and 'no ignition' in safeguard_oneline:
        return '設備防護(無點燃發生)'
    if 'Equipment safeguard' in safeguard_oneline and 'control of fire' in safeguard_oneline:
        return '設備防護(控制火焰擴散)'

    return safeguard_oneline

def copy_row_style(source_row, target_row):
    """複製列的格式（不複製內容）"""
    # 複製儲存格數量和基本屬性
    for i, (src_cell, tgt_cell) in enumerate(zip(source_row.cells, target_row.cells)):
        # 複製段落格式
        if src_cell.paragraphs and tgt_cell.paragraphs:
            tgt_cell.paragraphs[0].paragraph_format.alignment = src_cell.paragraphs[0].paragraph_format.alignment

def add_row_after(table, reference_row_idx):
    """在指定列後面新增一列"""
    from copy import deepcopy
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # 取得參考列
    ref_row = table.rows[reference_row_idx]
    ref_tr = ref_row._tr

    # 複製列
    new_tr = deepcopy(ref_tr)

    # 清空新列的文字
    for tc in new_tr.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
        for p in tc.findall('.//w:p', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            for r in list(p):
                if r.tag.endswith('}r'):
                    for t in list(r):
                        if t.tag.endswith('}t'):
                            t.text = ''

    # 在參考列後插入
    ref_tr.addnext(new_tr)

    return len(table.rows) - 1  # 回傳新列的索引（可能不準確，需要重新掃描）

def fill_overview_table_from_cb_p12(doc: Document, overview_cb_p12_rows: list):
    """
    使用 overview_cb_p12_rows 更新安全防護總攬表的 safeguard 欄位
    新策略：保留模板結構，只更新 safeguard（基本、補充、強化）欄位
    不刪除模板中的任何列
    """
    if not overview_cb_p12_rows:
        print("警告：overview_cb_p12_rows 為空，保留模板預設值")
        return 0

    # 找安全防護總攬表
    overview_table = None
    for tbl in doc.tables:
        if tbl.rows and '安全防護總攬表' in tbl.rows[0].cells[0].text:
            overview_table = tbl
            break

    if not overview_table:
        print("警告：找不到安全防護總攬表")
        return 0

    # 按 clause 分組 PDF 資料
    pdf_by_clause = {}
    for row_data in overview_cb_p12_rows:
        clause = row_data.get('cb_clause', 0)
        if clause not in pdf_by_clause:
            pdf_by_clause[clause] = []
        pdf_by_clause[clause].append(row_data)

    # 掃描模板中各章節的資料列
    def scan_clause_sections():
        sections = {}
        current = None
        for idx, row in enumerate(overview_table.rows):
            first_cell = row.cells[0].text.strip()
            if first_cell == '5.1':
                current = 5
                sections[5] = {'start': idx, 'data_rows': []}
            elif first_cell == '6.1':
                current = 6
                sections[6] = {'start': idx, 'data_rows': []}
            elif first_cell == '7.1':
                current = 7
                sections[7] = {'start': idx, 'data_rows': []}
            elif first_cell == '8.1':
                current = 8
                sections[8] = {'start': idx, 'data_rows': []}
            elif first_cell == '9.1':
                current = 9
                sections[9] = {'start': idx, 'data_rows': []}
            elif first_cell == '10.1':
                current = 10
                sections[10] = {'start': idx, 'data_rows': []}

            # 偵測資料列
            if current and (
                first_cell.startswith('ES') or
                first_cell.startswith('PS') or
                first_cell.startswith('MS') or
                first_cell.startswith('TS') or
                first_cell.startswith('RS') or
                first_cell in ['N/A', '無']
            ):
                sections[current]['data_rows'].append(idx)
        return sections

    clause_sections = scan_clause_sections()
    updated_rows = 0

    # 建立 PDF 資料的 energy source 類型映射（用於匹配）
    def get_energy_type(text):
        """從能源來源文字中提取類型 (ES1, ES3, PS2, PS3, MS1, TS1, TS3, RS1, N/A)"""
        text = text.upper()
        for prefix in ['ES3', 'ES2', 'ES1', 'PS3', 'PS2', 'PS1', 'MS3', 'MS2', 'MS1', 'TS3', 'TS2', 'TS1', 'RS3', 'RS2', 'RS1']:
            if prefix in text:
                return prefix
        if 'N/A' in text or text == '無':
            return 'N/A'
        return None

    # 對每個章節，根據能源類型匹配更新 safeguard 欄位
    for clause in [5, 6, 7, 8, 9, 10]:
        if clause not in clause_sections:
            continue

        section = clause_sections[clause]
        template_data_rows = section['data_rows']
        pdf_rows = pdf_by_clause.get(clause, [])

        if not pdf_rows:
            continue

        # 建立 PDF 資料的能源類型映射
        pdf_by_type = {}
        for pdf_row in pdf_rows:
            energy_source = pdf_row.get('energy_source', '') or pdf_row.get('class_energy_source', '')
            etype = get_energy_type(energy_source)
            if etype:
                if etype not in pdf_by_type:
                    pdf_by_type[etype] = []
                pdf_by_type[etype].append(pdf_row)

        # 遍歷模板中的每個資料列，根據能源類型匹配更新
        for row_idx in template_data_rows:
            table_row = overview_table.rows[row_idx]
            template_energy = table_row.cells[0].text.strip()
            template_type = get_energy_type(template_energy)

            if not template_type or template_type not in pdf_by_type:
                continue

            # 取出對應類型的第一個 PDF 資料（後續的用於相同類型的不同子項目）
            if pdf_by_type[template_type]:
                pdf_row = pdf_by_type[template_type][0]

                # 更新 safeguard 欄位（保留模板的能源來源和身體部位）
                basic = pdf_row.get('safeguard_basic', '') or pdf_row.get('basic', '')
                supp1 = pdf_row.get('safeguard_supplementary', '') or pdf_row.get('supp1', '')
                supp2 = pdf_row.get('safeguard_reinforced', '') or pdf_row.get('supp2', '')

                # 只更新 safeguard 欄位（第2-4欄），保留第0-1欄的模板內容
                if basic:
                    table_row.cells[2].text = translate_safeguard(basic, clause)
                if supp1:
                    table_row.cells[3].text = translate_safeguard(supp1, clause)
                if supp2:
                    table_row.cells[4].text = translate_safeguard(supp2, clause)

                updated_rows += 1

                # 對於 PS3/PS2 等有多個子項目的類型，不移除 PDF 資料
                # 讓相同類型的模板列共享相同的 safeguard 資訊

    print(f"已更新 {updated_rows} 列的 safeguard 欄位")
    return updated_rows

def rebuild_clause_tables_v2(doc: Document, pdf_clause_rows: list) -> dict:
    """
    完全動態重建主條款表格
    - 清空模板所有條款資料列
    - 依照 PDF 資料順序新增列
    - remark 完全來自 PDF（空就是空）
    - 處理合併的 B-M 表格（模板中 Table 13 包含 B~M 所有章節）

    Args:
        doc: Word 文件
        pdf_clause_rows: PDF 抽取的條款列表 [{'clause_id', 'req', 'remark', 'verdict', 'pdf_page'}, ...]

    Returns:
        dict: QA 比對結果
    """
    if not pdf_clause_rows:
        print("警告：pdf_clause_rows 為空，無法重建條款表格")
        return {'pdf_row_count': 0, 'word_row_count': 0, 'match': False}

    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # 數字章節（獨立表格）
    numeric_sections = ['4', '5', '6', '7', '8', '9', '10']
    # 字母章節（合併在 B 表格中）- 包含 B~M 和 N~Y 所有附錄
    letter_sections = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M',
                       'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']

    # 依章節分組 PDF 資料，保持順序
    # 關鍵：只處理主條款表格區域（page 12-43 左右），排除後面的附表（如 page 44+ 的 TABLE 數據）
    # 策略：追蹤章節順序，一旦進入下一章節就不再回頭；進入附表區域後停止處理
    pdf_by_section = {}
    current_section = None
    section_order = ['4', '5', '6', '7', '8', '9', '10',
                     'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M',
                     'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y']
    seen_sections = set()  # 已經完全處理過的章節
    in_appendix_table_area = False  # 是否已進入附表區域

    for row in pdf_clause_rows:
        cid = row.get('clause_id', '')
        req = row.get('req', '')

        # 檢查是否進入附表區域（TABLE: 開頭的條款）
        if cid and 'TABLE:' in req:
            in_appendix_table_area = True
            continue  # 跳過附表

        # 如果已在附表區域，跳過所有後續內容
        if in_appendix_table_area:
            continue

        if cid:
            sec = cid.split('.')[0]

            # 檢查是否是回頭的章節（附表）
            if sec in section_order:
                sec_idx = section_order.index(sec)
                # 如果當前章節在已處理章節之前，且已經處理過其他章節，這是附表，跳過
                if current_section and current_section != sec:
                    current_idx = section_order.index(current_section) if current_section in section_order else -1
                    if sec_idx < current_idx:
                        # 這是回頭的條款（如 page 44 的 5.2 TABLE），標記進入附表區域
                        in_appendix_table_area = True
                        continue

            current_section = sec
            if sec not in seen_sections:
                seen_sections.add(sec)
        else:
            # 無 clause_id 的列，歸入前一個章節
            sec = current_section

        if sec:
            if sec not in pdf_by_section:
                pdf_by_section[sec] = []
            pdf_by_section[sec].append(row)

    total_updated = 0
    word_rows_generated = []

    def fill_cell_text(tc, text):
        """填入儲存格文字"""
        p_elements = tc.findall('.//w:p', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if p_elements:
            p = p_elements[0]
            for child in list(p):
                if child.tag.endswith('}r'):
                    p.remove(child)
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = text or ''
            r.append(t)
            p.append(r)

    def set_cell_shading(tc, color):
        """設定儲存格背景色"""
        tcPr = tc.find('.//w:tcPr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        # 移除現有的 shd
        for shd in tcPr.findall('.//w:shd', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            tcPr.remove(shd)
        # 添加新的 shd
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)
        tcPr.append(shd)

    def insert_pdf_row(template_tr, pdf_row):
        """插入一列 PDF 資料"""
        new_tr = deepcopy(template_tr)
        cells = new_tr.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})

        clause_id = pdf_row.get('clause_id', '')
        req = pdf_row.get('req', '')
        remark = pdf_row.get('remark', '')
        verdict = pdf_row.get('verdict', '')

        # 判斷是否為章節標題行（需要灰色背景）
        # 章節標題：clause_id 是單一數字或字母（4, 5, 6, ..., B, C, ...）
        is_section_header = clause_id in section_order

        # 優先使用翻譯字典，否則嘗試翻譯常見英文片語
        if clause_id and clause_id in CLAUSE_TRANSLATIONS:
            title_cn = CLAUSE_TRANSLATIONS[clause_id].get('title_cn', req)
        else:
            title_cn = translate_req(req)
        verdict_cn = translate_verdict(verdict)
        remark_cn = translate_remark(remark, clause_id) if remark else ''

        cell_data = [clause_id, title_cn, remark_cn, verdict_cn]
        for i, tc in enumerate(cells[:4]):
            fill_cell_text(tc, cell_data[i] if i < len(cell_data) else '')
            # 為章節標題行添加灰色背景
            if is_section_header:
                set_cell_shading(tc, 'D9D9D9')  # 淺灰色

        # 當 verdict 為 ⎯ 時，為 verdict 欄位設定灰色背景並清空文字
        if verdict in ['⎯', '-', '—', '–']:
            verdict_tc = cells[3] if len(cells) > 3 else None
            if verdict_tc is not None:
                set_cell_shading(verdict_tc, 'D9D9D9')  # 淺灰色
                fill_cell_text(verdict_tc, '')  # 清空文字（符合人工版本）

        word_rows_generated.append({
            'clause_id': clause_id,
            'remark': remark_cn,
            'verdict': verdict_cn
        })

        return new_tr

    for tbl_idx, tbl in enumerate(doc.tables):
        if not tbl.rows or len(tbl.rows) < 2:
            continue

        first_cell = tbl.rows[0].cells[0].text.strip()

        # 處理數字章節表格
        if first_cell in numeric_sections:
            section = first_cell
            if section not in pdf_by_section:
                continue

            pdf_rows = pdf_by_section[section]

            if len(tbl.rows) < 2:
                print(f"  警告：表格 {tbl_idx} (Section {section}) 沒有足夠的列作為模板")
                continue

            # 尋找有 4 個 w:tc 元素的列作為模板
            template_row_idx = None
            for row_idx, row in enumerate(tbl.rows):
                tr = row._tr
                cells = tr.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if len(cells) == 4:
                    template_row_idx = row_idx
                    break

            if template_row_idx is None:
                print(f"  警告：表格 {tbl_idx} (Section {section}) 找不到 4 欄的模板列")
                continue

            template_tr = tbl.rows[template_row_idx]._tr

            # 清空表格（保留模板列）
            # 刪除模板列之前的所有列
            for _ in range(template_row_idx):
                first_tr = tbl.rows[0]._tr
                first_tr.getparent().remove(first_tr)

            # 刪除模板列之後的所有列
            while len(tbl.rows) > 1:
                tr = tbl.rows[-1]._tr
                tr.getparent().remove(tr)

            # 插入所有 PDF 列（反向插入以保持順序）
            for pdf_row in reversed(pdf_rows):
                new_tr = insert_pdf_row(template_tr, pdf_row)
                template_tr.addnext(new_tr)

            # 刪除模板列
            tbl.rows[0]._tr.getparent().remove(tbl.rows[0]._tr)

            total_updated += len(pdf_rows)
            print(f"  表格 {tbl_idx} (Section {section}): 新增 {len(pdf_rows)} 列")

        # 處理字母章節表格（B 表格包含 B~M 所有章節）
        elif first_cell == 'B':
            # 收集所有字母章節的 PDF 資料
            combined_pdf_rows = []
            for sec in letter_sections:
                if sec in pdf_by_section:
                    combined_pdf_rows.extend(pdf_by_section[sec])

            if not combined_pdf_rows:
                continue

            if len(tbl.rows) < 2:
                print(f"  警告：表格 {tbl_idx} (Section B) 沒有足夠的列作為模板")
                continue

            # 尋找有 4 個 w:tc 元素的列作為模板
            template_row_idx = None
            for row_idx, row in enumerate(tbl.rows):
                tr = row._tr
                cells = tr.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if len(cells) == 4:
                    template_row_idx = row_idx
                    break

            if template_row_idx is None:
                print(f"  警告：表格 {tbl_idx} (Section B) 找不到 4 欄的模板列")
                continue

            template_tr = tbl.rows[template_row_idx]._tr

            # 清空表格（保留模板列）
            # 刪除模板列之前的所有列
            for _ in range(template_row_idx):
                first_tr = tbl.rows[0]._tr
                first_tr.getparent().remove(first_tr)

            # 刪除模板列之後的所有列
            while len(tbl.rows) > 1:
                tr = tbl.rows[-1]._tr
                tr.getparent().remove(tr)

            # 插入所有字母章節的 PDF 列（反向插入以保持順序）
            for pdf_row in reversed(combined_pdf_rows):
                new_tr = insert_pdf_row(template_tr, pdf_row)
                template_tr.addnext(new_tr)

            # 刪除模板列
            tbl.rows[0]._tr.getparent().remove(tbl.rows[0]._tr)

            sections_included = [s for s in letter_sections if s in pdf_by_section]
            total_updated += len(combined_pdf_rows)
            print(f"  表格 {tbl_idx} (Sections {','.join(sections_included)}): 新增 {len(combined_pdf_rows)} 列")

    print(f"條款表格完全重建：共 {total_updated} 列")

    return {
        'pdf_row_count': len(pdf_clause_rows),
        'word_row_count': total_updated,
        'word_rows': word_rows_generated,
        'match': len(pdf_clause_rows) == total_updated
    }


def rebuild_clause_tables(doc: Document, clauses: list) -> dict:
    """
    舊版：只更新 verdict 和 remark，不清空模板
    保留向後相容
    """
    if not clauses:
        print("警告：clauses 為空，無法重建條款表格")
        return {'pdf_clause_ids': set(), 'word_clause_ids': set()}

    clause_map = {}
    for c in clauses:
        cid = c.get('clause_id', '')
        if cid:
            clause_map[cid] = c

    pdf_clause_ids = set(clause_map.keys())
    word_clause_ids = set()

    main_clause_prefixes = ['4', '5', '6', '7', '8', '9', '10', 'B']

    updated_tables = 0
    updated_rows = 0

    for tbl_idx, tbl in enumerate(doc.tables):
        if not tbl.rows:
            continue

        first_cell = tbl.rows[0].cells[0].text.strip()

        is_main_clause_table = False
        for prefix in main_clause_prefixes:
            if first_cell == prefix or (len(first_cell) <= 3 and first_cell.startswith(prefix)):
                is_main_clause_table = True
                break

        if not is_main_clause_table:
            continue

        for row_idx, row in enumerate(tbl.rows):
            if len(row.cells) < 4:
                continue

            template_clause_id = row.cells[0].text.strip()

            if not template_clause_id or template_clause_id.startswith('{'):
                continue

            if template_clause_id in clause_map:
                pdf_data = clause_map[template_clause_id]
                word_clause_ids.add(template_clause_id)

                remark = pdf_data.get('test_result_or_remark', '').replace('\n', ' ')
                verdict = pdf_data.get('verdict', '')
                verdict_cn = translate_verdict(verdict)

                # 關鍵改動：PDF remark 為空時，清空 Word 該欄（不保留模板）
                if remark:
                    remark_cn = translate_remark(remark, template_clause_id)
                else:
                    remark_cn = ''  # PDF 空就是空
                row.cells[2].text = remark_cn

                row.cells[3].text = verdict_cn

                updated_rows += 1

        updated_tables += 1

    print(f"條款表格更新：{updated_tables} 個表格，{updated_rows} 列")
    print(f"PDF clause_ids: {len(pdf_clause_ids)}, Word clause_ids: {len(word_clause_ids)}")

    return {
        'pdf_clause_ids': pdf_clause_ids,
        'word_clause_ids': word_clause_ids,
        'match': pdf_clause_ids == word_clause_ids
    }


def translate_req(req: str) -> str:
    """翻譯常見的 requirement 英文片語"""
    # 正規化：移除換行並壓縮多餘空白
    req_normalized = ' '.join(req.split())

    # 精確匹配翻譯（完整的 CB 報告術語）
    exact_translations = {
        # 通用術語
        'General': '一般',
        'General requirements': '一般要求',
        'General requirement': '一般要求',
        'Compliance': '符合性',
        'Requirements': '要求',
        'Test': '試驗',
        'Tests': '試驗',
        'Result': '結果',
        'Results': '結果',
        'N/A': '不適用',
        'NA': '不適用',
        'Pass': '符合',
        'Fail': '不符合',
        'Normal': '正常',
        'Interchangeable': '可互換',
        'Interchangeabl e': '可互換',
        'Test method and compliance': '試驗方法及符合性',

        # 4.x 一般要求
        'Protective bonding conductor size (mm2)': '保護連接導體截面積 (mm²)',
        'Protective bonding conductor size (mm2). ............... :': '保護連接導體截面積 (mm²)',
        'Terminal size for connecting protective bonding conductors (mm2)': '連接保護連接導體之端子尺寸 (mm²)',
        'Terminal size for connecting protective bonding conductors (mm) ......': '連接保護連接導體之端子尺寸 (mm)',
        'Terminal size for connecting protective bonding conductors (mm) ................': '連接保護連接導體之端子尺寸 (mm)',
        'Terminal size for connecting protective bonding conductors (mm) ....................................': '連接保護連接導體之端子尺寸 (mm)',
        'Terminal size for connecting protective earthing conductors (mm) .....': '連接保護接地導體之端子尺寸 (mm)',
        'Terminal size for connecting protective earthing conductors (mm) ...............': '連接保護接地導體之端子尺寸 (mm)',
        'Terminal size for connecting protective earthing conductors (mm) ...................................': '連接保護接地導體之端子尺寸 (mm)',
        'Protective current rating (A) .................................... :': '保護電流額定值 (A)',
        'Requirements for protective bonding conductors': '保護連接導體之要求',
        'Protective bonding conductors': '保護連接導體',
        'Terminals for protective conductors': '保護導體之端子',
        'Corrosion': '腐蝕',
        'Resistance to corrosion': '耐腐蝕性',

        # 試驗條件與參數
        'Normal, abnormal and fault condition': '正常、異常及故障條件',
        'Normal, abnormal and fault conditions': '正常、異常及故障條件',
        'Operating and fault condition': '操作及故障條件',
        'Operating and fault conditions': '操作及故障條件',
        'Worst-case fault': '最差情況故障',
        'Overload': '過載',
        'Material': '材料',
        'Tapes': '膠帶',
        'English': '英文',

        # 尺寸與參數項目
        'Wall thickness (mm) .............................................. :': '壁厚 (mm)',
        'Conditioning (C) ................................................... :': '調節 (°C)',
        'Conditioning, T (°C) ............................................. : C': '調節溫度 T (°C)',
        'Force applied (N) .................................................... :': '施加力 (N)',
        'Wheels diameter (mm) ............................................ :': '輪子直徑 (mm)',
        'Number of handles .................................................. :': '把手數量',
        'Loading force applied (N) ........................................ :': '施加負載力 (N)',
        'Lasers ..................................................................... :': '雷射',
        'Lamps and lamp systems ...................................... :': '燈具及燈具系統',
        'Personal music player ............................................ :': '個人音樂播放器',
        'Risk group marking and location ............................. :': '風險群組標示及位置',
        'UV radiation exposure............................................. :': '紫外線輻射曝露',
        'Warning for MEL ≥ 100 dB(A) ................................. :': 'MEL ≥ 100 dB(A) 警告',
        'Language .............................................................. :': '語言',
        'Test temperature (C) ............................................ :': '試驗溫度 (°C)',
        'Position .................................................................. :': '位置',
        'FIW wire nominal diameter .................................... :': 'FIW 導線標稱直徑',
        'Operating voltage ................................................. :': '工作電壓',
        'Type ....................................................................... :': '類型',
        "Manufacturers' defined drift .................................. :": '製造商定義之偏移量',
        'Type test voltage V ............................................ : ini,a': '型式試驗電壓 Vini,a',
        'Routine test voltage, V ...................................... : ini, b': '例行試驗電壓 Vini,b',
        'Solid round winding wire, diameter (mm) .............. :': '實心圓繞線直徑 (mm)',
        'Check of the charge/discharge function': '充放電功能檢查',
        'Minimum air flow rate, Q (m3/h) ............................. :': '最小空氣流量 Q (m³/h)',
        'Material(s) used ..................................................... :': '使用之材料',
        'Value of X (mm) ..................................................... :': 'X 值 (mm)',
        'X=1. Pollution degree': 'X=1. 污染等級',
        'Location and Dimensions (mm) ............................ :': '位置及尺寸 (mm)',
        'Duration (weeks) .................................................... :': '持續時間 (週)',

        # 音頻相關
        'Max. acoustic output L , dB(A) ........................... : Aeq,T': '最大聲輸出 LAeq,T, dB(A)',
        'Rated load impedance (Ω) .................................... :': '額定負載阻抗 (Ω)',

        # 熱切斷器相關
        'Thermal cut-outs separately approved according to IEC 60730 with conditions indi': '依 IEC 60730 單獨認證之熱切斷器及條件',
        'Thermal cut-outs tested as part of the equipment as indicated in c)': '依 c) 所述作為設備一部分測試之熱切斷器',
        'Over current protection by circuit design.': '電路設計之過電流保護。',

        # 光耦合器與繞線
        'Approved opto-coupler used. (See appended table 4.1.2)': '使用認可之光耦合器。(見附表 4.1.2)',
        'Smallest capacitance and smallest resistance specified by ICX manufacturer for i': 'ICX 製造商指定之最小電容與最小電阻',
        'Certified triple insulation wire used as secondary winding. (See appended table': '經認證之三重絕緣導線用於二次繞組。(見附表',

        # ES3/PS3 區域
        'The ES3 and PS3 keep-out volume in Figure P.3': '圖 P.3 中之 ES3 及 PS3 禁區體積',

        # 更多測試項目
        'e) IC current limiter complying with G.9': 'e) 符合 G.9 之 IC 電流限制器',
        'Test for external circuits – paired conductor': '外部電路試驗 – 成對導體',
        'Maximum output current (A) ................................. :': '最大輸出電流 (A)',
        'Current limiting method.......................................... :': '電流限制方法',
        'Test setup': '試驗設置',
        'Cord/cable used for test ........................................ :': '試驗用電線/電纜',
        'Test flame according to IEC 60695-11-5 with': '依 IEC 60695-11-5 之試驗火焰',
        'Flammability test for the bottom of a fire enclosure': '防火外殼底部可燃性試驗',
        'Mounting of samples': '樣品安裝',
        'Mounting of samples ............................................. :': '樣品安裝',
        'Steady force test, 10 N ....................................... :': '穩定力試驗，10 N',
        'Steady force test, 30 N ....................................... :': '穩定力試驗，30 N',
        'Conditioning (C) ................................................... :': '調節 (°C)',
        'Operating and fault condition': '操作及故障條件',

        # 更多試驗項目
        'Steady force test, 100 N ..................................... :': '穩定力試驗，100 N',
        'Steady force test, 250 N ..................................... :': '穩定力試驗，250 N',
        'Enclosure impact test': '外殼撞擊試驗',
        'Fall test': '落下試驗',
        'Swing test': '擺動試驗',
        'Tapes': '膠帶',
        'Overall diameter or minor overall dimension, D (mm) ............................': '整體直徑或較小整體尺寸 D (mm)',
        'Largest capacitance and smallest resistance for ICX tested by itself for 10000 c': 'ICX 單獨測試 10000 週期之最大電容與最小電阻',
        'Terminal size for connecting protective bonding conductors (mm) ................': '連接保護連接導體之端子尺寸 (mm)',
        'b) Equipment connected to unearthed external circuits, current (mA) ............': 'b) 連接至非接地外部電路之設備，電流 (mA)',

        'Class II with functional earthing marking': 'Class II 帶功能接地標誌',
        'Special conditions for temperature limited by fuse': '保險絲限溫之特殊條件',
        'Glass impact test (1J)': '玻璃撞擊試驗 (1J)',
        'Push/pull test (10 N)': '推/拉試驗 (10 N)',
        'No harm by explosion during single fault conditions': '單一故障條件下爆炸不致造成傷害',
        'Fix conductors not to defeat a safeguard': '導線固定不致使安全防護失效',
        'Open torque test': '開啟扭力試驗',
        '30N force test with test probe': '30N 測試探針力量試驗',
        '20N force test with test hook': '20N 測試鉤力量試驗',
        'Interchangeable': '可互換',
        'Stress relief test': '應力消除試驗',
        'Battery replacement test': '電池更換試驗',
        'Drop test': '落下試驗',
        'Impact test': '撞擊試驗',
        'Crush test': '擠壓試驗',
        'Steady force test': '穩定力試驗',
        'Adhesion test': '附著力試驗',
        'Legibility test': '清晰度試驗',
        'Durability test': '耐久性試驗',
        'Air comprising a safeguard': '空氣構成之安全防護',
        'Acceptance of materials, components and subassemblies': '材料、組件及次組件之允收',
        'Access protection override': '接觸防護覆蓋',
        'Accessibility, glass, safeguard effectiveness': '可接近性、玻璃、安全防護有效性',
        'Accessible part criterion': '可接觸部位判斷準則',
        'Accessible parts of equipment': '設備可接觸部位',
        'Alternative method': '替代方法',
        'Ball pressure test': '球壓試驗',
        'Blocked motor test': '馬達堵轉試驗',
        'Tilt test': '傾斜試驗',
        'Explosion test': '爆炸試驗',

        # 5.x 電擊防護
        'General Requirements for accessible parts to ordinary, instructed and skilled persons': '普通人員、受指導人員及技術人員可接觸部位之一般要求',
        'Accessible ES1/ES2 derived from ES2/ES3 circuits': '由 ES2/ES3 電路衍生之可接觸 ES1/ES2',
        'Skilled persons not unintentional contact ES3 bare conductors': '技術人員不致意外接觸 ES3 裸導體',
        'Skilled persons not unintentional contact ES3 bare conductor': '技術人員不致意外接觸 ES3 裸導體',
        'Contact requirements': '接觸要求',
        'Test with test probe from Annex V': '使用附錄 V 測試探針進行試驗',
        'Air gap – electric strength test potential (V) .............': '空氣間隙 - 耐電壓試驗電壓 (V)',
        'Air gap – distance (mm) .........................................': '空氣間隙 - 距離 (mm)',

        # 5.4.4.9 固態絕緣 - 多種格式
        'Solid insulation at frequencies >30 kHz, E , K , d, P R V (V) ......................................': '頻率 >30 kHz 之固態絕緣，E、K、d、P、R、V (V)',
        'Solid insulation at frequencies >30 kHz, E , K , d, P R V (V) .................................................................... : PW': '頻率 >30 kHz 之固態絕緣，E、K、d、P、R、VPW (V)',
        'TABLE: Solid insulation at frequencies >30 kHz': '表格：頻率 >30 kHz 之固態絕緣',
        'Alternative by electric strength test, tested voltage (V), K .......................................': '耐電壓試驗替代方法，試驗電壓 (V)，K',

        # 5.4.5 天線端子
        'Antenna terminal insulation': '天線端子絕緣',

        # 5.4.11.2 要求
        'SPDs bridge separation between external circuit and earth': 'SPD 跨越外部電路與接地間之隔離',
        'Rated operating voltage U (V) ............................. : op': '額定工作電壓 Uop (V)',
        'Rated operating voltage U (V) ............................. :': '額定工作電壓 U (V)',
        'Nominal voltage U (V) ....................................... : peak': '標稱電壓 Upeak (V)',
        'Nominal voltage U (V) ....................................... :': '標稱電壓 U (V)',
        'Nominal voltage U (V) ...................................... :': '標稱電壓 U (V)',
        'Nominal voltage U (V) ............................': '標稱電壓 U (V)',
        'Max increase due to variation U ........................ : sp': '因變動產生之最大增量 ΔUsp',
        'Max increase due to variation U ........................ :': '因變動產生之最大增量 ΔU',
        'Max increase due to variation ΔU': '因變動產生之最大增量 ΔU',
        'Max increase due to variation \uf044U ........................ : sp': '因變動產生之最大增量 ΔUsp',
        'Max increase due to variation \uf044U ........................ :': '因變動產生之最大增量 ΔU',
        'Max increase due to ageing U .......................... : sa': '因老化產生之最大增量 ΔUsa',
        'Max increase due to ageing U .......................... :': '因老化產生之最大增量 ΔU',
        'Max increase due to ageing ΔU': '因老化產生之最大增量 ΔU',
        'Max increase due to ageing \uf044U .......................... : sa': '因老化產生之最大增量 ΔUsa',
        'Max increase due to ageing \uf044U .......................... :': '因老化產生之最大增量 ΔU',

        # 5.6.8 功能接地
        'Functional earthing': '功能接地',
        'Conductor size (mm2) ............................................. :': '導體截面積 (mm²)',

        # 5.7.6 觸電電流
        'Requirements when touch current exceeds ES2 limits': '觸電電流超過 ES2 限值時之要求',
        'Protective conductor current (mA) ......................... :': '保護導體電流 (mA)',

        # 5.7.8 觸電電流總和
        'Summation of touch currents from external circuits': '外部電路觸電電流總和',
        'a) Equipment connected to earthed external circuits, current (mA) ..................................': 'a) 連接至接地外部電路之設備，電流 (mA)',
        'a) Equipment connected to earthed external circuits, current (mA) ..............': 'a) 連接至接地外部電路之設備，電流 (mA)',
        'a) Equipment connected to earthed external circuits, current (mA) ..................................': 'a) 連接至接地外部電路之設備，電流 (mA)',
        'b) Equipment connected to unearthed external circuits, current (mA) .............................': 'b) 連接至非接地外部電路之設備，電流 (mA)',
        'b) Equipment connected to unearthed external circuits, current (mA) ............': 'b) 連接至非接地外部電路之設備，電流 (mA)',
        'b) Equipment connected to unearthed external circuits, current (mA) ................................': 'b) 連接至非接地外部電路之設備，電流 (mA)',

        # 5.8 電池備援
        'Backfeed safeguard in battery backed up supplies': '電池備援供電之反饋安全防護',
        'Mains terminal ES ................................................... :': '主電源端子 ES',

        'Accessibility to outdoor equipment bare parts': '室外設備裸露部位之可接觸性',
        'Test with test probe from Annex V': '使用附錄 V 測試探針進行試驗',
        'Electric strength test': '耐電壓試驗',
        'Temporary overvoltage': '暫態過電壓',
        'Clearances in circuits connected to AC Mains, Alternative method': '連接 AC 主電源電路之間隙，替代方法',
        'SPDs bridge separation between external circuit and earth': 'SPD 跨越外部電路與接地間之隔離',
        'Protective earthing conductor serving as a reinforced safeguard': '保護接地導體作為強化安全防護',
        'Protective earthing conductor serving as a double safeguard': '保護接地導體作為雙重安全防護',
        'Accessibility to electrical energy sources and safeguards': '電能源及安全防護之可接觸性',

        # 6.x 火災防護
        'Combustible materials outside fire enclosure': '防火外殼外之可燃材料',
        'Flammability tests for the bottom of a fire enclosure': '防火外殼底部可燃性試驗',
        'Flammability test for fire enclosures and fire barrier materials of equipment': '設備防火外殼及防火屏障材料之可燃性試驗',
        'Flammability test for fire enclosure and fire barrier integrity': '防火外殼及防火屏障完整性可燃性試驗',
        'Flammability test for fire enclosure materials of': '防火外殼材料之可燃性試驗',
        'Flammability classification of materials': '材料可燃性分類',
        'Test method and compliance': '試驗方法及符合性',
        'Bottom openings and properties': '底部開口及特性',
        '- Material extinguishes within 30s': '- 材料在 30 秒內熄滅',
        '- Material not consumed completely': '- 材料未完全燒盡',
        '- Mechanical function check and visual inspection': '- 機械功能檢查及目視檢驗',
        '- No burning of layer or wrapping tissue': '- 包覆層或包裝紙未燃燒',
        'TESTS FOR RESISTANCE TO HEAT AND FIRE': '耐熱與耐火試驗',

        # 電池相關
        'Overcharging of a rechargeable battery': '可充電電池過充電',
        'Excessive discharging': '過度放電',
        'Unintentional charging of a non-rechargeable': '非充電電池意外充電',
        'Reverse charging of a rechargeable battery': '可充電電池反向充電',
        'Calculated hydrogen generation rate': '計算之氫氣產生率',
        'Hydrogen gas concentration (%)': '氫氣濃度 (%)',
        'Obtained hydrogen generation rate': '實測氫氣產生率',

        # 外物進入防護
        'Safeguards against entry or consequences of entry of a foreign object': '防止外物進入或進入後果之安全防護',
        'Safeguards against entry of a foreign object': '防止外物進入之安全防護',
        'Safeguards against the consequences of entry of a foreign object': '防止外物進入後果之安全防護',
        'Safeguard requirements': '安全防護要求',
        'Transportable equipment with metalized plastic': '具金屬化塑膠之可攜式設備',
        'Consequence of entry test': '進入後果試驗',
        'Safeguards against spillage of internal liquids': '防止內部液體溢出之安全防護',
        'Determination of spillage consequences': '溢出後果判定',
        'Spillage safeguards': '溢出安全防護',
        'Metallized coatings and adhesives securing parts': '金屬化塗層及固定零件之黏著劑',
        'Glass fragmentation test': '玻璃碎裂試驗',
        'Test for telescoping or rod antennas': '伸縮或桿狀天線試驗',

        # 輸出限制
        'a) Inherently limited output': 'a) 固有限制輸出',
        'c) Regulating network limited output': 'c) 調節網路限制輸出',
        'd) Overcurrent protective device limited output': 'd) 過電流保護裝置限制輸出',
        'Current rating of overcurrent protective device (A)': '過電流保護裝置額定電流 (A)',
        'Overcurrent protective device for test': '試驗用過電流保護裝置',

        # 電路互連
        'CIRCUITS INTENDED FOR INTERCONNECTION WITH BUILDING WIRING': '預定與建築配線互連之電路',
        'In circuit isolated from mains, separation distance': '與主電源隔離之電路中，隔離距離',

        # 其他大寫標題
        'ELECTROCHEMICAL POTENTIALS': '電化學電位',
        'MEASUREMENT OF CREEPAGE DISTANCES AND CLEARANCES': '沿面距離與間隙之量測',
        'SAFEGUARDS AGAINST CONDUCTIVE OBJECTS': '導電物體之安全防護',
        'COMPONENTS': '元件',
        'CONSTRUCTION REQUIREMENTS FOR OUTDOOR ENCLOSURES': '室外外殼構造要求',
        'DETERMINATION OF ACCESSIBLE PARTS': '可接觸部位之判定',
        'DISCONNECT DEVICES': '斷開裝置',
        'ELECTRICALLY-CAUSED INJURY': '電氣導致之傷害',
        'ELECTRICALLY- CAUSED FIRE': '電氣導致之火災',
        'EQUIPMENT CONTAINING BATTERIES AND THEIR PROTECTION CIRCUITS': '含電池及其保護電路之設備',
        'EQUIPMENT MARKINGS, INSTRUCTIONS, AND INSTRUCTIONAL SAFEGUARDS': '設備標示、說明及指示型安全防護',
        'ALTERNATIVE METHOD FOR DETERMINING CLEARANCES FOR INSULATION': '絕緣間隙測定之替代方法',

        # 更多技術術語
        'Classification': '分類',
        'Classification and limits of electrical energy sources': '電氣能量源之分類及限值',
        'Classification of PS and PIS': 'PS 及 PIS 之分類',
        'Classification of potential ignition sources': '潛在點火源之分類',
        'Clearances': '間隙',
        'Coating on components terminals': '元件端子之塗層',
        'Colour of insulation': '絕緣顏色',
        'Component requirements': '元件要求',
        'Components as safeguards': '作為安全防護之元件',
        'Components of safety interlock safeguard mechanism': '安全連鎖防護機構之元件',
        'Compression test': '壓縮試驗',
        'Conditioning': '調節',
        'Conditioning of capacitors and RC units': '電容器及 RC 單元之調節',
        'Conditions for use of a tripping device or a monitoring voltage': '使用跳脫裝置或監測電壓之條件',
        'Connectors': '連接器',
        'Constructional requirements for a fire enclosure and a fire barrier': '防火外殼及防火屏障之構造要求',
        'Constructions and components not specifically covered': '未特別涵蓋之構造及元件',
        'Continuous operation of components': '元件連續運轉',
        'Cord anchorages and strain relief for non-detachable power supply cords': '不可拆卸電源線之線夾及應力消除',
        'Covering of ventilation openings': '通風開口之覆蓋',
        'Creep resistance test': '抗潛變試驗',
        'Discharging current (A)': '放電電流 (A)',
        'Disconnect Device': '斷開裝置',
        'Disconnection from the supply': '與電源斷開',
        'Displacement of a safeguard by an insulating liquid': '絕緣液體對安全防護之位移',
        'Drop test of equipment containing a secondary lithium battery': '含二次鋰電池設備之落下試驗',
        'Durability, legibility and permanence of marking': '標示之耐久性、清晰度及持久性',
        'Capacitors and RC units': '電容器及 RC 單元',
        'Charging safeguards': '充電安全防護',
        'Enamelled winding wire insulation': '漆包繞線絕緣',
        'Endurance requirement': '耐久性要求',
        'Endurance requirements': '耐久性要求',
        'Equipment containing coin/button cell batteries': '含鈕扣電池之設備',
        'Equipment containing work cells with MS3 parts': '含具 MS3 部件工作室之設備',
        'Equipment design and construction': '設備設計與構造',
        'Equipment having electromechanical device for destruction of media': '具破壞媒體之機電裝置設備',
        'Equipment identification markings': '設備識別標示',
        'Equipment markings related to equipment classification': '與設備分類相關之設備標示',
        'Equipment safeguards': '設備安全防護',
        'Equipment set-up, supply connections and earth connections': '設備設定、電源連接及接地連接',
        'Equipment with direct connection to mains': '直接連接主電源之設備',
        'Equipment with multiple supply connections': '具多個電源連接之設備',
        'Equipment without direct connection to mains': '未直接連接主電源之設備',
        'Exceptions to separation between external circuits and earth': '外部電路與接地間隔離之例外',
        'Electronic pulse generator': '電子脈衝產生器',
        'Electrical energy source classification for audio signals': '音頻訊號之電氣能量源分類',

        # F~M 更多術語
        'Functional insulation': '功能絕緣',
        'GENERAL REQUIREMENTS': '一般要求',
        'General classification': '一般分類',
        'General requirement': '一般要求',
        'General requirements': '一般要求',
        'General test requirements': '一般試驗要求',
        'Graphic symbols according to IEC, ISO or manufacturer specific': '依 IEC、ISO 或製造商規定之圖形符號',
        'Humidity conditioning': '濕度調節',
        'Hydrostatic pressure test': '靜水壓試驗',
        'INJURY CAUSED BY HAZARDOUS SUBSTANCES': '危險物質導致之傷害',
        'INSULATED WINDING WIRES FOR USE WITHOUT INTERLEAVED INSULATION': '不使用層間絕緣之絕緣繞線',
        'Impulse test generators': '脈衝試驗產生器',
        'Inadvertent change of operating mode': '操作模式意外改變',
        'Instructional safeguards': '指示型安全防護',
        'Instructions': '說明',
        'Instructions to prevent reasonably foreseeable': '防止合理可預見之說明',
        'Insulating compound forming cemented joints': '形成膠合接合之絕緣化合物',
        'Insulating compound forming solid insulation': '形成固態絕緣之絕緣化合物',
        'Insulating liquid': '絕緣液體',
        'Insulating surfaces': '絕緣表面',
        'Insulation': '絕緣',
        'Insulation between conductors on different surfaces': '不同表面導體間之絕緣',
        'Insulation between conductors on the same inner surface': '同一內表面導體間之絕緣',
        'Insulation in circuits generating starting pulses': '產生起動脈衝電路之絕緣',
        'Insulation in transformers with varying dimensions': '尺寸變化變壓器之絕緣',
        'Insulation materials and requirements': '絕緣材料及要求',
        'Insulation of internal wire as part of supplementary safeguard': '作為補充安全防護一部分之內部導線絕緣',
        'Integrated circuit (IC) current limiters': '積體電路 (IC) 電流限制器',
        'Internal accessible safeguard tests': '內部可接觸安全防護試驗',
        'Likelihood of fire or shock due to entry of conductive object': '導電物體進入導致火災或電擊之可能性',
        'Liquids and liquid filled components (LFC)': '液體及充液元件 (LFC)',
        'MECHANICAL STRENGTH OF CATHODE RAY TUBES (CRT) AND PROTECTION': '陰極射線管 (CRT) 之機械強度及防護',
        'MECHANICAL STRENGTH TESTS': '機械強度試驗',
        'MECHANICALLY-CAUSED INJURY': '機械導致之傷害',
        'Markings and instructions': '標示及說明',
        'Material is non-hygroscopic': '材料非吸濕性',
        'Measurement': '量測',
        'Measurement methods': '量測方法',
        'Measurement of touch current': '觸電電流量測',
        'Measurement of voltage': '電壓量測',
        'Mechanical energy source classifications': '機械能量源分類',
        'Mechanical stability': '機械穩定性',
        'Mechanical strength of enclosures': '外殼機械強度',
        'Mechanical strength test': '機械強度試驗',
        'Mechanically operated safety interlocks': '機械操作之安全連鎖',
        'Metallic parts of outdoor enclosures are resistant to': '室外外殼金屬部件之耐蝕性',
        'Motor overload test conditions': '馬達過載試驗條件',
        'Motors with capacitors': '帶電容器之馬達',
        'No insulation breakdown': '無絕緣擊穿',
        'Non-detachable cord bend protection': '不可拆卸電線彎曲保護',
        'Non-resettable devices suitably rated and marking provided': '適當額定之不可復位裝置並提供標示',
        'Normal operating conditions': '正常工作條件',
        'Oil resistance': '耐油性',
        'Operating and fault conditions': '工作及故障條件',
        'Optocouplers': '光耦合器',
        'Overcurrent protection devices': '過電流保護裝置',
        'PTC thermistors': 'PTC 熱敏電阻',
        'Permanently connected equipment': '永久連接設備',
        'Plugs as disconnect devices': '作為斷開裝置之插頭',
        'Plugs, jacks, connectors tested with blunt probe': '以鈍頭探針測試之插頭、插孔、連接器',
        'Preparation and procedure for the drop test': '落下試驗之準備及程序',
        'Pressurized liquid filled components': '加壓充液元件',
        'Preventing electrolyte spillage': '防止電解液溢出',
        'Procedure 1 for determining clearance': '間隙測定程序 1',
        'Procedure 2 for determining clearance': '間隙測定程序 2',
        'Properties of insulating material': '絕緣材料特性',
        'Prospective touch voltage and touch current associated with external circuits': '與外部電路相關之預期觸電電壓及觸電電流',

        # 7.x 化學危害
        'Instructional safeguard': '指示型安全防護',
        'Instructional Safeguard': '指示型安全防護',

        # 8.x 機械危害
        'Maximum stopping distance from the point of activation (mm)': '從啟動點到停止的最大距離 (mm)',
        'Maximum stopping distance from the point of activation (m)': '從啟動點到停止的最大距離 (m)',
        'Space between end point and nearest fixed mechanical part (mm)': '端點與最近固定機械部件間之距離 (mm)',
        'Space between end point and nearest fixed mechanical part (m)': '端點與最近固定機械部件間之距離 (m)',
        'Mechanical system subjected to 100 000 cycles of operation': '機械系統經受 100,000 次操作循環',
        'MS2 or MS3 part required to be accessible for the function of the equipment': '為設備功能需可接觸之 MS2 或 MS3 部件',
        'Moving MS3 parts only accessible to skilled person': '移動的 MS3 部件僅技術人員可接觸',
        'Personal safeguards and instructions': '人員安全防護及說明',
        'Openings dimensions (mm)': '開口尺寸 (mm)',
        'Test 2, number of attachment points and test force (N)': '試驗 2，附著點數量及試驗力 (N)',

        # 9.x 熱危害

        # 10.x 輻射危害
        'The standard(s) equipment containing laser(s) complies with': '含雷射設備符合之標準',
        'The standard(s) equipment containing laser(s) comply': '含雷射設備符合之標準',
        'Instructional safeguard provided for accessible radiation': '針對可接觸輻射提供指示型安全防護',
        'Instructional safeguard provided for accessible radiation level needs to exceed 1': '針對需超過 1 級之可接觸輻射提供指示型安全防護',
        'Instructional safeguard for skilled persons': '針對技術人員之指示型安全防護',
        'Image projectors': '影像投影機',

        # B~M 附錄
        'Audio Amplifiers and equipment with audio amplifiers': '音頻放大器及含音頻放大器之設備',
        'Audio amplifier abnormal operating conditions': '音頻放大器異常工作狀態',
        'Audio amplifier normal operating conditions': '音頻放大器正常工作狀態',
        'Batteries and their cells comply with relevant IEC': '電池及其電芯符合相關 IEC',
        'Batteries and their protection circuits': '電池及其保護電路',
        'Battery charging and discharging under single fault conditions': '單一故障條件下電池充放電',
        'Battery compartment door/cover construction': '電池室門/蓋構造',
        'Backfeed safeguard in battery backed up supplies': '電池備援供電之反饋安全防護',
        'Additional safeguards for equipment containing a portable secondary lithium': '含可攜式二次鋰電池設備之額外安全防護',
        'Antenna interface test generator': '天線介面測試產生器',
        'Antenna terminal insulation': '天線端子絕緣',
        'Audio signals': '音頻訊號',
        # B 附錄 - 說明性條款
        'Information for safe operation and installation': '安全操作及安裝資訊',
        'Requirements for temperature measurement': '溫度量測要求',
        'a) Information prior to installation and initial use': 'a) 安裝及初次使用前之資訊',
        'c) Instructions for installation and interconnection': 'c) 安裝及互連說明',
        'd) Equipment intended for use only in restricted access area': 'd) 僅供限制進入區域使用之設備',
        'h) Protective conductor current exceeding ES2 limits': 'h) 保護導體電流超過 ES2 限值',
        'j) Permanently connected equipment not provided with all-pole mains switch': 'j) 未提供全極主電源開關之永久連接設備',
        'k) Replaceable components or modules providing safeguard function': 'k) 提供安全防護功能之可更換組件或模組',
        'b) Equipment for use in locations where children normally have': 'b) 用於兒童通常能接觸場所之設備',
        'b) Equipment for use in locations where children not likely to be present': 'b) 用於兒童通常不會出現場所之設備',
        'd) Equipment intended for use only in restricted access locations': 'd) 僅供限制進入場所使用之設備',
        'd) Equipment intended for use only in restricted access areas': 'd) 僅供限制進入區域使用之設備',
        'e) Equipment intended to be fastened in place': 'e) 預定固定安裝之設備',
        'f) Instructions for audio equipment terminals': 'f) 音頻設備端子說明',
        'g) Protective earthing used as a safeguard': 'g) 保護接地作為安全防護',
        'i) Graphic symbols used on equipment': 'i) 設備上使用之圖形符號',
        'j) Permanently connected equipment not provided with all-pole disconnection': 'j) 未提供全極斷開之永久連接設備',
        'k) Replaceable components or modules providing safeguard functions': 'k) 提供安全防護功能之可更換組件或模組',
        'l) Equipment containing insulating liquid': 'l) 含絕緣液體之設備',
        'm) Installation instructions for outdoor equipment': 'm) 室外設備安裝說明',
        # G 附錄 - 元件
        'Thermal cut-outs separately approved according to IEC 60730 series': '依 IEC 60730 系列單獨認證之熱切斷器',
        'Thermal cut-outs tested as part of the equipment as indicated in table G.4': '依表 G.4 作為設備一部分測試之熱切斷器',
        'b) Thermal links tested as part of the equipment': 'b) 作為設備一部分測試之熱熔斷器',
        'Thermal links tested as part of the equipment': '作為設備一部分測試之熱熔斷器',
        # G 附錄 - 更多元件翻譯
        'Thermal cut-outs separately approved according to IEC 60730 with conditions indi': '依 IEC 60730 單獨認證之熱切斷器及條件',
        'Test temperature (C)': '試驗溫度 (°C)',
        'Over current protection by circuit design.': '電路設計之過電流保護。',
        'Maximum Temperature': '最高溫度',
        'Manufacturers\' defined drift': '製造商定義之偏移量',
        'Optocouplers comply with IEC 60747-5-5 with specifics': '光耦合器符合 IEC 60747-5-5 及其細則',
        'Distance through insulation': '穿過絕緣之距離',
        'Number of insulation layers (pcs)': '絕緣層數 (片)',
        'ICX tested separately': 'ICX 單獨測試',
        'Winding wire insulation': '繞線絕緣',
        'Certified triple insulation wire used as secondary winding.': '經認證之三重絕緣導線用於二次繞組。',
        'Solid square and rectangular (flatwise bending) winding wire, cross-sectional ar': '實心方形及矩形 (平向彎曲) 繞線，截面積',
        'Tests and Manufacturing': '試驗與製造',
        'In circuit connected to mains, separation distance': '連接至主電源之電路中，隔離距離',

        # H 附錄 - 變壓器
        'Protection from displacement of windings': '繞組位移防護',
        # K 附錄 - 電容器
        'IC limiter output current (max. 5A)': 'IC 限流器輸出電流 (最大 5A)',
        'ICX with associated circuitry tested in equipment': 'ICX 與相關電路在設備中測試',
        'Smallest capacitance and smallest resistance specified by ICX manufacturer': 'ICX 製造商指定之最小電容與最小電阻',
        'Largest capacitance and smallest resistance for ICX tested by ICX manufacturer': 'ICX 製造商測試之最大電容與最小電阻',
        'Mains voltage that impulses to be superimposed on': '脈衝疊加之主電源電壓',
        # 5.x 電擊防護補充
        'a) Equipment connected to earthed external circuits': 'a) 連接至接地外部電路之設備',
        'a) Equipment connected to earthed external circuit': 'a) 連接至接地外部電路之設備',
        'b) Equipment connected to unearthed external circuits': 'b) 連接至非接地外部電路之設備',
        'b) Equipment connected to unearthed external circuit': 'b) 連接至非接地外部電路之設備',
        'Output +/- to earth': '輸出 +/- 對地',
        # 8.x 機械危害補充
        'MS2 or MS3 part required to be accessible for the purpose of operation': '為操作目的需可接觸之 MS2 或 MS3 部件',

        # 測試參數項目（帶冒號的）
        'Air gap (mm)': '空氣間隙 (mm)',
        'Air gap – distance (mm)': '空氣間隙 - 距離 (mm)',
        'Air gap – electric strength test potential (V)': '空氣間隙 - 耐電壓試驗電壓 (V)',
        'Alternative by electric strength test, tested voltage (V), K': '耐電壓試驗替代方法，試驗電壓 (V)，K',
        'Number of layers (pcs)': '層數 (片)',
        'Relative humidity (%), temperature (°C), duration (h)': '相對濕度 (%)，溫度 (°C)，時間 (h)',
        'Rated operating voltage U (V)': '額定工作電壓 U (V)',
        'Nominal voltage U (V)': '標稱電壓 U (V)',
        'RCD rated residual operating current (mA)': 'RCD 額定剩餘動作電流 (mA)',
        'Protective earthing conductor size (mm2)': '保護接地導體截面積 (mm²)',
        'Acoustic output L , dB(A)': '聲輸出 L, dB(A)',
        'Unweighted RMS output voltage (mV)': '非加權 RMS 輸出電壓 (mV)',
        'Digital output signal (dBFS)': '數位輸出訊號 (dBFS)',
        'Listening device input voltage (mV)': '聆聽裝置輸入電壓 (mV)',
        'Max. acoustic output L , dB(A)': '最大聲輸出 L, dB(A)',
        'Maximum non-clipped output power (W)': '最大無削峰輸出功率 (W)',
        'Open-circuit output voltage (V)': '開路輸出電壓 (V)',
        'Audio output power (W)': '音頻輸出功率 (W)',
        'Audio output voltage (V)': '音頻輸出電壓 (V)',
        'Audio signal source type': '音頻訊號源類型',
        '30 s integrated exposure level (MEL30)': '30 秒積分曝露量 (MEL30)',
        'Appliance inlet cl & cr (mm)': '器具插座 cl & cr (mm)',
        'Arcing PIS': '電弧 PIS',
        'Button/ball diameter (mm)': '按鈕/球直徑 (mm)',
        'Open circuit voltage': '開路電壓',

        # 試驗項目
        'Test 1, additional downwards force (N)': '試驗 1，額外向下力 (N)',
        'Test 2, number of attachment points and test force': '試驗 2，附著點數量及試驗力',
        'Test 3 Nominal diameter (mm) and applied torque (Nm)': '試驗 3 標稱直徑 (mm) 及施加扭矩 (Nm)',
        '- Cable assembly': '- 電纜組件',
        '90V/60Hz Horizontal': '90V/60Hz 水平',
        'ALTERNATIVE METHOD FOR DETERMINING CLEARANCES FOR INSULATION': '絕緣間隙測定之替代方法',
        # G 附錄補充 - 帶結尾點號的項目
        'Samples, material': '樣品，材料',
        'Test time (days per cycle)': '試驗時間 (天/週期)',
        'Test temperature (C)': '試驗溫度 (°C)',
        'Method of protection': '防護方法',
        'Test duration (days)': '試驗時間 (天)',
        'Strain relief test force (N)': '應力消除試驗力 (N)',
        'Radius of curvature after test (mm)': '試驗後曲率半徑 (mm)',
        'Type test voltage V': '型式試驗電壓 V',
        'Routine test voltage, V': '例行試驗電壓，V',
        'See appended table 4.1.2': '見附表 4.1.2',

        # 防護與保護相關
        'Protection against access to hazardous parts': '危險部位接觸防護',
        'Protection against electric shock': '電擊防護',
        'Protection against fire': '火災防護',
        'Protection against mechanical hazards': '機械危害防護',
        'Protection against thermal hazards': '熱危害防護',
        'Protection against radiation hazards': '輻射危害防護',
        'Protective bonding conductor': '保護連接導體',
        'Protective conductor': '保護導體',
        'Protective earth connection': '保護接地連接',
        'Protective earthing': '保護接地',
        'Protective earthing and bonding': '保護接地與連接',
        'Protective earthing terminals and conductors': '保護接地端子及導體',
        'Protective impedance': '保護阻抗',
        'Protective screening': '保護屏蔽',
        'Protective separation': '保護隔離',

        # 元件與零件
        'Power supply': '電源供應器',
        'Power supply cord': '電源線',
        'Power supply unit': '電源供應單元',
        'Power transformer': '電源變壓器',
        'Printed circuit board': '印刷電路板',
        'Printed wiring board': '印刷配線板',
        'Primary circuit': '初級電路',
        'Primary winding': '初級繞組',
        'Secondary circuit': '二次電路',
        'Secondary winding': '二次繞組',
        'Switch': '開關',
        'Switches': '開關',
        'Switch-mode power supply': '交換式電源供應器',
        'Switching device': '開關裝置',

        # 試驗與量測
        'Temperature measurement': '溫度量測',
        'Temperature rise test': '溫升試驗',
        'Temperature test': '溫度試驗',
        'Tensile test': '拉力試驗',
        'Test conditions': '試驗條件',
        'Test equipment': '試驗設備',
        'Test for resistance to heat': '耐熱試驗',
        'Test for resistance to fire': '耐火試驗',
        'Test method': '試驗方法',
        'Test procedure': '試驗程序',
        'Test probe': '測試探針',
        'Test requirements': '試驗要求',
        'Test results': '試驗結果',
        'Test sample': '試驗樣品',
        'Test voltage': '試驗電壓',
        'Touch current': '觸電電流',
        'Touch current measurement': '觸電電流量測',
        'Touch voltage': '觸電電壓',

        # 絕緣相關
        'Basic insulation': '基本絕緣',
        'Supplementary insulation': '補充絕緣',
        'Double insulation': '雙重絕緣',
        'Reinforced insulation': '強化絕緣',
        'Working insulation': '工作絕緣',
        'Insulation coordination': '絕緣配合',
        'Insulation resistance': '絕緣電阻',
        'Insulation test': '絕緣試驗',
        'Solid insulation': '固態絕緣',
        'Creepage distance': '沿面距離',
        'Clearance': '間隙',
        'Creepage distances and clearances': '沿面距離與間隙',
        'Minimum clearance': '最小間隙',
        'Minimum creepage distance': '最小沿面距離',

        # 環境與條件
        'Ambient temperature': '環境溫度',
        'Environmental conditions': '環境條件',
        'Humidity': '濕度',
        'Moisture resistance': '耐濕性',
        'Operating conditions': '操作條件',
        'Operating temperature': '操作溫度',
        'Storage conditions': '儲存條件',
        'Working conditions': '工作條件',

        # 電氣參數
        'AC mains': '交流主電源',
        'DC supply': '直流電源',
        'Input current': '輸入電流',
        'Input power': '輸入功率',
        'Input voltage': '輸入電壓',
        'Leakage current': '漏電流',
        'Load current': '負載電流',
        'Mains frequency': '電源頻率',
        'Mains supply': '主電源',
        'Mains voltage': '主電源電壓',
        'No-load condition': '無載條件',
        'Nominal current': '標稱電流',
        'Nominal power': '標稱功率',
        'Nominal voltage': '標稱電壓',
        'Output current': '輸出電流',
        'Output power': '輸出功率',
        'Output voltage': '輸出電壓',
        'Overcurrent': '過電流',
        'Overvoltage': '過電壓',
        'Peak voltage': '峰值電壓',
        'Rated current': '額定電流',
        'Rated power': '額定功率',
        'Rated voltage': '額定電壓',
        'RMS voltage': '均方根電壓',
        'Short-circuit current': '短路電流',
        'Supply voltage': '供應電壓',
        'Working voltage': '工作電壓',

        # 結構與外殼
        'Accessible part': '可接觸部位',
        'Accessible parts': '可接觸部位',
        'Bottom of enclosure': '外殼底部',
        'Cover': '蓋',
        'Door': '門',
        'Earthing terminal': '接地端子',
        'Enclosure opening': '外殼開口',
        'External surface': '外部表面',
        'Fire enclosure': '防火外殼',
        'Fire barrier': '防火屏障',
        'Housing': '外殼',
        'Internal part': '內部部件',
        'Internal parts': '內部部件',
        'Internal surface': '內部表面',
        'Internal wiring': '內部配線',
        'Live part': '帶電部件',
        'Live parts': '帶電部件',
        'Metal enclosure': '金屬外殼',
        'Mounting': '安裝',
        'Opening': '開口',
        'Openings in enclosures': '外殼開口',
        'Plastic enclosure': '塑膠外殼',
        'Ventilation opening': '通風開口',
        'Ventilation openings': '通風開口',

        # 材料相關
        'Combustible material': '可燃材料',
        'Combustible materials': '可燃材料',
        'Flame retardant': '阻燃劑',
        'Flammability': '可燃性',
        'Flammability rating': '可燃性等級',
        'Flammable material': '易燃材料',
        'Flammable materials': '易燃材料',
        'Insulating material': '絕緣材料',
        'Insulating materials': '絕緣材料',
        'Material group': '材料組別',
        'Non-combustible': '不可燃',
        'Non-flammable': '不易燃',
        'Polymeric material': '聚合物材料',
        'Polymeric materials': '聚合物材料',
        'Thermoplastic material': '熱塑性材料',
        'Thermosetting material': '熱固性材料',

        # 故障與異常
        'Abnormal operation': '異常操作',
        'Abnormal operating condition': '異常操作條件',
        'Abnormal operating conditions': '異常操作條件',
        'Failure': '故障',
        'Failure mode': '故障模式',
        'Fault': '故障',
        'Fault condition': '故障條件',
        'Fault conditions': '故障條件',
        'Malfunction': '故障',
        'Open circuit': '開路',
        'Open-circuit fault': '開路故障',
        'Short circuit': '短路',
        'Short-circuit fault': '短路故障',
        'Single fault': '單一故障',
        'Single fault condition': '單一故障條件',
        'Single fault conditions': '單一故障條件',

        # 人員分類
        'Instructed person': '受指導人員',
        'Instructed persons': '受指導人員',
        'Ordinary person': '普通人員',
        'Ordinary persons': '普通人員',
        'Skilled person': '技術人員',
        'Skilled persons': '技術人員',

        # 設備類型
        'Building-in equipment': '嵌入式設備',
        'Class I equipment': 'Class I 設備',
        'Class II equipment': 'Class II 設備',
        'Class III equipment': 'Class III 設備',
        'Fixed equipment': '固定式設備',
        'Hand-held equipment': '手持式設備',
        'IT equipment': '資訊技術設備',
        'Mobile equipment': '移動式設備',
        'Movable equipment': '可移動式設備',
        'Permanently connected equipment': '永久連接設備',
        'Pluggable equipment': '插接式設備',
        'Portable equipment': '可攜式設備',
        'Stationary equipment': '固定式設備',
        'Transportable equipment': '可運輸設備',

        # 標示與說明
        'Hazard warning': '危害警告',
        'Instruction': '說明',
        'Label': '標籤',
        'Labelling': '標示',
        'Marking': '標示',
        'Marking requirements': '標示要求',
        'Nameplate': '銘牌',
        'Rating label': '額定標籤',
        'Safety instruction': '安全說明',
        'Safety instructions': '安全說明',
        'Safety marking': '安全標示',
        'Safety sign': '安全標誌',
        'Warning': '警告',
        'Warning label': '警告標籤',
        'Warning marking': '警告標示',
        'Warning sign': '警告標誌',

        # 連接與接線
        'Cable entry': '電纜入口',
        'Cable gland': '電纜固定頭',
        'Connection': '連接',
        'Connection terminal': '連接端子',
        'Cord': '線',
        'Cord anchorage': '線夾',
        'Cord connection': '線連接',
        'Cord entry': '線入口',
        'Cord set': '電源線組',
        'Detachable cord': '可拆卸電線',
        'Earthing connection': '接地連接',
        'External cord': '外部電線',
        'External wiring': '外部配線',
        'Interconnection': '互連',
        'Internal cord': '內部電線',
        'Mains connection': '電源連接',
        'Non-detachable cord': '不可拆卸電線',
        'Plug': '插頭',
        'Power cord': '電源線',
        'Socket': '插座',
        'Socket-outlet': '插座',
        'Supply cord': '供電電線',
        'Terminal': '端子',
        'Terminal block': '端子台',
        'Terminals': '端子',
        'Wire': '導線',
        'Wiring': '配線',

        # 安全裝置
        'Circuit breaker': '斷路器',
        'Current limiter': '電流限制器',
        'Disconnect device': '斷開裝置',
        'Earthing device': '接地裝置',
        'Fuse': '保險絲',
        'Fuse holder': '保險絲座',
        'Interlock': '連鎖',
        'Overcurrent protection': '過電流保護',
        'Overcurrent protective device': '過電流保護裝置',
        'Overload protection': '過載保護',
        'Overtemperature protection': '過溫保護',
        'Protective device': '保護裝置',
        'Protective devices': '保護裝置',
        'RCD': '剩餘電流裝置',
        'Residual current device': '剩餘電流裝置',
        'Safety device': '安全裝置',
        'Safety interlock': '安全連鎖',
        'Safety interlocks': '安全連鎖',
        'Thermal cut-out': '熱切斷器',
        'Thermal link': '熱熔斷器',
        'Thermal protection': '熱保護',
        'Thermal protector': '熱保護器',
        'Tripping device': '跳脫裝置',

        # 能量來源
        'Electrical energy source': '電氣能量源',
        'Energy source': '能量源',
        'Hazardous energy source': '危險能量源',
        'Mechanical energy source': '機械能量源',
        'Potential ignition source': '潛在點火源',
        'Thermal energy source': '熱能量源',

        # 審查與驗證
        'Assessment': '評估',
        'Certification': '認證',
        'Compliance': '符合性',
        'Conformity': '符合性',
        'Declaration': '聲明',
        'Declaration of conformity': '符合性聲明',
        'Evaluation': '評估',
        'Inspection': '檢驗',
        'Investigation': '調查',
        'Review': '審查',
        'Type test': '型式試驗',
        'Verification': '驗證',

        # CB 報告特有術語
        'CB Test Certificate': 'CB 測試證書',
        'CB Test Report': 'CB 測試報告',
        'Clause': '條款',
        'Deviations': '偏差',
        'National Differences': '國家差異',
        'Not Applicable': '不適用',
        'Requirement': '要求',
        'Result': '結果',
        'Remark': '備註',
        'Test item': '試驗項目',
        'Test method': '試驗方法',
        'Test report': '試驗報告',
        'Verdict': '判定',

        # 其他常見術語
        'Alternative': '替代',
        'Applicability': '適用性',
        'Application': '應用',
        'Compliance criteria': '符合準則',
        'Definition': '定義',
        'Definitions': '定義',
        'Document': '文件',
        'Documentation': '文件',
        'Equipment': '設備',
        'Example': '範例',
        'Exception': '例外',
        'General': '一般',
        'Guidance': '指引',
        'Information': '資訊',
        'Introduction': '簡介',
        'Limit': '限值',
        'Limits': '限值',
        'Method': '方法',
        'Note': '註',
        'Notes': '註',
        'Objective': '目的',
        'Procedure': '程序',
        'Purpose': '目的',
        'Reference': '參考',
        'References': '參考',
        'Requirement': '要求',
        'Scope': '範圍',
        'Specification': '規格',
        'Specifications': '規格',
        'Standard': '標準',
        'Summary': '摘要',
        'Symbol': '符號',
        'Symbols': '符號',
        'Table': '表',
        'Term': '術語',
        'Terms': '術語',
        'Terms and definitions': '術語與定義',

        # 條款 4-10 章節標題
        'GENERAL REQUIREMENTS': '一般要求',
        'ELECTRICALLY-CAUSED INJURY': '電氣導致之傷害',
        'ELECTRICALLY-CAUSED FIRE': '電氣導致之火災',
        'INJURY CAUSED BY HAZARDOUS SUBSTANCES': '危險物質導致之傷害',
        'MECHANICALLY-CAUSED INJURY': '機械導致之傷害',
        'THERMAL BURN INJURY': '熱灼傷',
        'RADIATION': '輻射',

        # 附錄標題
        'EQUIPMENT MARKINGS, INSTRUCTIONS, AND INSTRUCTIONAL SAFEGUARDS': '設備標示、說明及指示型安全防護',
        'COMPONENTS': '元件',
        'TRANSFORMERS AND INDUCTORS USED AS SAFEGUARDS': '用作安全防護之變壓器及電感器',
        'ADDITIONAL REQUIREMENTS FOR AUDIO/VIDEO EQUIPMENT': '音頻/視頻設備之附加要求',
        'MEASUREMENT OF CREEPAGE DISTANCES AND CLEARANCES': '沿面距離與間隙之量測',
        'TESTS FOR RESISTANCE TO HEAT AND FIRE': '耐熱與耐火試驗',
        'MECHANICAL STRENGTH TESTS': '機械強度試驗',

        # 更多 B 附錄條款
        'Requirements for temperature measurement': '溫度量測要求',
        'Information for safe operation and installation': '安全操作及安裝資訊',
        'Specific requirements for pluggable equipment': '插接式設備之特定要求',
        'Permanently connected equipment without disconnect device': '無斷開裝置之永久連接設備',
        'Equipment with replaceable modules': '具可更換模組之設備',
        'Equipment containing insulating liquid': '含絕緣液體之設備',
        'Instructions for outdoor equipment': '室外設備說明',
        'Symbols used on equipment': '設備上使用之符號',
        'Instructions for equipment with functional earthing': '具功能接地設備之說明',

        # 更多 G 附錄條款 - 元件
        'Capacitor requirements': '電容器要求',
        'Resistor requirements': '電阻器要求',
        'Semiconductor requirements': '半導體要求',
        'Transformer requirements': '變壓器要求',
        'Motor requirements': '馬達要求',
        'Relay requirements': '繼電器要求',
        'Switch requirements': '開關要求',
        'Connector requirements': '連接器要求',
        'Fuse requirements': '保險絲要求',
        'Thermal cut-out requirements': '熱切斷器要求',
        'Thermal link requirements': '熱熔斷器要求',
        'PTC thermistor requirements': 'PTC 熱敏電阻要求',
        'Varistor requirements': '變阻器要求',
        'Optocoupler requirements': '光耦合器要求',

        # H 附錄 - 變壓器
        'Transformer construction': '變壓器構造',
        'Winding insulation': '繞組絕緣',
        'Layer insulation': '層間絕緣',
        'Barrier insulation': '屏障絕緣',
        'Margin': '邊距',
        'Creepage and clearance in transformers': '變壓器之沿面距離與間隙',
        'Transformer marking': '變壓器標示',
        'Transformer tests': '變壓器試驗',
        'Impulse voltage test': '脈衝電壓試驗',

        # J 附錄 - 繞線
        'Winding wire': '繞線',
        'Enamelled wire': '漆包線',
        'Grade 1 wire': '1 級導線',
        'Grade 2 wire': '2 級導線',
        'Grade 3 wire': '3 級導線',
        'Triple insulated wire': '三重絕緣導線',
        'Wire insulation': '導線絕緣',
        'Number of layers': '層數',

        # K 附錄 - 電路
        'Circuit design': '電路設計',
        'Circuit requirements': '電路要求',
        'Current limiting': '電流限制',
        'Current limiting circuit': '電流限制電路',
        'Voltage limiting': '電壓限制',
        'Voltage limiting circuit': '電壓限制電路',
        'Power limiting': '功率限制',
        'Power limiting circuit': '功率限制電路',

        # L 附錄 - 間隙測定
        'Clearance determination': '間隙測定',
        'Creepage distance determination': '沿面距離測定',
        'Pollution degree': '污染等級',
        'Pollution degree 1': '污染等級 1',
        'Pollution degree 2': '污染等級 2',
        'Pollution degree 3': '污染等級 3',
        'Overvoltage category': '過電壓類別',
        'Overvoltage category I': '過電壓類別 I',
        'Overvoltage category II': '過電壓類別 II',
        'Overvoltage category III': '過電壓類別 III',
        'Overvoltage category IV': '過電壓類別 IV',
        'Material group I': '材料組 I',
        'Material group II': '材料組 II',
        'Material group IIIa': '材料組 IIIa',
        'Material group IIIb': '材料組 IIIb',

        # M 附錄 - 電池
        'Battery': '電池',
        'Battery cell': '電池芯',
        'Battery pack': '電池組',
        'Battery charger': '電池充電器',
        'Battery charging': '電池充電',
        'Battery discharging': '電池放電',
        'Rechargeable battery': '可充電電池',
        'Non-rechargeable battery': '不可充電電池',
        'Lithium battery': '鋰電池',
        'Lithium-ion battery': '鋰離子電池',
        'Lead-acid battery': '鉛酸電池',
        'Nickel-cadmium battery': '鎳鎘電池',
        'Nickel-metal hydride battery': '鎳氫電池',
        'Button cell': '鈕扣電池',
        'Coin cell': '鈕扣電池',
        'Battery compartment': '電池室',
        'Battery holder': '電池座',
        'Battery protection': '電池保護',
        'Battery safety': '電池安全',
        'Overcharge protection': '過充電保護',
        'Overdischarge protection': '過放電保護',
        'Short-circuit protection': '短路保護',
        'Battery ventilation': '電池通風',
        'Hydrogen release': '氫氣釋放',

        # 更多測試參數
        'Leakage current (mA)': '漏電流 (mA)',
        'Touch current (mA)': '觸電電流 (mA)',
        'Protective conductor current (mA)': '保護導體電流 (mA)',
        'Electric strength (V)': '耐電壓 (V)',
        'Insulation resistance (MΩ)': '絕緣電阻 (MΩ)',
        'Temperature rise (K)': '溫升 (K)',
        'Temperature rise (°C)': '溫升 (°C)',
        'Applied voltage (V)': '施加電壓 (V)',
        'Test duration (s)': '試驗時間 (s)',
        'Test duration (min)': '試驗時間 (min)',
        'Test duration (h)': '試驗時間 (h)',
        'Force (N)': '力 (N)',
        'Torque (Nm)': '扭矩 (Nm)',
        'Mass (kg)': '質量 (kg)',
        'Dimension (mm)': '尺寸 (mm)',
        'Distance (mm)': '距離 (mm)',
        'Diameter (mm)': '直徑 (mm)',
        'Thickness (mm)': '厚度 (mm)',
        'Width (mm)': '寬度 (mm)',
        'Length (mm)': '長度 (mm)',
        'Height (mm)': '高度 (mm)',
        'Area (mm²)': '面積 (mm²)',
        'Cross-sectional area (mm²)': '截面積 (mm²)',

        # 更多 5.x 條款術語
        'Accessible parts of equipment': '設備可接觸部位',
        'Equipment containing work cells': '含工作室之設備',
        'Accessibility to energy sources': '能量源之可接觸性',
        'Determination of energy source class': '能量源類別之判定',
        'Limits for ES1': 'ES1 限值',
        'Limits for ES2': 'ES2 限值',
        'Limits for ES3': 'ES3 限值',
        'Touch current limits': '觸電電流限值',
        'Protective conductor current limits': '保護導體電流限值',
        'Earthing requirements': '接地要求',
        'Protective earthing requirements': '保護接地要求',
        'Functional earthing requirements': '功能接地要求',
        'Clearance requirements': '間隙要求',
        'Creepage distance requirements': '沿面距離要求',
        'Solid insulation requirements': '固態絕緣要求',
        'Electric strength requirements': '耐電壓要求',
        'Insulation resistance requirements': '絕緣電阻要求',

        # 更多 6.x 條款術語
        'Fire hazard': '火災危害',
        'Fire protection': '火災防護',
        'Fire enclosure requirements': '防火外殼要求',
        'Fire barrier requirements': '防火屏障要求',
        'Flammability requirements': '可燃性要求',
        'Material flammability': '材料可燃性',
        'Flame spread': '火焰蔓延',
        'Flame test': '火焰試驗',
        'Glow wire test': '灼熱絲試驗',
        'Needle flame test': '針焰試驗',
        'Ball pressure test requirements': '球壓試驗要求',
        'Resistance to heat': '耐熱性',
        'Resistance to fire': '耐火性',

        # 更多 8.x 條款術語
        'Mechanical hazard': '機械危害',
        'Moving parts': '移動部件',
        'Sharp edges': '銳邊',
        'Sharp corners': '銳角',
        'Stability': '穩定性',
        'Mechanical stability requirements': '機械穩定性要求',
        'Handle strength': '把手強度',
        'Enclosure strength': '外殼強度',
        'Drop test requirements': '落下試驗要求',
        'Impact test requirements': '撞擊試驗要求',
        'Crush test requirements': '擠壓試驗要求',
        'Steady force test requirements': '穩定力試驗要求',

        # 更多 9.x 條款術語
        'Thermal hazard': '熱危害',
        'Thermal burn': '熱灼傷',
        'Surface temperature': '表面溫度',
        'Temperature limits': '溫度限值',
        'Touch temperature': '觸摸溫度',
        'Hot surface': '高溫表面',
        'Thermal insulation': '熱絕緣',
        'Heat dissipation': '散熱',
        'Cooling': '冷卻',
        'Ventilation': '通風',
        'Thermal protection requirements': '熱保護要求',

        # 更多 10.x 條款術語
        'Radiation hazard': '輻射危害',
        'Electromagnetic radiation': '電磁輻射',
        'Ionizing radiation': '游離輻射',
        'Non-ionizing radiation': '非游離輻射',
        'Laser radiation': '雷射輻射',
        'UV radiation': '紫外線輻射',
        'Infrared radiation': '紅外線輻射',
        'X-ray radiation': 'X 射線輻射',
        'Acoustic radiation': '聲波輻射',
        'Acoustic hazard': '聲波危害',
        'Sound pressure level': '聲壓級',
        'Radiation limits': '輻射限值',
        'Laser class': '雷射等級',
        'Laser safety': '雷射安全',

        # ===== Page 22-27 區域翻譯 (6.x, 7.x, 8.x, 9.x, 10.x 章節) =====

        # 6.4.8.3.4 防火外殼底部可燃性試驗
        'Flammability tests for the bottom of a fire enclosure': '防火外殼底部可燃性試驗',
        'Instructional Safeguard ........................................... :': '指示型安全防護',

        # 6.4.8.3.5 側邊開口
        'Side openings and properties': '側邊開口及特性',
        'Openings dimensions (mm) ................................... :': '開口尺寸 (mm)',

        # 6.4.8.4 防火外殼完整性
        'Integrity of a fire enclosure, condition met: a), b) or c) .........................................': '防火外殼完整性，符合條件：a), b) 或 c)',
        'Integrity of a fire enclosure, condition met: a), b) or c)': '防火外殼完整性，符合條件：a), b) 或 c)',

        # 6.4.8.5 PIS 分離
        'Separation of a PIS from a fire enclosure and a fire barrier distance (mm) or flammability rating ..': 'PIS 與防火外殼及防火屏障之間距 (mm) 或可燃性等級',
        'Separation of a PIS from a fire enclosure and a fire barrier distance (mm) or flammability rating': 'PIS 與防火外殼及防火屏障之間距 (mm) 或可燃性等級',

        # 6.4.10 絕緣液體可燃性
        'Flammability of insulating liquid ............................. :': '絕緣液體可燃性',
        'Flammability of insulating liquid': '絕緣液體可燃性',

        # 6.5 內部及外部配線
        'Internal and external wiring': '內部及外部配線',

        # 6.5.1 一般要求
        # 'General requirements' 已存在

        # 6.5.2 建築配線互連要求
        'Requirements for interconnection to building wiring ................................................': '建築配線互連要求',
        'Requirements for interconnection to building wiring': '建築配線互連要求',

        # 6.5.3 插座內部配線尺寸
        'Internal wiring size (mm2) for socket-outlets .......... :': '插座內部配線尺寸 (mm²)',
        'Internal wiring size (mm2) for socket-outlets': '插座內部配線尺寸 (mm²)',

        # 6.6 連接附加設備之防火安全
        'Safeguards against fire due to the connection to additional equipment': '連接附加設備之防火安全',

        # 7.x 危險物質導致之傷害
        # 'INJURY CAUSED BY HAZARDOUS SUBSTANCES' 已存在

        # 7.2 減少危險物質曝露
        'Reduction of exposure to hazardous substances': '減少危險物質曝露',

        # 7.3 臭氧曝露
        'Ozone exposure': '臭氧曝露',

        # 7.4 使用人員安全防護或個人防護裝備
        'Use of personal safeguards or personal protective equipment (PPE)': '使用人員安全防護或個人防護裝備 (PPE)',
        'Personal safeguards and instructions .................... :': '人員安全防護及說明',

        # 7.5 使用指示型安全防護及說明
        'Use of instructional safeguards and instructions': '使用指示型安全防護及說明',
        'Instructional safeguard (ISO 7010) ........................ :': '指示型安全防護 (ISO 7010)',

        # 7.6 電池及其保護電路
        # 'Batteries and their protection circuits' 已存在

        # 8.x 機械導致之傷害
        # 'MECHANICALLY-CAUSED INJURY' 已存在

        # 8.2 機械能量源分類
        # 'Mechanical energy source classifications' 已存在

        # 8.4 銳邊及尖角之安全防護
        'Safeguards against parts with sharp edges and corners': '銳邊及尖角部件之安全防護',

        # 8.4.1 安全防護
        'Safeguards': '安全防護',

        # 8.4.2 銳邊或尖角
        'Sharp edges or corners': '銳邊或尖角',

        # 8.5 移動部件之安全防護
        'Safeguards against moving parts': '移動部件之安全防護',

        # 8.5.1 手指、珠寶、衣物、頭髮等接觸
        'Fingers, jewellery, clothing, hair, etc., contact with MS2 or MS3 parts': '手指、珠寶、衣物、頭髮等接觸 MS2 或 MS3 部件',

        # 8.5.2 指示型安全防護
        'Instructional safeguard ............................................ :': '指示型安全防護',

        # 8.5.3 含移動部件之特殊類別設備
        'Special categories of equipment containing moving parts': '含移動部件之特殊類別設備',

        # 8.5.4 含 MS3 部件之工作室設備
        # 'Equipment containing work cells with MS3 parts' 已存在

        # 8.5.4.2.1 工作室內人員保護
        'Protection of persons in the work cell': '工作室內人員保護',

        # 8.5.4.2.2 接觸防護覆蓋
        # 'Access protection override' 已存在

        # 8.5.4.2.2.2 覆蓋系統
        'Override system': '覆蓋系統',

        # 8.5.4.2.2.3 視覺指示器
        'Visual indicator': '視覺指示器',

        # 8.5.4.2.2.4 緊急停止系統
        'Emergency stop system': '緊急停止系統',

        # 8.5.4.2.3 啟動點至停止之最大距離
        'Maximum stopping distance from the point of activation (m) .........................................': '啟動點至停止之最大距離 (m)',
        'Space between end point and nearest fixed mechanical part (mm) .....................................': '端點與最近固定機械部件間之空間 (mm)',

        # 8.5.4.2.4 耐久性要求
        # 'Endurance requirements' 已存在
        'Mechanical system subjected to 100 000 cycles of operation': '機械系統經受 100,000 次操作循環',
        '- Cable assembly .................................................... :': '- 電纜組件',

        # 8.5.4.3 含媒體銷毀機電裝置之設備
        'Equipment having electromechanical device for destruction of media': '含媒體銷毀機電裝置之設備',

        # 8.5.4.3.2 設備安全防護
        # 'Equipment safeguards' 已存在

        # 8.5.4.3.3 移動部件之指示型安全防護
        'Instructional safeguards against moving parts ........ :': '移動部件之指示型安全防護',

        # 8.5.4.3.4 與電源斷開
        # 'Disconnection from the supply' 已存在

        # 8.5.4.3.5 切割類型及試驗力
        'Cut type and test force (N) ...................................... :': '切割類型及試驗力 (N)',

        # 8.5.5 高壓燈
        'High pressure lamps': '高壓燈',
        'Explosion test .......................................................... :': '爆炸試驗',
        'Glass particles dimensions (mm) ............................ :': '玻璃微粒尺寸 (mm)',

        # 8.6 設備穩定性
        'Stability of equipment': '設備穩定性',

        # 8.6.2 靜態穩定性
        'Static stability': '靜態穩定性',
        'Static stability test ................................................... :': '靜態穩定性試驗',

        # 8.6.2.2 向下力試驗
        'Downward force test': '向下力試驗',

        # 8.6.3 移動穩定性
        'Relocation stability': '移動穩定性',

        # 8.6.3.3 玻璃滑動試驗
        'Glass slide test': '玻璃滑動試驗',

        # 8.6.3.4 水平力試驗
        'Horizontal force test ................................................ :': '水平力試驗',

        # 8.7 安裝於牆壁、天花板或其他結構之設備
        'Equipment mounted to wall, ceiling or other structure': '安裝於牆壁、天花板或其他結構之設備',

        # 8.7.1 安裝方式類型
        'Mount means type .................................................. :': '安裝方式類型',

        # 8.7.2 試驗方法
        'Test methods': '試驗方法',
        'Test 1, additional downwards force (N) ................... :': '試驗 1，額外向下力 (N)',
        'Test 2, number of attachment points and test force (N) .............................................': '試驗 2，附著點數量及試驗力 (N)',
        'Test 3 Nominal diameter (mm) and applied torque (Nm) ...............................................': '試驗 3 標稱直徑 (mm) 及施加扭矩 (Nm)',

        # 8.8 把手強度
        'Handles strength': '把手強度',

        # 8.8.2 把手強度試驗
        'Handle strength test': '把手強度試驗',

        # 8.9 輪子或腳輪附著要求
        'Wheels or casters attachment requirements': '輪子或腳輪附著要求',

        # 8.9.1 拉力試驗
        'Pull test': '拉力試驗',

        # 8.10 推車、支架及類似載具
        'Carts, stands and similar carriers': '推車、支架及類似載具',

        # 8.10.1 標示及說明
        'Marking and instructions ......................................... :': '標示及說明',

        # 8.10.3 載具負載試驗
        'Cart, stand or carrier loading test': '推車、支架或載具負載試驗',
        'Loading force applied (N) ........................................ :': '施加負載力 (N)',

        # 8.10.4 載具撞擊試驗
        'Cart, stand or carrier impact test': '推車、支架或載具撞擊試驗',

        # 8.10.5 機械穩定性
        # 'Mechanical stability' 已存在

        # 8.10.6 熱塑性材料溫度穩定性
        'Thermoplastic temperature stability': '熱塑性材料溫度穩定性',

        # 8.11 滑軌安裝設備之安裝方式
        'Mounting means for slide-rail mounted equipment (SRME)': '滑軌安裝設備 (SRME) 之安裝方式',

        # 8.11.2 滑軌要求
        'Requirements for slide rails': '滑軌要求',

        # 8.11.3 機械強度試驗
        # 'Mechanical strength test' 已存在
        'Downward force test, force (N) applied ................... :': '向下力試驗，施加力 (N)',

        # 8.11.3.2 側向推力試驗
        'Lateral push force test': '側向推力試驗',

        # 8.11.4 滑軌端擋完整性
        'Integrity of slide rail end stops': '滑軌端擋完整性',

        # 8.12 伸縮或桿狀天線
        'Telescoping or rod antennas': '伸縮或桿狀天線',
        'Button/ball diameter (mm) ....................................... :': '按鈕/球直徑 (mm)',

        # 9.x 熱灼傷
        # 'THERMAL BURN INJURY' 已存在

        # 9.2 熱能量源分類
        'Thermal energy source classifications': '熱能量源分類',

        # 9.3 觸摸溫度限值
        'Touch temperature limits': '觸摸溫度限值',
        'Touch temperatures of accessible parts ................. :': '可接觸部位之觸摸溫度',

        # 9.4 熱能量源之安全防護
        'Safeguards against thermal energy sources': '熱能量源之安全防護',

        # 9.4.1 安全防護要求
        'Requirements for safeguards': '安全防護要求',

        # 9.4.2 設備安全防護
        'Equipment safeguard': '設備安全防護',
        'Instructional safeguard ............................................ :': '指示型安全防護',

        # 9.5 無線電力發射器要求
        'Requirements for wireless power transmitters': '無線電力發射器要求',

        # 9.5.3 外物規格
        'Specification of the foreign objects': '外物規格',
        'Test method and compliance .................................. :': '試驗方法及符合性',

        # 10.x 輻射
        # 'RADIATION' 已存在

        # 10.2 輻射能量源分類
        'Radiation energy source classification': '輻射能量源分類',
        # 'General classification' 已存在

        # 10.2.1 項目
        'Lasers ..................................................................... :': '雷射',
        'Lamps and lamp systems ...................................... :': '燈具及燈具系統',
        'Image projectors ..................................................... :': '影像投影機',
        'X-Ray ...................................................................... :': 'X 射線',
        'Personal music player ............................................ :': '個人音樂播放器',

        # 10.3 雷射輻射之安全防護
        'Safeguards against laser radiation': '雷射輻射之安全防護',
        'The standard(s) equipment containing laser(s) comply ...............................................': '含雷射設備符合之標準',

        # 10.4 燈具及燈具系統光輻射之安全防護
        'Safeguards against optical radiation from lamps and lamp systems (including': '燈具及燈具系統 (含) 光輻射之安全防護',
        'Safeguards against optical radiation from lamps and lamp systems': '燈具及燈具系統光輻射之安全防護',

        # 10.4.1 一般要求
        # 'General requirements' 已存在

        # 10.4.2 超過 1 級輻射之指示型安全防護
        'Instructional safeguard provided for accessible radiation level needs to exceed': '需超過可接觸輻射水平時提供之指示型安全防護',
        'Instructional safeguard provided for accessible radiation level needs to exceed 1': '需超過 1 級可接觸輻射水平時提供之指示型安全防護',

        # ===== 附錄 B-M 補充翻譯 =====

        # B 附錄
        'CONDITION TESTS AND SINGLE FAULT CONDITION TESTS': '狀態試驗及單一故障狀態試驗',
        'NORMAL OPERATING CONDITION TESTS, ABNORMAL OPERATING': '正常操作狀態試驗、異常操作',
        'NORMAL OPERATING CONDITION TESTS, ABNORMAL OPERATING CONDITION TESTS AND SINGLE FAULT CONDITION TESTS': '正常操作狀態試驗、異常操作狀態試驗及單一故障狀態試驗',

        # G 附錄 - 元件
        'Thermal cut-outs separately approved according to IEC 60730 with conditions indicated in Table G.1': '依 IEC 60730 單獨認證之熱切斷器，條件見表 G.1',
        'Thermal cut-outs separately approved according to IEC 60730': '依 IEC 60730 單獨認證之熱切斷器',
        'Test temperature (C)': '試驗溫度 (°C)',
        'Over current protection by circuit design': '電路設計之過電流保護',
        'Over current protection by circuit design.': '電路設計之過電流保護。',
        'Overall diameter or minor overall dimension, D (mm)': '整體直徑或較小整體尺寸 D (mm)',
        'Manufacturers\' defined drift': '製造商定義之漂移',
        'Manufacturers defined drift': '製造商定義之漂移',
        'Approved opto-coupler used': '使用經認證之光耦合器',
        'Approved opto-coupler used.': '使用經認證之光耦合器。',
        'Approved optocoupler used': '使用經認證之光耦合器',
        'Smallest capacitance and smallest resistance specified by ICX': 'ICX 規定之最小電容及最小電阻',
        'Smallest capacitance and smallest resistance specified by IC': 'IC 規定之最小電容及最小電阻',
        'Largest capacitance and smallest resistance for ICX tested by ICX': 'ICX 測試之 ICX 最大電容及最小電阻',
        'Largest capacitance and smallest resistance for ICX tested b': 'ICX 測試之最大電容及最小電阻',
        'Certified triple insulation wire used as secondary winding': '二次繞組使用經認證之三重絕緣導線',
        'Certified triple insulation wire used as secondary winding.': '二次繞組使用經認證之三重絕緣導線。',
        'Solid square and rectangular (flatwise bending) winding wire': '實心方形及矩形（平邊彎曲）繞線',
        'circuit elements': '電路元件',
        'for contact gaps (mm)': '接點間隙 (mm)',

        # K 附錄 - 安全連鎖
        'K.7.2': 'K.7.2',
        'standards': '標準',

        # M 附錄 - 電池
        'battery': '電池',
        '(V); voltage difference during 24 h period (%)': '(V)；24 小時期間電壓差 (%)',
        'misuse': '誤用',

        # O 附錄 - 沿面距離
        'X=1. Pollution degree': 'X=1. 污染等級',
        'Pollution degree': '污染等級',

        # P 附錄 - 導電物件
        'not applicable to transportable equipment': '不適用於可運輸設備',
        'parts': '部件',

        # Q 附錄 - 建築配線
        'b) Impedance limited output': 'b) 阻抗限制輸出',
        'a) Inherently limited output': 'a) 固有限制輸出',
        'c) Regulating network limited output': 'c) 調節網路限制輸出',
        'd) Overcurrent protective device limited output': 'd) 過電流保護裝置限制輸出',
        'e) IC current limiter complying with G.9': 'e) 符合 G.9 之 IC 電流限制器',

        # R 附錄 - 短路試驗
        'where the steady state power does not exceed 4 000 W': '穩態功率不超過 4000 W 時',
        'where the steady state power does not exceed 4000 W': '穩態功率不超過 4000 W 時',

        # S 附錄 - 耐熱耐燃
        'Conditioning (C)': '調節 (°C)',
        'Conditioning': '調節',
        'conditions as set out': '規定之條件',

        # 5.6 保護導體
        'Terminal size for connecting protective bonding conductors (mm)': '連接保護搭接導體之端子尺寸 (mm)',
        'Terminal size for connecting protective earthing conductors (mm)': '連接保護接地導體之端子尺寸 (mm)',
        'Terminal size for connecting protective bonding co': '連接保護搭接導體之端子尺寸',

        # 5.4.2 間隙
        'Alternative by electric strength test, tested voltage (V), K': '耐電壓試驗替代法，測試電壓 (V)，K',
        'Alternative by electric strength test, tested volt': '耐電壓試驗替代法，測試電壓',

        # 5.7.8 外部電路
        'a) Equipment connected to earthed external circuits, current (mA)': 'a) 連接至接地外部電路之設備，電流 (mA)',
        'b) Equipment connected to unearthed external circuits, current (mA)': 'b) 連接至非接地外部電路之設備，電流 (mA)',
        'a) Equipment connected to earthed external circuit': 'a) 連接至接地外部電路之設備',
        'b) Equipment connected to unearthed external circu': 'b) 連接至非接地外部電路之設備',
    }

    if req_normalized in exact_translations:
        return exact_translations[req_normalized]

    # 處理 Single fault 開頭的文字
    if req_normalized.startswith('Single fault'):
        result = req_normalized
        result = result.replace('Single fault', '單一故障')
        result = result.replace(' – SC ', ' – 短路 ')
        result = result.replace(' – OC ', ' – 開路 ')
        result = result.replace(' pin ', ' 腳位 ')
        result = result.replace(' to ', ' 至 ')
        return result

    # 模式匹配翻譯
    patterns = [
        # 帶冒號結尾的項目名稱
        (r'^(.+?)\s*\.{2,}\s*:?\s*$', None),  # 移除結尾的點和冒號，稍後處理
        (r'^Compliance is checked by test\s*\.+\s*:?$', '符合性以試驗檢查'),
        (r'^([\d.]+)\s*mm$', r'\1mm'),
        (r'^Max\.?\s*([\d.]+)\s*Nm$', r'最大 \1Nm'),
        # Instructional Safeguard 變體
        (r'^Instructional\s+[Ss]afeguard\s*\.+\s*:?$', '指示型安全防護'),
        (r'^Instructional\s+[Ss]afeguard\s+\(ISO\s+7010\)\s*\.+\s*:?$', '指示型安全防護 (ISO 7010)'),
    ]

    for pattern, replacement in patterns:
        if replacement:
            m = re.match(pattern, req_normalized, re.IGNORECASE)
            if m:
                return re.sub(pattern, replacement, req_normalized, flags=re.IGNORECASE)

    # 處理帶有大量點號 (...) 的項目 - 這些通常是測試參數
    # 移除結尾的點號和冒號，然後查詢翻譯
    clean_req = re.sub(r'\s*\.{2,}\s*:?\s*$', '', req_normalized).strip()
    if clean_req != req_normalized and clean_req in exact_translations:
        return exact_translations[clean_req]

    # 如果清理後的文字不在字典中，嘗試部分匹配
    partial_translations = {
        'Instructional Safeguard': '指示型安全防護',
        'Instructional safeguard': '指示型安全防護',
        'Air gap': '空氣間隙',
        'Electric strength test': '耐電壓試驗',
        'Ball pressure test': '球壓試驗',
        'Acoustic output': '聲輸出',
        'Audio output': '音頻輸出',
        'Open-circuit output': '開路輸出',
        'Maximum non-clipped output': '最大無削峰輸出',
        'Number of layers': '層數',
        'Relative humidity': '相對濕度',
    }

    for eng, chn in partial_translations.items():
        if clean_req.startswith(eng):
            # 保留後面的參數部分
            rest = clean_req[len(eng):].strip()
            if rest:
                return f'{chn} {rest}'
            return chn

    # 字典未匹配，嘗試 LLM 翻譯
    if HAS_LLM:
        translated = llm_translate(req_normalized)
        if translated != req_normalized:
            return translated

    return req_normalized


def translate_remark(remark: str, clause_id: str) -> str:
    """翻譯常見的 remark 模式"""
    # 正規化：移除換行並壓縮多餘空白（PDF 提取常有換行）
    remark_normalized = ' '.join(remark.split())

    # 精確匹配翻譯（長文本）- 使用正規化後的文字
    exact_translations = {
        '(See appended table 4.1.2)': '(見附表 4.1.2)',
        'The internal wire fixed by riveting, so that a loosen is not likely.': '內部導線以鉚接固定，不易鬆脫。',
        'The external enclosure cannot be opened without damaging the product.': '外殼若不損壞產品則無法開啟。',
        'All safeguards remain effective': '所有防護裝置均維持有效',
        'No such glass used.': '未使用此類玻璃。',
        'No coin/button batteries.': '無鈕扣電池。',
        'No openings.': '無開口。',
        'No ringing signals.': '無振鈴信號。',
        'No audio signals.': '無音頻信號。',
        'Only ES1 circuit can be accessed for this product.': '本產品僅可接觸 ES1 電路。',
        'Considered': '已予考量',
        'Indoor use': '室內使用',
        # 4.7.2 插頭標準說明
        'The US plug according to UL 1310 is used.': '使用符合 UL 1310 之美規插頭。',
        'The Japan plug according to JIS C 8303 is used.': '使用符合 JIS C 8303 之日規插頭。',
        'The EU plug according to EN 50075 is used.': '使用符合 EN 50075 之歐規插頭。',
        'The UK plug according to BS 1363 is used.': '使用符合 BS 1363 之英規插頭。',
        'The AU plug according to AS/NZS 3112 is used.': '使用符合 AS/NZS 3112 之澳規插頭。',
        # 常見測試備註
        'All safeguards remained effectively': '所有安全防護維持有效',
        'All safeguards remain effective': '所有安全防護維持有效',
        'ASRE: All safeguards remained effectively.': 'ASRE: 所有安全防護維持有效。',
        'Indoor use only': '僅限室內使用',
        'Indoor use': '室內使用',
        'Outdoor use': '室外使用',
        'Not applicable': '不適用',
        'See table': '見附表',
        'See clause': '見條款',
        'No opening': '無開口',
        'No openings': '無開口',
        'Not used': '未使用',
        'Not required': '不需要',
        'No such component': '無此類元件',
        'No such device': '無此類裝置',
        'No such material': '無此類材料',
        'No ringing signals': '無振鈴信號',
        'No audio signals': '無音頻信號',
        'No coin cells': '無鈕扣電池',
        'No button cells': '無鈕扣電池',
        'No coin/button batteries': '無鈕扣電池',
        'No coin/button cells': '無鈕扣電池',
        'No laser': '無雷射',
        'No lasers': '無雷射',
        'No moving parts': '無移動部件',
        'No sharp edges': '無銳邊',
        'No sharp corners': '無銳角',
        'No ventilation openings': '無通風開口',
        'No bottom openings': '無底部開口',
        'No thermal cut-offs': '無熱切斷器',
        'No thermal cut-off': '無熱切斷器',
        'No thermal links': '無熱熔斷器',
        'No thermal link': '無熱熔斷器',
        'No appliance outlet provided.': '未提供設備插座。',
        'No appliance outlet provided': '未提供設備插座',
        'The fuses are located within the equipment and not replaceable by an ordinary person or an instructe': '保險絲位於設備內部，非一般人或受指導人員可更換。',
        'The fuses are located within the equipment and not replaceable by an ordinary person or an instructed person.': '保險絲位於設備內部，非一般人或受指導人員可更換。',
        'Not permanently connected equipment': '非永久連接之設備',
        'Not permanently connected equipment.': '非永久連接之設備。',
        'Considered': '已予考量',
        'Class II with functional earthing': 'Class II 帶功能接地',
        'Class II with functional earthing marking': 'Class II 帶功能接地標誌',
        'Class II equipment': 'Class II 設備',
        'Class I equipment': 'Class I 設備',
        'USB-C output': 'USB-C 輸出',
        'USB output': 'USB 輸出',
        'Output + to -': '輸出 + 對 -',
        'Output +/- to earth': '輸出 +/- 對地',
        'Primary to secondary': '初級對二次',
        'Primary to earth': '初級對地',
        'Secondary to earth': '二次對地',
        # 常見測試結果
        'General': '一般',
        'General requirements': '一般要求',
        'Requirements': '要求',
        'Compliance': '符合性',
        'Normal': '正常',
        'Interchangeable': '可互換',
        'Interchangeabl e': '可互換',
        'Test method and compliance': '試驗方法及符合性',
        'Normal, abnormal and fault condition': '正常、異常及故障條件',
        'Normal, abnormal and fault conditions': '正常、異常及故障條件',
        'Worst-case fault': '最差情況故障',
        'Overload': '過載',
        'See below': '見下方',
        'Class B': 'Class B',
        # 常見溫度測試備註
        'Phenolic': '酚醛樹脂',
        # 單一故障測試
        'Single fault': '單一故障',
        'Short circuit': '短路',
        'Open circuit': '開路',
        'SC': '短路',
        'OC': '開路',

        # 附加翻譯 - 語言標記
        'English': '英文',

        # G 附錄 - 元件備註
        'Over current protection by circuit design.': '電路設計之過電流保護。',
        'Over current protection by circuit design': '電路設計之過電流保護',
        'Approved opto-coupler used. (See appended table 4.1.2)': '使用經認證之光耦合器。(見附表 4.1.2)',
        'Approved opto-coupler used.': '使用經認證之光耦合器。',
        'Approved optocoupler used. (See appended table 4.1.2)': '使用經認證之光耦合器。(見附表 4.1.2)',
        'Min. 4000VDC': '最小 4000VDC',
        'Certified triple insulation wire used as secondary winding. (See appended table 4.1.2)': '二次繞組使用經認證之三重絕緣導線。(見附表 4.1.2)',
        'Certified triple insulation wire used as secondary winding.': '二次繞組使用經認證之三重絕緣導線。',
        'Certified triple insulation wire used as secondary winding': '二次繞組使用經認證之三重絕緣導線',

        # O 附錄 - 污染等級
        'X=1. Pollution degree': 'X=1. 污染等級',
        'Pollution degree 1': '污染等級 1',
        'Pollution degree 2': '污染等級 2',
        'Pollution degree 3': '污染等級 3',

        # 防火外殼與防火屏障
        'Approved fire enclosure is used.': '使用經認可之防火外殼。',
        'Approved fire enclosure is used': '使用經認可之防火外殼',
        'V-0 fire enclosure used.': '使用 V-0 防火外殼。',
        'V-0 fire enclosure used': '使用 V-0 防火外殼',
        'No opening and fire barrier': '無開口及防火屏障',
        'No opening and fire barrier.': '無開口及防火屏障。',
        'Separation by a fire barrier': '以防火屏障隔離',
        'Separation by a fire barrier.': '以防火屏障隔離。',
        'Fire enclosure and fire barrier material properties': '防火外殼及防火屏障材料特性',
        'Fire enclosure used': '使用防火外殼',
        'Requirements for a fire enclosure': '防火外殼之要求',

        # 元件與符合性
        'Component requirements': '元件要求',
        'Endurance requirements': '耐久性要求',
        'Tested in the unit': '於本機測試',
        'Tested in the unit.': '於本機測試。',
        'Alternative method': '替代方法',
        'Alternative method.': '替代方法。',
        'Thermal cycling test': '熱循環試驗',
        'Voltage surge test': '電壓突波試驗',
        'Test methods': '試驗方法',
        'Transformers': '變壓器',
        'Optocouplers': '光耦合器',
        'Capacitors and RC units': '電容器及 RC 單元',
        'Supplementary safeguards': '補充性安全防護',
        'Audio amplifier abnormal operating conditions': '音頻放大器異常工作條件',
        'Overload test': '過載試驗',
        'Endurance test': '耐久性試驗',
        'Resistance to corrosion': '耐腐蝕性',

        # 常見試驗結果備註
        'See copy of marking plate': '見標示標籤',
        'See copy of marking plate.': '見標示標籤。',
        'See marking plate': '見銘牌',
        'Test with appliance': '隨機測試',
        'Test with appliance.': '隨機測試。',
        'Tested with appliance': '隨機測試',
        'Tested with appliance.': '隨機測試。',
        'Slot openings tested with wedge probe': '以楔形探針測試槽孔開口',
        'Slot openings tested with wedge probe.': '以楔形探針測試槽孔開口。',
        'No relays': '無繼電器',
        'No relays.': '無繼電器。',
        'No motors': '無馬達',
        'No motors.': '無馬達。',

        # 材料製造商
        'LIANG YI TAPE': '良義膠帶',
    }

    if remark_normalized in exact_translations:
        return exact_translations[remark_normalized]

    # 處理包含 (See appended table X) 開頭的複合備註
    if remark_normalized.startswith('(See appended table'):
        # 提取表格號碼
        m = re.match(r'^\(See appended table ([\d.]+)\)\s*(.*)', remark_normalized, re.DOTALL)
        if m:
            table_num = m.group(1)
            rest = m.group(2).strip()
            if rest:
                # 有附加內容，翻譯常見的附加內容（正規化後比對）
                rest_normalized = ' '.join(rest.split())
                rest_translations = {
                    'Components which are certified to IEC and/or national standards are used correctly within their ratings. Components not covered by IEC standards are tested under the conditions present in the equipment. See also Annex G':
                        '符合 IEC 及/或國家標準認證之組件在其額定值內正確使用。未涵蓋於 IEC 標準之組件在設備實際條件下測試。另見附錄 G',
                }
                if rest_normalized in rest_translations:
                    return f'(見附表 {table_num}) {rest_translations[rest_normalized]}'
                # 如果附加內容太長，只翻譯開頭部分
                return f'(見附表 {table_num}) {rest_normalized[:50]}...' if len(rest_normalized) > 50 else f'(見附表 {table_num}) {rest_normalized}'
            return f'(見附表 {table_num})'

    # 處理 The US plug / Japan plug 等插頭說明（已在 exact_translations 處理）
    # 這裡保留作為後備，處理複雜的多插頭說明
    if 'plug according to' in remark_normalized.lower():
        # 先檢查是否為複雜的多插頭說明
        complex_translations = {
            # MC-601 的複雜 4.7.2 描述
            'The US plug according to UL 1310 Class 2 Power Units (Mechanical Requirements on blades Only); The blade dimension was evaluated to be complied with NEMA configurations in accordance with Wiring Devices- Dimensional Specifications, ANSI/NEMA WD6. Japan plug according to JIS C 8303: 2007.':
                '美規插頭符合 UL 1310 Class 2 電源供應器 (僅限插刀機械要求)；插刀尺寸符合 ANSI/NEMA WD6 配線裝置尺寸規格之 NEMA 配置。日規插頭符合 JIS C 8303: 2007。',
        }
        if remark_normalized in complex_translations:
            return complex_translations[remark_normalized]

        # 嘗試通用模式翻譯
        plug_match = re.match(r'^The\s+(\w+)\s+plug\s+according\s+to\s+(.+?)\s+is\s+used\.?$', remark_normalized, re.IGNORECASE)
        if plug_match:
            region = plug_match.group(1)
            standard = plug_match.group(2)
            region_cn = {'US': '美規', 'Japan': '日規', 'EU': '歐規', 'UK': '英規', 'AU': '澳規', 'China': '中國'}.get(region, region)
            return f'使用符合 {standard} 之{region_cn}插頭。'
        return remark_normalized

    # 常見模式翻譯（使用正規化後的文字）
    patterns = [
        (r'^\(See\s+(?:Annex\s+)?([A-Z][\d.]+(?:,\s*[A-Z]?[\d.]+)*)\)$', r'(見附表 \1)'),
        (r'^\(See\s+(?:Clause\s+)?([A-Z]?[\d.]+(?:,\s*[A-Z]?[\d.]+)*)\)$', r'(見條款 \1)'),
        (r'^\(See Clause ([A-Z]?[\d.]+(?:,\s*[A-Z]?[\d.]+)*)\)$', r'(見條款 \1)'),
        (r'^See\s+(?:appended\s+)?table\s+([A-Z]?[\d.]+)', r'(見附表 \1)'),
        (r'^Indoor\s+use$', '室內使用'),
        (r'^No\s+coin\s+cells\.?$', '無鈕扣電池'),
        (r'^No coin/button batteries\.?$', '無鈕扣電池'),
        (r'^Considered$', '已予考量'),
        (r'^No such glass used\.?$', '未使用此類玻璃'),
        (r'^Evaluation of safeguards.*considered\.?$', '已評估用於限制輸出以符合 ES1 及防止火勢蔓延、機械性與熱灼傷風險之防護措施。'),
        # 數值類（保留 PDF 原值）
        (r'^Max\.?\s*([\d.]+)\s*Nm$', r'最大 \1Nm'),
        (r'^([\d.]+)\s*mm$', r'\1mm'),
    ]

    for pattern, replacement in patterns:
        m = re.match(pattern, remark_normalized, re.IGNORECASE)
        if m:
            return re.sub(pattern, replacement, remark_normalized, flags=re.IGNORECASE)

    # 處理 (See Annex X)clause_id 格式
    annex_match = re.match(r'^\(See Annex ([A-Z](?:\.[0-9.]+)?)\)([0-9A-Z.]+)$', remark_normalized)
    if annex_match:
        return f'(見附錄 {annex_match.group(1)}){annex_match.group(2)}'

    # 處理 (See Clause X, Y)clause_id 格式
    clause_match = re.match(r'^\(See Clause ([A-Z0-9., ]+)\)([0-9A-Z.]+)$', remark_normalized)
    if clause_match:
        return f'(見條款 {clause_match.group(1)}){clause_match.group(2)}'

    # 處理 (See appended table X)clause_id 格式
    table_match = re.match(r'^\(See appended table ([A-Z0-9.]+(?:\s+and\s+[A-Z0-9.]+)?)\)([0-9A-Z.]+)?$', remark_normalized)
    if table_match:
        table_ref = table_match.group(1).replace(' and ', ' 及 ')
        clause_ref = table_match.group(2) or ''
        return f'(見附表 {table_ref}){clause_ref}'

    # 處理 (See appended Tables X and Y) 格式
    tables_match = re.match(r'^\(See appended Tables? ([A-Z0-9., ]+(?:\s+and\s+[A-Z0-9.]+)?)\)$', remark_normalized)
    if tables_match:
        table_ref = tables_match.group(1).replace(' and ', ' 及 ')
        return f'(見附表 {table_ref})'

    # 處理 (See Test Item Particulars...) 格式
    if remark_normalized.startswith('(See Test Item Particulars'):
        return remark_normalized.replace('See Test Item Particulars', '見試驗項目詳情').replace('and appended test tables', '及附加試驗表')

    # 處理 Single fault 開頭的備註
    if remark_normalized.startswith('Single fault'):
        # 翻譯常見模式
        result = remark_normalized
        result = result.replace('Single fault', '單一故障')
        result = result.replace(' – SC ', ' – 短路 ')
        result = result.replace(' – OC ', ' – 開路 ')
        result = result.replace(' pin ', ' 腳位 ')
        result = result.replace(' to ', ' 至 ')
        return result

    # 處理 General / Compliance 等常見詞
    simple_translations = {
        'General': '一般',
        'General requirements': '一般要求',
        'Requirements': '要求',
        'Compliance': '符合性',
        'Normal': '正常',
        'Interchangeable': '可互換',
        'Interchangeabl e': '可互換',
        'Test method and compliance': '試驗方法及符合性',
        'Normal, abnormal and fault condition': '正常、異常及故障條件',
        'Worst-case fault': '最差情況故障',
        'Overload': '過載',
        'See below': '見下方',
    }
    if remark_normalized in simple_translations:
        return simple_translations[remark_normalized]

    # 如果模板中有翻譯，優先使用模板
    if clause_id and clause_id in CLAUSE_TRANSLATIONS:
        template_remark = CLAUSE_TRANSLATIONS[clause_id].get('remark_cn', '')
        if template_remark:
            return template_remark

    # 字典未匹配，嘗試 LLM 翻譯
    if HAS_LLM:
        translated = llm_translate(remark_normalized)
        if translated != remark_normalized:
            return translated

    return remark_normalized


def fill_table_5522(doc: Document, table_5522_data: dict):
    """
    填充 5.5.2.2 電容器存儲放電表格
    """
    if not table_5522_data or 'error' in table_5522_data:
        return

    # 找 5.5.2.2 表格
    target_table = None
    for tbl in doc.tables:
        if tbl.rows and '5.5.2.2' in tbl.rows[0].cells[0].text:
            target_table = tbl
            break

    if not target_table:
        print("警告：找不到 5.5.2.2 表格")
        return

    verdict = table_5522_data.get('verdict', '')
    rows_data = table_5522_data.get('rows', [])

    # 更新 verdict
    if verdict:
        # R0 最後一個 cell 是 verdict
        last_cell = target_table.rows[0].cells[-1]
        if verdict == 'P':
            last_cell.text = '符合'
        elif verdict == 'N/A':
            last_cell.text = '不適用'
        else:
            last_cell.text = verdict

    # 更新資料列（R2, R3 是資料列）
    for i, row_data in enumerate(rows_data[:2]):
        row_idx = i + 2  # R2 開始
        if row_idx < len(target_table.rows):
            row = target_table.rows[row_idx]
            row.cells[0].text = row_data.get('location', '--')
            row.cells[1].text = row_data.get('location', '--')  # 合併儲存格
            row.cells[2].text = row_data.get('supply_voltage', '--')
            row.cells[3].text = row_data.get('condition', '--')
            row.cells[4].text = row_data.get('switch_position', '--')
            row.cells[5].text = row_data.get('measured_voltage', '--')
            row.cells[6].text = row_data.get('es_class', '--')

    # 更新備註列（X 電容值）
    x_cap = table_5522_data.get('x_capacitors', '')
    bleeding = table_5522_data.get('bleeding_resistor', '')
    if x_cap or bleeding:
        # 找備註列
        for row in target_table.rows:
            if '備註' in row.cells[0].text or 'X電容' in row.cells[0].text:
                note_text = f'備註: 用於測試的X電容器: {x_cap}±10%  洩放電阻額定值: {bleeding}'
                row.cells[0].text = note_text
                break

def translate_component_part(part: str) -> str:
    """翻譯 4.1.2 零件名稱"""
    translations = {
        'Plastic enclosure and plug holder': '塑膠外殼及插頭座',
        'Japan plug': '日本插頭',
        'Insulation barrier': '絕緣屏障',
        'PCB': 'PCB',
        'Input wire': '輸入配線',
        'Fuse': '保險絲',
        'Heat shrinkable tube': '熱縮套管',
        'Line choke': '電感',
        'Thermistor': '熱敏電阻',
        'Electrolytic Capacitors': '電解電容',
        'Bridging Rectifier diode': '橋式整流器',
        'Current sensor resistor': '限流電阻',
        'Opto-coupler': '光耦合器',
        'Y-Capacitor': '跨接電容',
        'Transformer': '變壓器',
        'Bobbin': '線架',
        'Magnet wire': '漆包線',
        'Triple insulation wire': '三層絕緣線',
        'Insulation tape': '絕緣膠帶',
        'Coil': '線圈',
        'Insulation system': '絕緣系統',
        '(Alternative)': '(替代)',
        'Interchangeable': '互換',
    }

    result = part
    for eng, chn in translations.items():
        if eng.lower() in result.lower():
            result = result.replace(eng, chn)
            # 也處理大小寫變體
            result = result.replace(eng.lower(), chn)
            result = result.replace(eng.upper(), chn)

    return result


def translate_component_mark(mark: str) -> str:
    """翻譯 4.1.2 認證標誌欄位"""
    if not mark:
        return mark

    # 正規化換行
    mark_norm = ' '.join(mark.split())

    translations = {
        'Test with appliance': '隨機測試',
        'Tested with appliance': '隨機測試',
        'Same as applicant': '與申請者相同',
        'See appended table': '見附表',
        'Interchangeable': '可互換',
    }

    for eng, chn in translations.items():
        if eng.lower() in mark_norm.lower():
            mark_norm = re.sub(re.escape(eng), chn, mark_norm, flags=re.IGNORECASE)

    return mark_norm


def translate_test_observation(obs: str) -> str:
    """翻譯 B.3/B.4 表格的觀察結果欄位"""
    if not obs:
        return obs

    # 正規化換行
    result = obs

    # 詞組翻譯（保持量測值）
    phrase_translations = [
        ('The maximum output current was', '最大輸出電流為'),
        ('The transformer maximum output current was', '變壓器最大輸出電流為'),
        ('when load to', '當負載至'),
        ('the unit shut down immediately', '本機立即關機'),
        ('shut down immediately', '立即關機'),
        ('Recoverable', '可恢復'),
        ('NT, NC, NB', 'NT, NC, NB'),  # 保持縮寫
        ('Prospective Touch Voltage:', '預期接觸電壓:'),
        ('Prospective Touch Voltage', '預期接觸電壓'),
        ('Touch current (output +/- to earth):', '接觸電流 (輸出 +/- 對地):'),
        ('Touch current', '接觸電流'),
        ('Plastic enclosure to earth:', '塑膠外殼對地:'),
        ('Plastic enclosure to earth', '塑膠外殼對地'),
        ('Maximum measured temperature:', '最高量測溫度:'),
        ('Maximum measured temperature', '最高量測溫度'),
        ('open immediately', '立即熔斷'),
        ('F1 open immediately', 'F1 立即熔斷'),
        ('Output port normal load,', '輸出埠正常負載,'),
        ('Output port normal load', '輸出埠正常負載'),
        ('Enclosure outside near', '外殼外部近'),
        ('Plug holder outside', '插頭座外部'),
        ('All safeguards remained effectively', '所有安全防護維持有效'),
        ('ASRE', 'ASRE'),  # 保持縮寫
        ('No insulation breakdown', '無絕緣擊穿'),
        ('Immediately following the humidity conditioning', '濕度調節後立即'),
    ]

    for eng, chn in phrase_translations:
        if eng in result:
            result = result.replace(eng, chn)

    return result


def translate_b34_observations(doc: Document):
    """翻譯 B.3/B.4 異常操作和故障條件試驗表格的觀察結果欄位"""
    # 找出 B.3, B.4 表格
    for tbl_idx, table in enumerate(doc.tables):
        if len(table.rows) < 3:
            continue

        # 檢查表格標題是否包含 B.3, B.4
        first_row_text = ' '.join([c.text for c in table.rows[0].cells])
        if 'B.3' not in first_row_text and 'B.4' not in first_row_text:
            continue
        if '異常操作' not in first_row_text and 'Abnormal' not in first_row_text:
            continue

        # 找出觀察結果欄的索引
        # 通常是 R3 的最後幾欄
        obs_col_indices = []
        if len(table.rows) > 3:
            header_row = table.rows[3]
            for i, cell in enumerate(header_row.cells):
                if '觀察結果' in cell.text or 'Observation' in cell.text:
                    obs_col_indices.append(i)

        if not obs_col_indices:
            # 預設使用最後兩欄
            obs_col_indices = [-2, -1]

        # 翻譯資料列的觀察結果
        translated_count = 0
        for row_idx in range(4, len(table.rows)):  # 從 R4 開始是資料列
            row = table.rows[row_idx]
            for col_idx in obs_col_indices:
                if col_idx < 0:
                    col_idx = len(row.cells) + col_idx
                if 0 <= col_idx < len(row.cells):
                    cell = row.cells[col_idx]
                    original_text = cell.text
                    if original_text and len(original_text) > 10:
                        translated_text = translate_test_observation(original_text)
                        if translated_text != original_text:
                            cell.text = translated_text
                            translated_count += 1

        if translated_count > 0:
            print(f"B.3/B.4 表格 (Table {tbl_idx + 1})：已翻譯 {translated_count} 個觀察結果欄位")


def translate_summary_table(doc: Document):
    """翻譯總覽表格（Table 48）的備註欄位中的英文短語"""
    # 找出總覽表格（特徵：第一列是 Clause | Title | Verdict | Remark）
    target_table = None
    target_table_idx = -1

    for i, tbl in enumerate(doc.tables):
        if len(tbl.rows) >= 2:
            first_row_cells = [c.text.strip() for c in tbl.rows[0].cells]
            if 'Clause' in first_row_cells and 'Title' in first_row_cells:
                target_table = tbl
                target_table_idx = i
                break

    if not target_table:
        return

    # 常見英文短語翻譯
    phrase_translations = [
        ('See appended table', '見附表'),
        ('See copy of marking plate', '見標示標籤'),
        ('Test method and compliance', '試驗方法及符合性'),
        ('Reinforced insulation', '強化絕緣'),
        ('Reinforced safeguard', '強化安全防護'),
        ('Basic insulation', '基本絕緣'),
        ('Supplementary insulation', '補充絕緣'),
        ('Double insulation', '雙重絕緣'),
        ('Functional insulation', '功能絕緣'),
        ('Instructional safeguard', '指示型安全防護'),
        ('Instructional Safeguard', '指示型安全防護'),
        ('Steady force test', '穩定力試驗'),
        ('Control of fire spread in PS', 'PS 火災蔓延控制'),
        ('Control of fire spread', '火災蔓延控制'),
        ('fire enclosure used', '使用防火外殼'),
        ('fire enclosure', '防火外殼'),
        ('fire barrier', '防火屏障'),
        ('Triple insulation wire', '三重絕緣線'),
        ('All circuits except for output circuits', '除輸出電路外之所有電路'),
        ('secondary part circuits', '二次側電路'),
        ('Overload test', '過載試驗'),
        ('No openings', '無開口'),
        ('No opening', '無開口'),
        ('Components which are certified to IEC and/or national standards are used correctly within their ratings',
         '符合 IEC 及/或國家標準認證之組件在其額定值內正確使用'),
        ('See also Annex G', '另見附錄 G'),
        ('Components not covered by IEC standards are tested under the conditions present in the equipment',
         '未涵蓋於 IEC 標準之組件在設備實際條件下測試'),
        ('Printed board:', '印刷電路板:'),
        ('Components other than PCB and wires are:', 'PCB 及導線以外之組件為:'),
        ('Compliance detailed as follows:', '符合性詳述如下:'),
        ('Flammability test for', '可燃性試驗用於'),
        ('fire enclosures and fire barrier materials of equipment', '設備防火外殼及防火屏障材料'),
        ('materials of equipment', '設備材料'),
        ('Same as applicant', '與申請者相同'),
        ('Dti', 'Dti'),  # 保持 Dti 不變
    ]

    translated_count = 0
    for row in target_table.rows[1:]:  # 跳過表頭
        for cell in row.cells:
            original_text = cell.text
            if not original_text or len(original_text) < 5:
                continue

            # 檢查是否有英文需要翻譯
            modified_text = original_text
            for eng, chn in phrase_translations:
                if eng in modified_text:
                    modified_text = modified_text.replace(eng, chn)

            if modified_text != original_text:
                cell.text = modified_text
                translated_count += 1

    if translated_count > 0:
        print(f"總覽表格 (Table {target_table_idx + 1})：已翻譯 {translated_count} 個儲存格")


def _needs_llm_translation(text: str) -> bool:
    """
    判斷文本是否需要 LLM 翻譯
    條件：
    1. 包含連續 3 個以上的英文字母（非專有名詞）
    2. 不是純數字/符號
    3. 不是標準編號（如 IEC 60950-1）
    4. 不是型號/認證編號（如 VDE 40050440, UL E538923）
    5. 不是公司名稱
    6. 不是已經以中文為主的文本
    """
    if not text or len(text.strip()) < 5:
        return False

    # 計算中文字符比例 - 如果超過 30% 是中文，則不需要翻譯
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    total_chars = len(re.sub(r'\s+', '', text))
    if total_chars > 0 and chinese_chars / total_chars > 0.3:
        return False

    # 排除純數字/符號/技術參數格式
    if re.match(r'^[\d\s\.\-\+\/%°℃Ω\(\)\[\],;:>=<≧≦]+$', text):
        return False

    # 排除技術參數格式（如 Dti>0.4 mm, Ext.dcr≧8.0 mm）
    if re.search(r'Dti|Ext\.dcr|Vini|Vpeak|mm,|°C', text):
        return False

    # 排除標準編號（IEC/EN/UL/VDE 等）
    if re.match(r'^(IEC|EN|UL|VDE|TUV|CCC|CSA|CB|CNS)\s*[\d\-\.]+', text.strip()):
        return False

    # 排除認證編號格式
    if re.match(r'^[A-Z]{2,4}\s+[A-Z]?\d{5,}', text.strip()):
        return False

    # 排除電氣參數（如 264 Vac, 50Hz）
    if re.match(r'^[\d\.]+\s*(V|A|W|Hz|kHz|MHz|mA|mV|Vac|Vdc|Vpk|Arms|Apk)', text.strip()):
        return False

    # 排除公司名稱模式（包含 Ltd, Co., Corp, Inc 等）
    if re.search(r'\b(Co\.?\s*,?\s*Ltd|Corp|Inc|GmbH|AG|Pte|B\.?V\.?|S\.?A\.?|LLC)\b', text, re.IGNORECASE):
        return False

    # 排除地名+公司模式（如 ChongQing JinLai Technology）
    if re.search(r'(Technology|Electronics|Electric|Plastics|Chemical|Industrial)\s*(Co|Corp|Ltd|Inc)?', text, re.IGNORECASE):
        return False

    # 檢查是否有需要翻譯的英文（連續 3 個以上英文字母的單詞）
    english_words = re.findall(r'\b[a-zA-Z]{3,}\b', text)

    # 過濾掉常見不需翻譯的專有名詞
    skip_words = {'PCB', 'LED', 'USB', 'AC', 'DC', 'RMS', 'PIN', 'PIS', 'ICX', 'FIW', 'TIW',
                  'SMD', 'MOV', 'NTC', 'PTC', 'EMC', 'ESD', 'LPS', 'RCD', 'CRT', 'AWG',
                  'VDE', 'TUV', 'CSA', 'CCC', 'ENEC', 'CB', 'UL', 'IEC', 'EN', 'CNS', 'JIS',
                  'Vac', 'Vdc', 'Vpk', 'mApk', 'kHz', 'MHz', 'mm', 'kg', 'Co', 'Ltd', 'Inc',
                  'CORP', 'LTD', 'INC', 'CO', 'AG', 'GMBH', 'PTE', 'BV', 'SA', 'SPA',
                  'ELECTRONICS', 'ELECTRONIC', 'TECHNOLOGY', 'PLASTICS', 'CHEMICAL',
                  'INNOVATIVE', 'INDUSTRIAL', 'MATERIALS', 'COMPONENTS', 'INDUSTRY',
                  'Dti', 'Ext', 'dcr', 'Vini', 'Vpeak', 'MAX', 'MIN', 'See', 'pages', 'model',
                  'list', 'details', 'mains', 'Delta', 'Wye'}

    # 過濾掉型號相關的字母組合
    meaningful_words = [w for w in english_words
                        if w.upper() not in skip_words
                        and w.lower() not in {s.lower() for s in skip_words}
                        and not re.match(r'^[A-Z]+\d+', w)  # 排除如 T1, U3, BD1 等
                        and not w.isupper()  # 排除全大寫單詞（通常是公司名/縮寫）
                        and len(w) > 2]  # 排除太短的單詞

    # 如果有 3 個以上有意義的英文單詞，才需要翻譯
    return len(meaningful_words) >= 3


def _apply_llm_translations(doc: Document, candidates: list):
    """
    使用 LLM 批次翻譯候選文本

    Args:
        doc: Word 文件
        candidates: [(tbl_idx, row_idx, cell_idx, text), ...]
    """
    if not candidates:
        return

    translator = get_translator()
    if not translator or not translator.enabled:
        print("[LLM] 翻譯器未啟用，跳過 LLM 翻譯")
        return

    print(f"[LLM] 收集到 {len(candidates)} 個需要智能翻譯的欄位")

    # 批次翻譯
    texts = [item[3] for item in candidates]
    translated = translator.translate_batch(texts)

    # 回寫翻譯結果
    llm_count = 0
    new_translations = {}  # 收集新翻譯，可用於更新字典

    for i, (tbl_idx, row_idx, cell_idx, original_text) in enumerate(candidates):
        if translated[i] != original_text:
            doc.tables[tbl_idx].rows[row_idx].cells[cell_idx].text = translated[i]
            llm_count += 1
            # 記錄新翻譯
            new_translations[original_text.strip()] = translated[i].strip()

    if llm_count > 0:
        print(f"[LLM] 智能翻譯完成：翻譯了 {llm_count} 個欄位")

        # 可選：將新翻譯保存到檔案，以便未來加入字典
        new_trans_path = Path(__file__).parent / 'new_llm_translations.json'
        existing = {}
        if new_trans_path.exists():
            try:
                with open(new_trans_path, 'r', encoding='utf-8') as f:
                    existing = json.load(f)
            except:
                pass
        existing.update(new_translations)
        with open(new_trans_path, 'w', encoding='utf-8') as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        print(f"[LLM] 新翻譯已保存至: {new_trans_path}")


def fill_annex_model_rows(doc: Document, annex_model_rows: list):
    """
    使用 PDF 抽取的 Model 行資料填充 Word 附表中的型號行

    Args:
        doc: Word 文件
        annex_model_rows: PDF 抽取的 Model 行列表
            [{'table_id': '5.2', 'page': 44, 'model_text': 'Model: MC-601 (output: 20.0Vdc, 3.0A)'}, ...]
    """
    if not annex_model_rows:
        return

    # 建立 table_id -> model_texts 的對應（同一個表可能有多個不同的 Model 行）
    table_models = {}
    for mr in annex_model_rows:
        table_id = mr.get('table_id', '')
        model_text = mr.get('model_text', '')
        if table_id and model_text:
            if table_id not in table_models:
                table_models[table_id] = []
            table_models[table_id].append(model_text)

    updated_count = 0

    # 遍歷 Word 文件中的表格
    for tbl_idx, table in enumerate(doc.tables):
        if not table.rows:
            continue

        # 檢查第一行第一格是否是附表 ID
        first_cell_text = table.rows[0].cells[0].text.strip()
        table_id_match = re.match(r'^(\d+\.\d+(?:\.\d+)?)$', first_cell_text)
        if not table_id_match:
            continue

        word_table_id = table_id_match.group(1)
        if word_table_id not in table_models:
            continue

        pdf_models = table_models[word_table_id]
        pdf_model_idx = 0  # 用於追蹤已使用的 PDF Model 行

        # 遍歷表格中的行，尋找型號行
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                text = cell.text.strip()
                # 檢查是否是需要替換的型號行（包含完整規格列表）
                if '型號:' in text and ' or ' in text:
                    # 取得對應的 PDF Model 行
                    if pdf_model_idx < len(pdf_models):
                        pdf_model_text = pdf_models[pdf_model_idx]
                        # 翻譯 PDF Model 行
                        # Model: MC-601 (output: 20.0Vdc, 3.0A) → 型號: MC-601 (輸出: 20.0 Vdc, 3.0 A)
                        translated = translate_model_text(pdf_model_text)
                        cell.text = translated
                        updated_count += 1

            # 檢查這行是否有 Model 行被替換（只在第一個儲存格有時才增加索引）
            first_cell = row.cells[0].text.strip() if row.cells else ''
            if '型號:' in first_cell and 'or' not in first_cell and pdf_model_idx < len(pdf_models):
                pdf_model_idx += 1

    if updated_count > 0:
        print(f"附表 Model 行：已從 PDF 填入 {updated_count} 個儲存格")


def translate_model_text(model_text: str) -> str:
    """
    翻譯 PDF 中的 Model 行文字
    Model: MC-601 (output: 20.0Vdc, 3.0A) → 型號: MC-601 (輸出: 20.0 Vdc, 3.0 A)
    Model: MC-601 (load with 20.0Vdc, 3.0A for...) → 型號: MC-601 (負載 20.0 Vdc, 3.0 A ...)
    """
    if not model_text:
        return model_text

    # 替換基本格式
    result = model_text.replace('Model:', '型號:')

    # output: X.XVdc, Y.YA → 輸出: X.X Vdc, Y.Y A
    output_match = re.search(r'\(output:\s*(\d+\.?\d*)Vdc,\s*(\d+\.?\d*)A\)', result)
    if output_match:
        voltage = output_match.group(1)
        current = output_match.group(2)
        result = re.sub(r'\(output:\s*\d+\.?\d*Vdc,\s*\d+\.?\d*A\)',
                       f'(輸出: {voltage} Vdc, {current} A)', result)
        return result

    # load with X.XVdc, Y.YA for... → 負載 X.X Vdc, Y.Y A ...
    load_match = re.search(r'\(load with\s*(\d+\.?\d*)Vdc,\s*(\d+\.?\d*)A\s*(.*?)\)', result, re.IGNORECASE)
    if load_match:
        voltage = load_match.group(1)
        current = load_match.group(2)
        extra = load_match.group(3).strip()
        # 翻譯常見的附加描述
        extra = extra.replace('for long working', '長時間工作')
        extra = extra.replace('for 5 minutes', '持續5分鐘')
        extra = extra.replace('then reduced to', '然後降至')
        extra = extra.replace('long working as specification', '依規格長時間工作')
        if extra:
            result = f"型號: {result.split('型號:')[1].split('(')[0].strip()}\n(負載{voltage} Vdc, {current}A{extra})"
        else:
            result = f"型號: {result.split('型號:')[1].split('(')[0].strip()} (負載{voltage} Vdc, {current}A)"
        return result

    return result


def translate_all_tables(doc: Document):
    """翻譯所有表格中的通用英文短語"""
    # 通用英文短語翻譯（適用於所有表格）
    phrase_translations = [
        ('Same as applicant', '與申請者相同'),
        ('Test with appliance', '隨機測試'),
        ('Tested with appliance', '隨機測試'),
        ('See appended table', '見附表'),
        ('See copy of marking plate', '見標示標籤'),
        ('Interchangeable', '可互換'),
        ('Interchangeabl e', '可互換'),  # 處理斷字
        ('Interchangeabl\ne', '可互換'),  # 處理換行斷字
        ('T of part/at:', '元件溫度/位置:'),
        ('T of part', '元件溫度'),
        ('All circuits except for output circuits', '除輸出電路外之所有電路'),
        ('secondary part circuits', '二次側電路'),
        ('The circuit connected to AC mains', '連接 AC 主電源之電路'),
        ('Operating surface temperature', '操作表面溫度'),
        ('Current sensor resistor', '電流感測電阻'),
        ('Heat shrinkable tube', '熱縮套管'),
        ('outside triple insulation wire', '三重絕緣線外部'),
        ('Stress relief test', '應力消除試驗'),
        ('Thermal cycling test', '熱循環試驗'),
        ('Voltage surge test', '電壓突波試驗'),
        ('Electric strength test', '耐電壓試驗'),
        ('Capacitors and RC units', '電容器及 RC 單元'),
        ('Prospective touch voltage and touch current', '預期接觸電壓及接觸電流'),
        ('Reduction of the likelihood of ignition', '降低點火可能性'),
        ('for determining clearance', '用於測定間隙'),
        ('mains transient voltage', '主電源暫態電壓'),
        ('See Clause', '見條款'),
        ('See Attachment No', '見附件編號'),
        ('The US plug according to UL', '美規插頭符合 UL'),
        ('Japan plug according to JIS', '日規插頭符合 JIS'),
        ('Mechanical Requirements on blades Only', '僅插刀機械要求'),
        ('The blade dimension was evaluated to', '插刀尺寸經評估為'),
        # 常見單詞翻譯
        ('General requirements', '一般要求'),
        ('General requirement', '一般要求'),
        ('Test method', '試驗方法'),
        ('Compliance', '符合性'),
        ('Requirements', '要求'),
        ('General', '一般'),
        ('Conditioning', '調節'),
        # 處理帶換行的情況
        ('-Triple insulation\nwire', '-三重絕緣線'),
        ('Triple insulation\nwire', '三重絕緣線'),
        # 材料規格翻譯
        ('Double protection', '雙重保護'),
        ('thickness', '厚度'),
        ('Reinforced\ninsulation', '強化\n絕緣'),
        # 條件與輸出
        ('Conditioning (\uf0b0C)', '調節 (°C)'),
        ('Conditioning (°C)', '調節 (°C)'),
        ('Output circuits', '輸出電路'),
        ('Output connector with a shape\nthat insertion into a mains\nconnector or socket is', '輸出連接器之形狀使其插入主電源連接器或插座'),
        ('Output', '輸出'),
        # 元件與試驗術語
        ('Transformers', '變壓器'),
        ('Optocouplers', '光耦合器'),
        ('Relays', '繼電器'),
        ('Resistors', '電阻器'),
        ('Supplementary safeguards', '補充性安全防護'),
        ('Audio amplifier abnormal operating conditions', '音頻放大器異常工作條件'),
        ('Endurance test', '耐久性試驗'),
        ('Tested in the unit', '於本機測試'),
        ('Alternative method', '替代方法'),
        ('- Insulation\nsystem', '- 絕緣\n系統'),
        ('Insulation system', '絕緣系統'),
        ('Tests', '試驗'),
        ('Operating surface\ntemperature:', '操作表面\n溫度:'),
        # 時間與條件
        ('10 mins', '10 分鐘'),
        ('Conditioning (\uf0b0C)', '調節 (°C)'),
    ]

    translated_count = 0
    field_translated_count = 0
    llm_candidates = []  # 收集需要 LLM 翻譯的候選項 [(tbl_idx, row_idx, cell_idx, text)]

    for tbl_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                original_text = cell.text
                if not original_text or len(original_text) < 3:
                    continue

                # 跳過包含 FORMCHECKBOX 的儲存格（避免破壞 checkbox 結構）
                cell_xml = cell._element.xml
                if 'FORMCHECKBOX' in cell_xml:
                    continue

                modified_text = original_text

                # 1. 先嘗試使用智能欄位翻譯（處理帶填充符號的欄位標題）
                field_result = translate_field_title(modified_text)
                if field_result != modified_text:
                    modified_text = field_result
                    field_translated_count += 1
                else:
                    # 2. 使用常規短語翻譯
                    for eng, chn in phrase_translations:
                        if eng in modified_text:
                            modified_text = modified_text.replace(eng, chn)

                # 處理帶有可變長度點號的 Conditioning
                if 'Conditioning (\uf0b0C)' in modified_text:
                    modified_text = re.sub(r'Conditioning \(\uf0b0C\)\s*\.+\s*:?', '調節 (°C) :', modified_text)

                # 3. 處理附表中的「型號」行 - 將完整規格替換成簡潔格式
                # 匹配: 型號: MC-601 (5.0V 3.0A 15.0W or 9.0V ...) → 型號: MC-601 (輸出: 20.0 Vdc, 3.0 A)
                model_match = re.match(r'^(型號:\s*\S+)\s*\(([^)]+or[^)]+)\)$', modified_text.strip())
                if model_match:
                    model_name = model_match.group(1)  # 型號: MC-601
                    full_spec = model_match.group(2)   # 5.0V 3.0A 15.0W or 9.0V ...
                    # 抽取最後一個電壓規格（通常是最大功率的）
                    # 格式: 20.0V 3.0A 60.0W 或 5.0-20.0V 3.0A 60.0W MAX
                    last_spec_match = re.search(r'(\d+\.?\d*)V\s+(\d+\.?\d*)A\s*(?:\d+\.?\d*W)?\s*(?:MAX)?$', full_spec)
                    if last_spec_match:
                        voltage = last_spec_match.group(1)
                        current = last_spec_match.group(2)
                        modified_text = f"{model_name} (輸出: {voltage} Vdc, {current} A)"

                # 4. 格式正規化
                modified_text = normalize_text_format(modified_text)

                if modified_text != original_text:
                    cell.text = modified_text
                    translated_count += 1

                # 4. 檢查是否仍有大量英文（需要 LLM 翻譯）
                if HAS_LLM and _needs_llm_translation(modified_text):
                    llm_candidates.append((tbl_idx, row_idx, cell_idx, modified_text))

    if translated_count > 0:
        print(f"全文件表格：已翻譯 {translated_count} 個儲存格（含 {field_translated_count} 個欄位標題）")

    # 5. 使用 LLM 批次翻譯剩餘英文內容
    if HAS_LLM and llm_candidates:
        _apply_llm_translations(doc, llm_candidates)


def fill_table_412(doc: Document, table_412_data: list):
    """
    從 PDF 提取的 4.1.2 Critical components 數據填充 Word 表格

    Args:
        doc: Word 文件
        table_412_data: list of dict，每個包含 part, manufacturer, model, spec, standard, mark
    """
    if not table_412_data:
        print("警告：4.1.2 表格數據為空")
        return

    # 找到 4.1.2 表格 (Table 45，標題含「重要零件列表」)
    target_table = None
    target_table_idx = -1

    for i, tbl in enumerate(doc.tables):
        if len(tbl.rows) > 2:
            first_row_text = ' '.join([c.text for c in tbl.rows[0].cells])
            if '4.1.2' in first_row_text and '重要零件' in first_row_text:
                target_table = tbl
                target_table_idx = i
                break

    if not target_table:
        print("警告：找不到 4.1.2 重要零件列表表格")
        return

    print(f"找到 4.1.2 表格 (Table {target_table_idx + 1})，原有 {len(target_table.rows)} 行")

    # 保留表頭行（前 2 行：標題行和欄位名稱行）
    header_rows = 2

    # 刪除舊的數據行（從後往前刪除）
    while len(target_table.rows) > header_rows:
        target_table._tbl.remove(target_table.rows[-1]._tr)

    # 添加新的數據行
    for row_data in table_412_data:
        new_row = target_table.add_row()
        cells = new_row.cells

        # 翻譯零件名稱
        part_translated = translate_component_part(row_data.get('part', ''))
        # 翻譯認證標誌
        mark_translated = translate_component_mark(row_data.get('mark', ''))

        # 填入數據
        if len(cells) >= 8:
            cells[0].text = part_translated
            cells[1].text = row_data.get('manufacturer', '')
            cells[2].text = row_data.get('model', '')
            cells[3].text = row_data.get('spec', '')
            cells[4].text = row_data.get('standard', '')
            cells[5].text = mark_translated
            cells[6].text = mark_translated  # 重複認證標誌欄
            cells[7].text = ''  # 索引欄（可選）

    # 添加備註行
    note_row = target_table.add_row()
    note_row.cells[0].text = '備註:'

    print(f"4.1.2 表格：已填入 {len(table_412_data)} 行零件資料")


def fill_table_t7_t8(doc: Document, cb_tables: list):
    """
    從 PDF 提取 T.7 和 T.8 表格數據並填入 Word

    Args:
        doc: Word 文件
        cb_tables: cb_tables_text.json 的內容
    """
    if not cb_tables:
        return

    # 提取 T.7 和 T.8 數據
    t7_rows = []
    t8_rows = []

    for tbl in cb_tables:
        rows = tbl.get('rows', [])
        if not rows:
            continue

        first_row = str(rows[0])

        if 'T.7' in first_row and 'Drop test' in first_row:
            for r in rows[3:]:  # 跳過表頭行
                if r and r[0] and 'Enclosure' in str(r[0]):
                    t7_rows.append({
                        'part': (r[0] or '').replace('#', '').strip(),
                        'material': r[2] if len(r) > 2 else '',
                        'thickness': r[3] if len(r) > 3 else '',
                        'height': r[4] if len(r) > 4 else '',
                        'observation': r[5] if len(r) > 5 else ''
                    })

        if 'T.8' in first_row and 'Stress relief' in first_row:
            for r in rows[3:]:  # 跳過表頭行
                part = (r[0] or '').replace('#', '').strip() if r else ''
                if part and ('Enclosure' in part or 'barrier' in part.lower() or 'Insulation' in part):
                    t8_rows.append({
                        'part': part,
                        'material': r[2] if len(r) > 2 else '',
                        'thickness': r[3] if len(r) > 3 else '',
                        'temperature': r[4] if len(r) > 4 else '',
                        'duration': r[5] if len(r) > 5 else '',
                        'observation': r[6] if len(r) > 6 else ''
                    })

        # 也檢查 T.8 延續表格（page 69，只有數據沒有標題）
        # Table 156 格式: [part, material, thickness, temperature, duration, observation]
        # 必須含有 "#" 前綴才是正確的 T.8 數據行
        for r in rows:
            part_raw = (r[0] or '').strip() if r else ''
            if part_raw.startswith('#') and 'Insulation barrier' in part_raw and len(r) == 6:
                # 檢查是否含有觀察結果關鍵字（排除溫度數據表）
                obs = r[5] if len(r) > 5 else ''
                if obs and ('distortion' in obs.lower() or 'softening' in obs.lower() or 'cracking' in obs.lower()):
                    t8_rows.append({
                        'part': 'Insulation barrier',
                        'material': r[1] if len(r) > 1 else '',
                        'thickness': r[2] if len(r) > 2 else '',
                        'temperature': r[3] if len(r) > 3 else '',
                        'duration': r[4] if len(r) > 4 else '',
                        'observation': obs
                    })

    # 翻譯觀察結果
    def translate_observation(obs: str) -> str:
        obs = (obs or '').replace('\n', ' ')
        translations = {
            'No distortion': '無變形',
            'no distortion': '無變形',
            'no damaged': '無損壞',
            'No damaged': '無損壞',
            'Clearances and creepage were not reduced': '間隙及沿面距離未減少',
            'No softening': '無軟化',
            'no softening': '無軟化',
            'no cracking': '無裂紋',
            'No cracking': '無裂紋',
        }
        result = obs
        for eng, chn in translations.items():
            result = result.replace(eng, chn)
        return result

    # 翻譯部件名稱
    def translate_part(part: str) -> str:
        translations = {
            'Enclosure top': '外殼頂部',
            'Enclosure side': '外殼側邊',
            'Enclosure bottom': '外殼底部',
            'Enclosure': '外殼',
            'Insulation barrier': '絕緣屏障',
        }
        return translations.get(part, part)

    # 填充 T.7 表格
    for i, tbl in enumerate(doc.tables):
        first_row_text = ' '.join([c.text for c in tbl.rows[0].cells]) if tbl.rows else ''
        if 'T.7' in first_row_text and '落下試驗' in first_row_text:
            print(f"找到 T.7 表格 (Table {i+1})")
            # 更新數據行（跳過表頭行 0, 1）
            data_row_idx = 0
            for row_idx in range(2, len(tbl.rows) - 1):  # 最後一行是備註
                row = tbl.rows[row_idx]
                if data_row_idx < len(t7_rows):
                    pdf_data = t7_rows[data_row_idx]
                    cells = row.cells
                    # 更新觀察結果欄（通常是第 5, 6 欄）
                    if len(cells) >= 6:
                        translated_obs = translate_observation(pdf_data.get('observation', ''))
                        cells[5].text = translated_obs
                        if len(cells) > 6:
                            cells[6].text = translated_obs
                    # 更新部件名稱
                    if len(cells) >= 1:
                        cells[0].text = translate_part(pdf_data.get('part', ''))
                        if len(cells) > 1:
                            cells[1].text = translate_part(pdf_data.get('part', ''))
                    data_row_idx += 1
            # 更新備註
            if tbl.rows:
                last_row = tbl.rows[-1]
                last_row.cells[0].text = '備註: 所有外殼材料來源均經過評估試驗，結果相同。'
            print(f"  已更新 {data_row_idx} 行觀察結果")

        # 填充 T.8 表格
        if 'T.8' in first_row_text and '應力釋放試驗' in first_row_text:
            print(f"找到 T.8 表格 (Table {i+1})")
            data_row_idx = 0
            for row_idx in range(2, len(tbl.rows) - 1):  # 最後一行是備註
                row = tbl.rows[row_idx]
                if data_row_idx < len(t8_rows):
                    pdf_data = t8_rows[data_row_idx]
                    cells = row.cells
                    # 更新觀察結果欄（通常是第 6, 7 欄）
                    if len(cells) >= 7:
                        translated_obs = translate_observation(pdf_data.get('observation', ''))
                        cells[6].text = translated_obs
                        if len(cells) > 7:
                            cells[7].text = translated_obs
                    # 更新部件名稱
                    if len(cells) >= 1:
                        cells[0].text = translate_part(pdf_data.get('part', ''))
                        if len(cells) > 1:
                            cells[1].text = translate_part(pdf_data.get('part', ''))
                    data_row_idx += 1
            # 更新備註
            if tbl.rows:
                last_row = tbl.rows[-1]
                last_row.cells[0].text = '備註: 所有外殼及絕緣屏障材料來源均經過評估試驗，結果相同。'
            print(f"  已更新 {data_row_idx} 行觀察結果")


def fill_table_b25(doc: Document, table_b25_data: dict, special_tables: dict):
    """
    確保 B.2.5 表格使用正確的 I rated 值（不是舊案殘留）
    """
    if not table_b25_data or 'error' in table_b25_data:
        return

    i_rated_values = table_b25_data.get('i_rated_values', [])
    if not i_rated_values:
        return

    correct_i_rated = i_rated_values[0]  # 例如 '0.8'

    # 找 B.2.5 或輸入試驗表格
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                # 檢查是否有錯誤的 1.7 值
                if '1.7' in cell.text and ('額定' in cell.text or 'rated' in cell.text.lower()):
                    # 替換為正確值
                    cell.text = cell.text.replace('1.7', correct_i_rated)
                    print(f"已修正 I rated: 1.7 -> {correct_i_rated}")


def translate_paragraph_placeholders(doc):
    """
    替換段落中的佔位符文字
    """
    replacements = {
        '自動生成：安全防護總攬表（由 CB TRF Overview 匯入）': '安全防護總攬表',
        '自動生成：條款判定清單（Clause / Verdict / Remark）': '條款判定清單',
    }

    for para in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in para.text:
                # 替換文字但保留格式
                for run in para.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)


def remove_template_example_tables(doc):
    """
    刪除模板末尾的多餘範例表格
    這些表格包含 Jinja2 標記（如 {% for ... %}），是模板設計時的備用版本
    """
    tables_to_remove = []

    for i, tbl in enumerate(doc.tables):
        if not tbl.rows:
            continue

        # 檢查第一列的內容
        first_row_text = ' '.join([c.text for c in tbl.rows[0].cells])

        # 檢查是否包含未處理的 Jinja2 標記
        if '{%' in first_row_text or '{{' in first_row_text:
            tables_to_remove.append((i, tbl))
            continue

        # 檢查特定的範例表格標誌
        # 1. 能量來源類別表頭（空的範例表格）
        if '能量來源類別' in first_row_text and '身體部位' in first_row_text and len(tbl.rows) <= 2:
            tables_to_remove.append((i, tbl))
            continue

        # 2. Clause Title Verdict 表頭（空的範例表格）
        if 'Clause' in first_row_text and 'Title' in first_row_text and 'Verdict' in first_row_text and len(tbl.rows) <= 2:
            tables_to_remove.append((i, tbl))
            continue

    # 從後往前刪除，避免索引問題
    for i, tbl in reversed(tables_to_remove):
        # 刪除表格的 XML 元素
        tbl._element.getparent().remove(tbl._element)
        print(f"已刪除多餘的範例表格（索引 {i}）")

    if tables_to_remove:
        print(f"共刪除 {len(tables_to_remove)} 個多餘的範例表格")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", default="output/cns_report_data.json")
    ap.add_argument("--template", default="templates/CNS_15598_1_109_template.docx")
    ap.add_argument("--out", default="output/CNS_15598_1_report.docx")
    ap.add_argument("--special_tables", default=None, help="特殊表格 JSON 路徑")
    ap.add_argument("--pdf_clause_rows", default=None, help="PDF 主幹條款列 JSON 路徑")
    ap.add_argument("--table_412", default=None, help="4.1.2 零件表格 JSON 路徑")
    ap.add_argument("--cb_tables", default=None, help="CB 表格原始資料 JSON 路徑 (cb_tables_text.json)")
    ap.add_argument("--annex_model_rows", default=None, help="附表 Model 行 JSON 路徑 (cb_annex_model_rows.json)")
    # 封面欄位（用戶填入）
    ap.add_argument("--cover_report_no", default="", help="封面報告編號")
    ap.add_argument("--cover_applicant_name", default="", help="封面申請者名稱")
    ap.add_argument("--cover_applicant_address", default="", help="封面申請者地址")
    args = ap.parse_args()

    json_path = Path(args.json)
    tpl_path = Path(args.template)
    out_path = Path(args.out)

    if not json_path.exists():
        raise FileNotFoundError(f"JSON not found: {json_path}")
    if not tpl_path.exists():
        raise FileNotFoundError(f"Template not found: {tpl_path}")

    data = load_json(json_path)
    ctx = normalize_context(data)

    # 封面欄位處理：
    # - 如果用戶有填入封面欄位 → 使用用戶填入的值
    # - 如果用戶沒有填入封面欄位 → 保持空白（不使用 PDF 抽取的值）
    # 這三個欄位永遠使用用戶提供的值（即使是空字串）
    ctx['meta']['cb_report_no'] = args.cover_report_no  # 覆蓋報告編號
    ctx['meta']['applicant'] = args.cover_applicant_name  # 覆蓋申請者名稱
    ctx['meta']['applicant_address'] = args.cover_applicant_address  # 覆蓋申請者地址

    if args.cover_report_no:
        print(f"封面報告編號: {args.cover_report_no}")
    if args.cover_applicant_name:
        print(f"封面申請者名稱: {args.cover_applicant_name}")
    if args.cover_applicant_address:
        print(f"封面申請者地址: {args.cover_applicant_address}")

    # 第一階段：docxtpl 渲染
    doc = DocxTemplate(str(tpl_path))
    doc.render(ctx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    # 第二階段：後處理填充特殊表格
    special_tables = {}
    if args.special_tables:
        special_tables_path = Path(args.special_tables)
        if special_tables_path.exists():
            special_tables = load_json(special_tables_path)

    # 重新打開文件進行後處理
    docx = Document(str(out_path))

    # 使用 overview_cb_p12_rows 填充安全防護總攬表（方案A）
    overview_cb_p12_rows = data.get('overview_cb_p12_rows', [])
    if overview_cb_p12_rows:
        rendered_count = fill_overview_table_from_cb_p12(docx, overview_cb_p12_rows)
        print(f"安全防護總攬表：已從 CB p.12 資料填入 {rendered_count} 列")
    else:
        print("警告：overview_cb_p12_rows 不存在，無法填充安全防護總攬表")

    # 填充 5.5.2.2 表格
    if special_tables.get('table_5522'):
        fill_table_5522(docx, special_tables.get('table_5522', {}))

    # 修正 B.2.5 I rated 值
    if special_tables.get('table_b25'):
        fill_table_b25(docx, special_tables.get('table_b25', {}), special_tables)

    # 填充 4.1.2 零件表格（從 PDF 動態提取）
    if args.table_412:
        table_412_path = Path(args.table_412)
        if table_412_path.exists():
            table_412_data = load_json(table_412_path)
            fill_table_412(docx, table_412_data)

    # 填充 T.7 和 T.8 表格（從 PDF 動態提取觀察結果）
    if args.cb_tables:
        cb_tables_path = Path(args.cb_tables)
        if cb_tables_path.exists():
            cb_tables_data = load_json(cb_tables_path)
            fill_table_t7_t8(docx, cb_tables_data)

    # 翻譯 B.3/B.4 表格的觀察結果欄位
    translate_b34_observations(docx)

    # 翻譯總覽表格（Table 48）的英文短語
    translate_summary_table(docx)

    # 填充附表 Model 行（從 PDF 動態提取）
    if args.annex_model_rows:
        annex_model_path = Path(args.annex_model_rows)
        if annex_model_path.exists():
            annex_model_data = load_json(annex_model_path)
            fill_annex_model_rows(docx, annex_model_data)

    # 翻譯所有表格中的通用英文短語
    translate_all_tables(docx)

    # 替換段落佔位符文字
    translate_paragraph_placeholders(docx)

    # 動態更新主條款表格
    # 優先使用 pdf_clause_rows（完全動態生成），否則使用舊版 clauses（只更新）
    pdf_clause_rows = []
    if args.pdf_clause_rows:
        pdf_clause_rows_path = Path(args.pdf_clause_rows)
        if pdf_clause_rows_path.exists():
            pdf_clause_rows = load_json(pdf_clause_rows_path)

    if pdf_clause_rows:
        # 新版：完全動態生成（清空模板列）
        print(f"使用 pdf_clause_rows 完全重建條款表格 ({len(pdf_clause_rows)} 列)...")
        clause_match_result = rebuild_clause_tables_v2(docx, pdf_clause_rows)

        # 儲存條款比對報告
        qa_clause_match_path = out_path.parent / 'qa_clause_table_match.json'
        clause_match_report = {
            'status': 'PASS' if clause_match_result.get('match', False) else 'WARN',
            'pdf_clause_count': clause_match_result.get('pdf_row_count', 0),
            'word_clause_count': clause_match_result.get('word_row_count', 0),
            'pdf_rows_file': str(args.pdf_clause_rows)
        }
        with open(qa_clause_match_path, 'w', encoding='utf-8') as f:
            json.dump(clause_match_report, f, ensure_ascii=False, indent=2)
        print(f"條款比對報告: {qa_clause_match_path}")
    else:
        # 舊版相容
        clauses = data.get('clauses', [])
        if clauses:
            clause_match_result = rebuild_clause_tables(docx, clauses)

            qa_clause_match_path = out_path.parent / 'qa_clause_table_match.json'
            clause_match_report = {
                'status': 'PASS' if clause_match_result.get('match', False) else 'WARN',
                'pdf_clause_count': len(clause_match_result.get('pdf_clause_ids', set())),
                'word_clause_count': len(clause_match_result.get('word_clause_ids', set())),
                'pdf_only': sorted(list(clause_match_result.get('pdf_clause_ids', set()) - clause_match_result.get('word_clause_ids', set())))[:20],
                'word_only': sorted(list(clause_match_result.get('word_clause_ids', set()) - clause_match_result.get('pdf_clause_ids', set())))[:20],
            }
            with open(qa_clause_match_path, 'w', encoding='utf-8') as f:
                json.dump(clause_match_report, f, ensure_ascii=False, indent=2)
            print(f"條款比對報告: {qa_clause_match_path}")
        else:
            print("警告：clauses 為空，跳過條款表格更新")

    # 條款表格重建後，再次翻譯所有表格中的通用英文短語
    # 此函數已整合 LLM 智能翻譯功能（當字典無法匹配時自動調用 LLM）
    translate_all_tables(docx)

    # 刪除模板末尾的多餘範例表格（含 Jinja2 標記）
    remove_template_example_tables(docx)

    # 保存
    docx.save(str(out_path))
    print("已完成特殊表格後處理")

    print("OK")
    print("Rendered:", out_path)

if __name__ == "__main__":
    main()
