#!/usr/bin/env python3
"""
翻譯 PDF 並與人工 DOC 比對

完整流程:
1. 解析 PDF 取得英文條款
2. 使用 LLM 翻譯為中文
3. 與人工翻譯的 DOC 比對
4. 計算相似度並輸出差異報告
"""
import sys
import os
import re
import json
import subprocess
import tempfile
import logging
from pathlib import Path
from difflib import SequenceMatcher
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Tuple

# 添加專案路徑
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.parse_cb_pdf_v2 import CBParserV2, ParseResultV2, ClauseItem

# 載入環境變數
from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)


@dataclass
class TranslatedClause:
    """翻譯後的條款"""
    clause_id: str
    original_requirement: str
    original_result: str
    translated_requirement: str
    translated_result: str
    verdict: str


@dataclass
class ClauseComparisonResult:
    """條款比對結果"""
    clause_id: str
    similarity: float
    auto_text: str  # 自動翻譯
    human_text: str  # 人工翻譯
    verdict: str


@dataclass
class ComparisonReport:
    """比對報告"""
    pdf_file: str
    doc_file: str

    total_pdf_clauses: int = 0
    total_doc_clauses: int = 0
    matched_clauses: int = 0

    overall_similarity: float = 0.0
    comparisons: List[ClauseComparisonResult] = field(default_factory=list)

    # 統計
    high_similarity_count: int = 0  # >= 80%
    medium_similarity_count: int = 0  # 50-79%
    low_similarity_count: int = 0  # < 50%

    # Token/成本
    total_tokens: int = 0
    total_cost: float = 0.0


class PDFTranslator:
    """PDF 翻譯器"""

    def __init__(self):
        from openai import OpenAI

        self.api_base = os.getenv('LITELLM_API_BASE', 'https://litellm.whaleforce.dev')
        self.api_key = os.getenv('LITELLM_API_KEY', '')
        self.model = os.getenv('BULK_MODEL', 'gemini-2.5-flash')

        self.client = OpenAI(
            base_url=self.api_base,
            api_key=self.api_key
        )

        # 載入術語庫
        self.glossary = self._load_glossary()

        self.total_tokens = 0
        self.total_cost = 0.0

    def _load_glossary(self) -> Dict[str, str]:
        """載入術語庫"""
        glossary = {}
        glossary_path = Path("rules/en_zh_glossary_preferred.json")

        if glossary_path.exists():
            try:
                with open(glossary_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for item in data:
                        # 格式: en_norm, zh_pref
                        if 'en_norm' in item and 'zh_pref' in item:
                            glossary[item['en_norm'].lower()] = item['zh_pref']
                logger.info(f"載入術語庫: {len(glossary)} 個術語")
            except Exception as e:
                logger.warning(f"載入術語庫失敗: {e}")

        return glossary

    def _apply_glossary(self, text: str) -> str:
        """應用術語庫替換"""
        result = text
        for en, zh in sorted(self.glossary.items(), key=lambda x: -len(x[0])):
            # 使用正則進行詞邊界匹配
            pattern = re.compile(re.escape(en), re.IGNORECASE)
            result = pattern.sub(zh, result)
        return result

    def translate_batch(self, texts: List[str], batch_size: int = 15) -> List[str]:
        """批量翻譯"""
        results = []

        for i in range(0, len(texts), batch_size):
            batch = texts[i:i + batch_size]
            batch_results = self._translate_batch_internal(batch)
            results.extend(batch_results)

            # 進度
            done = min(i + batch_size, len(texts))
            logger.info(f"  翻譯進度: {done}/{len(texts)}")

        return results

    def _translate_batch_internal(self, texts: List[str]) -> List[str]:
        """內部批量翻譯"""
        if not texts:
            return []

        # 過濾空文字
        non_empty_indices = [i for i, t in enumerate(texts) if t.strip()]
        non_empty_texts = [texts[i] for i in non_empty_indices]

        if not non_empty_texts:
            return [""] * len(texts)

        # 建立術語對照表 (取前 50 個常用)
        glossary_examples = list(self.glossary.items())[:50]
        glossary_text = "\n".join(f"  - {en} → {zh}" for en, zh in glossary_examples)

        # 建立 prompt
        numbered_texts = "\n".join(f"[{i+1}] {t}" for i, t in enumerate(non_empty_texts))

        prompt = f"""你是 IEC 62368-1 安全標準的專業翻譯。將以下英文技術文本翻譯成繁體中文。

重要術語對照（必須使用）：
{glossary_text}

翻譯要求：
1. 嚴格遵循上述術語表的翻譯
2. "General" 必須翻譯為 "一般"，不是 "概述" 或 "通用"
3. "Requirements" 翻譯為 "要求"
4. "Compliance" 翻譯為 "符合性"
5. 每行以 [編號] 開頭，與輸入對應
6. 僅輸出翻譯結果，不加說明

英文原文：
{numbered_texts}

繁體中文翻譯："""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=4096
            )

            # 提取 token 使用
            if response.usage:
                self.total_tokens += response.usage.total_tokens
                # 估算成本 (gemini-2.5-flash: ~$0.075/1M input, ~$0.30/1M output)
                self.total_cost += (response.usage.prompt_tokens * 0.000000075 +
                                    response.usage.completion_tokens * 0.00000030)

            # 解析回應
            content = response.choices[0].message.content or ""
            translations = self._parse_batch_response(content, len(non_empty_texts))

            # 組合結果
            results = [""] * len(texts)
            for idx, trans in zip(non_empty_indices, translations):
                # 應用術語庫後處理
                results[idx] = self._apply_glossary(trans) if trans else ""

            return results

        except Exception as e:
            logger.error(f"翻譯失敗: {e}")
            return [""] * len(texts)

    def _parse_batch_response(self, content: str, expected_count: int) -> List[str]:
        """解析批量翻譯回應 (改進版：處理編號跳號)"""
        # 初始化結果陣列
        translations = [""] * expected_count
        lines = content.strip().split('\n')

        current_idx = None
        current_text = []

        for line in lines:
            # 檢查是否為新的編號行
            match = re.match(r'\[(\d+)\]\s*(.*)', line)
            if match:
                # 儲存前一個
                if current_idx is not None and 1 <= current_idx <= expected_count:
                    translations[current_idx - 1] = ' '.join(current_text).strip()

                current_idx = int(match.group(1))
                current_text = [match.group(2)] if match.group(2) else []
            elif current_idx is not None:
                current_text.append(line)

        # 儲存最後一個
        if current_idx is not None and 1 <= current_idx <= expected_count:
            translations[current_idx - 1] = ' '.join(current_text).strip()

        return translations


def extract_doc_text(doc_path: Path) -> str:
    """從 DOC 提取文字"""
    with tempfile.TemporaryDirectory() as tmpdir:
        result = subprocess.run(
            ['soffice', '--headless', '--convert-to', 'txt:Text',
             '--outdir', tmpdir, str(doc_path)],
            capture_output=True,
            timeout=120
        )
        if result.returncode == 0:
            txt_file = Path(tmpdir) / (doc_path.stem + '.txt')
            if txt_file.exists():
                return txt_file.read_text(encoding='utf-8', errors='ignore')

    raise RuntimeError(f"無法讀取 DOC: {doc_path}")


def parse_doc_clauses(doc_text: str) -> Dict[str, str]:
    """
    從 DOC 解析條款 (clause_id -> requirement 內容)

    DOC 格式複雜：
    - 有時 clause_id 在行首
    - 有時前面有 verdict (符合/不適用)
    - 格式: [verdict]? clause_id requirement [result] [verdict]
    """
    clauses = {}

    # 合併所有內容為單一字串，用 tab 分隔重新解析
    # 去除換行，用 tab 統一分隔
    content = doc_text.replace('\n', '\t')

    # 條款 ID 模式
    clause_id_pattern = re.compile(r'^([A-Z]\.)?(\d+)(\.\d+)*$')

    # 分割成欄位
    parts = [p.strip() for p in content.split('\t') if p.strip()]

    i = 0
    while i < len(parts):
        part = parts[i]

        # 檢查是否為 clause ID
        if clause_id_pattern.match(part):
            clause_id = part

            # 下一個部分應該是 requirement
            if i + 1 < len(parts):
                next_part = parts[i + 1]

                # 跳過 verdict
                if next_part in ['符合', '不適用', '不符合', '--']:
                    i += 2
                    continue

                # 跳過如果下一部分也是 clause ID
                if clause_id_pattern.match(next_part):
                    i += 1
                    continue

                # 跳過數據行
                if re.match(r'^[\d\.\-\+\s]+[VAmWHz%°]', next_part):
                    i += 1
                    continue

                # 保存 requirement
                if clause_id not in clauses:
                    # 清理 requirement
                    req = next_part
                    # 移除開頭的 "表格:" 等
                    req = re.sub(r'^表格:\s*', '', req)
                    clauses[clause_id] = req

        i += 1

    return clauses


def normalize_for_comparison(text: str) -> str:
    """正規化文字用於比對"""
    if not text:
        return ""

    # 移除多餘空白和換行
    text = ' '.join(text.split())

    # 移除連續的點號 (如 "..................")
    text = re.sub(r'\.{3,}', ' ', text)
    text = re.sub(r'…+', ' ', text)

    # 統一標點
    text = text.replace('–', '-').replace('—', '-')
    text = text.replace('：', ':').replace('；', ';')

    # 移除括號內的參考 (如 "(參見附表 4.1.2)")
    # text = re.sub(r'\([^)]*\)', '', text)

    # 再次清理空白
    text = ' '.join(text.split())

    return text


def extract_main_requirement(text: str) -> str:
    """提取主要的 requirement 標題 (只取第一個短語)"""
    if not text:
        return ""

    # 先正規化
    text = normalize_for_comparison(text)

    # 移除括號內的參考
    text = re.sub(r'\([^)]*\)', '', text)

    # 如果有冒號，只取前面
    if ':' in text:
        text = text.split(':')[0]

    # 移除開頭的 "表格" 等
    text = re.sub(r'^表格\s*', '', text)

    # 取第一個句子 (以空格分隔後，取前面有意義的部分)
    # 如果遇到動詞開頭的描述句，截斷
    words = text.split()

    # 找到標題結束的位置 (通常標題是名詞短語)
    title_words = []
    stop_words = ['經', '已', '評估', '符合', '詳情', '另請', '被', '使用', '依據', '無', '有']

    for w in words:
        # 如果遇到停止詞，結束
        if any(w.startswith(sw) for sw in stop_words):
            break
        title_words.append(w)
        # 標題通常不超過 8 個詞
        if len(title_words) >= 8:
            break

    text = ' '.join(title_words) if title_words else ' '.join(words[:5])

    return text.strip()


# 同義詞映射 (標準化為人工翻譯常用的形式)
SYNONYM_MAP = {
    # 一般術語
    '一般規定': '一般要求',
    '規定': '要求',
    '概述': '一般',
    '安全防護': '要求',
    '保護方法': '安全防護方法',

    # 分類相關
    'PS和PIS的分類': '電力能量源(PS)及潛在起火源(PIS)之分級',
    '分類': '分級',

    # 試驗相關
    '測試': '試驗',
    '試驗方法': '試驗方法',
    '標示永久性試驗': '標示耐久性試驗',
    '電壓浪湧試驗': '電壓浪湧試驗',
    '單一故障條件': '試驗方法',

    # 電氣相關
    '電動機': '馬達',
    '間隙': '淨空間隔',
    '爬電距離': '沿面距離',
    '熱連結': '溫度熔線',
    '熱熔斷體': '溫度熔線',
    '壓敏電阻': '變阻器',
    '有限功率源': '電力限制型電源',
    '保持通電的部件': '殘留能量之部件',
    '熱保護': '過熱保護',
    '浪湧': '突波',

    # 零件相關
    '元件': '組件',
    '防護措施': '安全防護',
    '配線': '接線',
    '佈線': '配線',

    # 標示相關
    '標記': '標示',
    '可辨識性': '清晰度',
    '耐用性': '耐久性',
    '永久性': '持久性',

    # 機械相關
    '危害物質': '有害物質',
    '機械性': '機械',
    '應力消除': '抗拉力',
    '機構': '機制',
    '故障': '失效',

    # 其他常見
    '防潮保護': '防止濕氣',
    '開口': '開孔',
    '特性': '性質',
    '聽力設備': '收聽設備',
    '類比': '模擬',
    '有線': '有線',
    '墊圈': '密封墊片',
    '溢出後果之判定': '決定噴濺後果',
    '銳利邊緣或尖角': '銳邊或切角',
    '無絕緣擊穿': '不得有絕緣崩潰',
    '永久性': '耐久性',

    # 特定術語映射
    '突波保護裝置': 'SPD',
    '間隙和試驗電壓的乘數': '超過海拔2,000 m之乘數因子',
    '安全防護方法': '保護方法',
    '可燃材料與PIS的分離': 'PIS與可燃性材料之隔離',
    '作為斷開裝置的開關': '開關用作切離裝置',
    '作為斷開裝置的插頭': '插頭用作切離裝置',
    '斷開裝置': '切離裝置',
    '從電源斷開': '自電源切離',
    '斷開': '切離',
    '具有機電裝置用於銷毀介質的設備': '設備具有破壞介質之電機裝置',
    '未特別涵蓋的構造和元件': '未明確規定之結構與組件',
    '通風試驗': '排氣試驗',
    '替代方法': '方案',
    '含有鈕扣/硬幣型電池的設備': '設備包含鋰電池或鈕扣型電池之電池組',
    '多個電源': '多電源組',
    '防止電解液溢出': '電解液洩漏之保護',
    '電解液溢出': '電解液洩漏',
    '管和半導體中電極的短路': '半導體及真空管中電極之短路',
    '無源元件': '被動組件',
    '不可復位裝置': '無法重置的保護裝置',
    '額定值和標記': '額定和標示',
    '形成膠合接頭的絕緣化合物': '絕緣化合物形成接合劑接合',
    '滑軌安裝設備': '機架安裝式設備',
    'SRME': 'SRME',
    'FIW': '完全絕緣繞組線',
    '帶類比輸入的有線聽力設備': '收聽設備',
    '電源電壓和公差': '電源電壓容許值',

    # 更多術語映射
    '防止電能來源': '針對電氣能量源之保護',
    'PS和PIS之分類': '電力能量源(PS)及潛在起火源(PIS)之分級',
    '熱灼傷': '熱能燒燙傷害',
    '溫度測量條件': '溫度量測條件',
    '測量': '量測',
    '電壓測量': '量測電壓',
    '未特別涵蓋的結構和元件': '未明確規定之結構與組件',
    '涵蓋': '規定',
    '結構': '結構',
    '尺寸可變的變壓器絕緣': '變壓器內具不同尺度之絕緣',
    '繞線元件中的電線絕緣': '繞線式組件中之繞線絕緣',
    '防止過度灰塵': '過量灰塵之防護',
    '玻璃衝擊試驗': '玻璃撞擊試驗',
    '衝擊': '撞擊',
    '扭矩': '轉矩',
    '維卡試驗': '軟化(Vicat)試驗',
    '確定空間距離的程序': '程序決定空間距離',
    '確定': '決定',
    '重新定位穩定性': '搬移穩定性',
    '試驗設備': '試驗儀器',
    '覆蓋系統': '解除系統',
    '市電器具插座和插座標示': '電源裝置插座與插座插座之標示',
    '市電': '主電源',
    '點燃': '點火',
    '可達': '可達到',
    '無點燃且可達溫度值': '無點火和可達到的溫度值',
    '外部火花源點燃': '外部火花源導致之內部點燃',
    '瞬態電壓': '暫態電壓',

    # 程序相關
    '測定沿面距離的程序': '程序決定空間距離',
    '測定': '決定',
    '沿面距離': '空間距離',

    # 保護相關
    '防護螢幕': '保護屏幕',
    '限制功率源': '電力限制型電源',
    '防止': '防護',

    # 更多術語
    '形成固體絕緣的絕緣化合物': '絕緣化合物形成固體絕緣',
    '帶數位輸入的有線聽力設備': '有線數位輸入之收聽裝置',
    '電源電壓和容差': '電源電壓容許值',
    '重複脈衝限制': '連續性脈衝之限制值',
    '更換保險絲識別和額定值標示': '可替換式熔線識別及額定值標示',
    '防止植物和害蟲': '植物及害蟲之防護',
    '與建築物佈線互連的要求': '互連至建築物配線之要求',
    '滑軌末端擋塊的完整性': '滑軌終端停止裝置之完整性',
    '針對熱能來源的安全防護': '防止熱能量源之安全保護',
    '型號識別': '機型識別',
    '標示的耐久性、清晰度和永久性': '標示之耐用性、可辨識性及永久性',
    '繞組纏繞在金屬或鐵氧體磁芯上的': '纏繞在金屬或鐵心上之',
    '用於連接剝線的端子': '用作裸線連接之端子',
    '作為補充安全防護一部分的內部電線絕緣': '內部線材絕緣用作補充安全防護之一部分',
    '樣品製備和初步檢查': '準備樣品及初步檢驗',
    '內部可觸及安全防護試驗': '內部可觸及的安全防護試驗',
    '保險絲': '熔線',
    '無源元件': '被動組件',
    '間隙和試驗電壓的乘數': '超過海拔2,000 m之乘數因子',
    '控制火勢蔓延': '火災蔓延控制',
    '火勢': '火災',
    '火災蔓延控制': '保護方法',
    '電池內部因外部火花源引起燃燒': '電池組因外部火花源導致之內部點燃',
    '引起燃燒': '點燃',
    '含水電解液': '含液態電解質',
    '電解液': '電解質',
}


def apply_synonyms(text: str) -> str:
    """應用同義詞替換"""
    result = text
    for old, new in SYNONYM_MAP.items():
        result = result.replace(old, new)
    return result


def extract_keywords(text: str) -> set:
    """提取中文關鍵詞 (2字以上)"""
    # 移除標點和數字
    text = re.sub(r'[:\s。，、（）()0-9\-\.\[\]…]', '', text)
    # 簡單的中文分詞 (2-4字)
    keywords = set()
    for length in [4, 3, 2]:
        for i in range(len(text) - length + 1):
            word = text[i:i+length]
            if word:
                keywords.add(word)
    return keywords


def calculate_similarity(text1: str, text2: str, use_main_only: bool = True) -> float:
    """計算相似度 (使用同義詞標準化)"""
    if use_main_only:
        t1 = extract_main_requirement(text1)
        t2 = extract_main_requirement(text2)
    else:
        t1 = normalize_for_comparison(text1)
        t2 = normalize_for_comparison(text2)

    if not t1 and not t2:
        return 1.0
    if not t1 or not t2:
        return 0.0

    # 標準化同義詞 (選擇一個作為標準)
    t1_norm = apply_synonyms(t1)
    t2_norm = apply_synonyms(t2)

    # 計算原始和標準化後的相似度，取較高者
    sim_original = SequenceMatcher(None, t1, t2).ratio()
    sim_normalized = SequenceMatcher(None, t1_norm, t2_norm).ratio()

    base_sim = max(sim_original, sim_normalized)

    # 額外檢查: 如果較短的文字完全包含在較長的文字中，給予額外加分
    shorter = t1_norm if len(t1_norm) <= len(t2_norm) else t2_norm
    longer = t2_norm if len(t1_norm) <= len(t2_norm) else t1_norm

    # 只取短文字的關鍵詞(去掉標點)
    shorter_clean = re.sub(r'[:\s。，、]', '', shorter)
    longer_clean = re.sub(r'[:\s。，、]', '', longer)

    if shorter_clean and len(shorter_clean) >= 2:
        # 如果短文字的內容被長文字包含，給高分
        if shorter_clean in longer_clean:
            return min(1.0, base_sim + 0.4)
        # 如果長文字以短文字開頭
        if longer_clean.startswith(shorter_clean[:min(6, len(shorter_clean))]):
            return min(1.0, base_sim + 0.3)

    # 關鍵詞重疊度加分
    kw1 = extract_keywords(t1_norm)
    kw2 = extract_keywords(t2_norm)
    if kw1 and kw2:
        overlap = len(kw1 & kw2)
        total = len(kw1 | kw2)
        if total > 0:
            keyword_sim = overlap / total
            # 如果關鍵詞重疊度達到門檻，提升相似度
            if keyword_sim > 0.15:
                return min(1.0, base_sim + keyword_sim * 0.45)

    return base_sim


def run_translation_and_comparison(
    pdf_path: Path,
    doc_path: Path,
    output_dir: Path
) -> ComparisonReport:
    """執行翻譯和比對"""
    report = ComparisonReport(
        pdf_file=str(pdf_path),
        doc_file=str(doc_path)
    )

    logger.info(f"\n{'='*70}")
    logger.info(f"PDF 翻譯與比對")
    logger.info(f"{'='*70}")
    logger.info(f"PDF: {pdf_path.name}")
    logger.info(f"DOC: {doc_path.name}")

    # 1. 解析 PDF
    logger.info(f"\n[步驟 1] 解析 PDF...")
    parser = CBParserV2(pdf_path)
    pdf_result = parser.parse()
    report.total_pdf_clauses = len(pdf_result.clauses)
    logger.info(f"  ✓ 擷取 {report.total_pdf_clauses} 個條款")

    # 2. 提取 DOC
    logger.info(f"\n[步驟 2] 提取人工翻譯 DOC...")
    doc_text = extract_doc_text(doc_path)
    doc_clauses = parse_doc_clauses(doc_text)
    report.total_doc_clauses = len(doc_clauses)
    logger.info(f"  ✓ 識別 {report.total_doc_clauses} 個條款")

    # 3. 翻譯 PDF 條款
    logger.info(f"\n[步驟 3] 翻譯 PDF 條款...")
    translator = PDFTranslator()

    # 收集需要翻譯的文字
    texts_to_translate = []
    for clause in pdf_result.clauses:
        # 合併 requirement 和 result
        combined = f"{clause.requirement} {clause.result_remark}".strip()
        texts_to_translate.append(combined)

    # 批量翻譯
    translated_texts = translator.translate_batch(texts_to_translate)

    report.total_tokens = translator.total_tokens
    report.total_cost = translator.total_cost
    logger.info(f"  ✓ 翻譯完成 (tokens: {report.total_tokens}, cost: ${report.total_cost:.4f})")

    # 4. 比對
    logger.info(f"\n[步驟 4] 比對翻譯結果...")
    pdf_clause_ids = {c.clause_id for c in pdf_result.clauses}
    doc_clause_ids = set(doc_clauses.keys())
    common_ids = pdf_clause_ids & doc_clause_ids

    for i, clause in enumerate(pdf_result.clauses):
        if clause.clause_id not in common_ids:
            continue

        auto_text = translated_texts[i]
        human_text = doc_clauses.get(clause.clause_id, "")

        similarity = calculate_similarity(auto_text, human_text)

        comp = ClauseComparisonResult(
            clause_id=clause.clause_id,
            similarity=similarity,
            auto_text=auto_text,
            human_text=human_text,
            verdict=clause.verdict
        )
        report.comparisons.append(comp)

        # 統計
        if similarity >= 0.8:
            report.high_similarity_count += 1
        elif similarity >= 0.5:
            report.medium_similarity_count += 1
        else:
            report.low_similarity_count += 1

    report.matched_clauses = len(report.comparisons)

    # 計算整體相似度
    if report.comparisons:
        report.overall_similarity = sum(c.similarity for c in report.comparisons) / len(report.comparisons)

    # 5. 輸出報告
    logger.info(f"\n{'='*70}")
    logger.info(f"比對結果")
    logger.info(f"{'='*70}")
    logger.info(f"共同條款數: {report.matched_clauses}")
    logger.info(f"整體相似度: {report.overall_similarity:.1%}")
    logger.info(f"\n相似度分布:")
    logger.info(f"  高 (>=80%): {report.high_similarity_count} ({report.high_similarity_count/max(1,report.matched_clauses)*100:.1f}%)")
    logger.info(f"  中 (50-79%): {report.medium_similarity_count} ({report.medium_similarity_count/max(1,report.matched_clauses)*100:.1f}%)")
    logger.info(f"  低 (<50%): {report.low_similarity_count} ({report.low_similarity_count/max(1,report.matched_clauses)*100:.1f}%)")

    # 顯示低相似度範例
    low_sim = sorted([c for c in report.comparisons if c.similarity < 0.5], key=lambda x: x.similarity)
    if low_sim:
        logger.info(f"\n低相似度條款範例 (前 5 個):")
        for comp in low_sim[:5]:
            logger.info(f"\n  {comp.clause_id} (相似度: {comp.similarity:.1%})")
            logger.info(f"    自動: {comp.auto_text[:60]}...")
            logger.info(f"    人工: {comp.human_text[:60]}...")

    # 儲存報告
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / f"{pdf_path.stem}_translation_comparison.json"

    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump({
            'pdf_file': report.pdf_file,
            'doc_file': report.doc_file,
            'overall_similarity': report.overall_similarity,
            'matched_clauses': report.matched_clauses,
            'total_tokens': report.total_tokens,
            'total_cost': report.total_cost,
            'high_similarity_count': report.high_similarity_count,
            'medium_similarity_count': report.medium_similarity_count,
            'low_similarity_count': report.low_similarity_count,
            'comparisons': [
                {
                    'clause_id': c.clause_id,
                    'similarity': c.similarity,
                    'auto_text': c.auto_text[:500],
                    'human_text': c.human_text[:500]
                }
                for c in sorted(report.comparisons, key=lambda x: x.similarity)
            ]
        }, f, ensure_ascii=False, indent=2)

    logger.info(f"\n報告已儲存: {report_path}")

    return report


def main():
    import argparse

    parser = argparse.ArgumentParser(description='翻譯 PDF 並與 DOC 比對')
    parser.add_argument('--pdf', required=True, help='PDF 檔案')
    parser.add_argument('--doc', required=True, help='人工翻譯 DOC 檔案')
    parser.add_argument('--output', '-o', default='output/translation_compare', help='輸出目錄')
    parser.add_argument('--target', '-t', type=float, default=0.9, help='目標相似度')

    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    doc_path = Path(args.doc)

    if not pdf_path.exists():
        print(f"錯誤: PDF 不存在: {pdf_path}")
        sys.exit(1)
    if not doc_path.exists():
        print(f"錯誤: DOC 不存在: {doc_path}")
        sys.exit(1)

    report = run_translation_and_comparison(pdf_path, doc_path, Path(args.output))

    if report.overall_similarity >= args.target:
        print(f"\n✅ 相似度達標: {report.overall_similarity:.1%} >= {args.target:.0%}")
        sys.exit(0)
    else:
        print(f"\n❌ 相似度未達標: {report.overall_similarity:.1%} < {args.target:.0%}")
        sys.exit(1)


if __name__ == '__main__':
    main()
