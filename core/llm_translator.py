# core/llm_translator.py
"""
LLM 翻譯模組 - 使用 Azure OpenAI 進行專業安規術語翻譯
支援 IEC 62368-1 & CNS 15598-1 標準術語
支援併發翻譯加速處理
支援 Redis 快取持久化

翻譯策略（純 LLM 模式）:
1. 特殊映射（P, N/A, --）→ 直接映射
2. LLM 翻譯 → 唯一翻譯來源
3. LLM 失敗或未啟用 → 保留原文

二次翻譯機制:
- 第一階段：大片翻譯（分 chunk + 併發）
- 第二階段：掃描殘留英文並重新翻譯

快取機制:
- Redis 可用時：持久化快取（30 天過期）
- Redis 不可用時：記憶體快取（程序重啟後消失）

翻譯規則參考: LLM翻譯術語表.md
"""
import os
import re
import hashlib
from typing import Optional, List, Dict
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time

# 嘗試導入 OpenAI
try:
    from openai import AzureOpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

# 嘗試導入 Redis
try:
    import redis
    HAS_REDIS = True
except ImportError:
    HAS_REDIS = False


# ============================================================
# 強制術語表 (MANDATORY GLOSSARY) - 翻譯時必須使用
# ============================================================
MANDATORY_GLOSSARY = {
    # 零件 / 元件 (Parts / Components)
    'Bleeding resistor': '洩放電阻',
    'Electrolytic capacitor': '電解電容',
    'MOSFET': '電晶體',
    'Current limit resistor': '限流電阻',
    'Varistor': '突波吸收器',
    'MOV': '突波吸收器',
    'Primary wire': '一次側引線',
    'Line choke': '電感',
    'Line chock': '電感',
    'Bobbin': '線架',
    'Plug holder': '刃片插座塑膠材質',
    'AC connector': 'AC 連接器',
    'Fuse': '保險絲',
    'Triple insulated wire': '三層絕緣線',
    'Trace': '銅箔',  # PCB Trace

    # 電路側與繞線 (Circuit Sides & Windings)
    'primary winding': '一次側繞線',
    'primary circuit': '一次側電路',
    'primary': '一次側',
    'secondary': '二次側',
    'Sec.': '二次側',
    'winding': '繞線',
    'core': '鐵芯',
    'magnetic core': '鐵芯',
    'circuit': '電路',
    'connected to': '連接至',

    # 測試條件、環境、狀態
    'Unit shutdown immediately': '設備立即中斷',
    'Unit shutdown': '設備中斷',
    'Ambient': '室溫',
    'Plastic enclosure outside near': '塑膠外殼內側靠近',
    'For model': '適用型號',
    'Optional': '可選',
    'Interchangeable': '不限',
    'Minimum': '至少',
    'at least': '至少',
    'approx.': '約',
    'Approx.': '約',

    # 型號類型
    'For direct plug-in models': '直插式型號',
    'direct plug-in models': '直插式型號',
    'direct plug-in': '直插式',
    'For desktop models': '桌上型型號',
    'desktop models': '桌上型型號',
    'desktop': '桌上型',

    # 判定結果 (Verdict / Result)
    'Pass': '符合',
    'Fail': '不符合',
    'Not applicable': '不適用',

    # 產品類型 (Product Types)
    'SWITCHING MODE POWER SUPPLY': '電源供應器',
    'Switching Mode Power Supply': '電源供應器',
    'switching mode power supply': '電源供應器',
    'POWER SUPPLY': '電源供應器',
    'Power Supply': '電源供應器',
    'power supply': '電源供應器',
}

# 特殊翻譯映射 - 直接映射不經過 LLM
SPECIAL_TRANSLATIONS = {
    'P': '符合',
    'p': '符合',
    'N/A': '不適用',
    '--': '--',
    'F': '不符合',
}

# LLM 拒絕消息開頭 - 遇到這些開頭時替換為 '--'
LLM_REFUSAL_PREFIXES = [
    '抱歉，我無法',
    '抱歉，我不能',
    '對不起，我無法',
    '對不起，我不能',
    '很抱歉，我無法',
    '很抱歉，我不能',
    "I'm sorry",
    'I cannot',
    "I can't",
]

# 英文檢測排除清單 - 這些詞彙不觸發重新翻譯
ENGLISH_EXCLUDE_LIST = {
    'iec', 'en', 'ul', 'csa', 'vde', 'tuv', 'cb', 'ict', 'mosfet', 'pcb',
    'ac', 'dc', 'led', 'usb', 'hdmi', 'wifi', 'http', 'https', 'api',
    'pass', 'fail', 'n/a', 'max', 'min', 'typ', 'nom', 'ref', 'see',
    'table', 'figure', 'note', 'page', 'item', 'model', 'type', 'class',
}


# 系統提示詞 - 專業安規工程師角色（一次翻譯和二次翻譯共用）
SYSTEM_PROMPT = """You are a senior bilingual technical translator. Your ONLY task is to translate from **English to Traditional Chinese (Taiwan)**.

The documents are CB / IEC safety test reports and power electronics specifications. Your translation MUST sound like it was written by an experienced compliance engineer familiar with IEC/EN standards and safety reports used in Taiwan.

### Core rules
1. **Direction:** Always translate **from English to Traditional Chinese**. Never translate Chinese back to English.
2. **Style:**
   - Use formal, concise wording suitable for test reports, specifications, and certification documents.
   - Use clear engineering wording, not marketing language.
   - Keep sentence structure close to the source when it improves traceability in audits or cross-checking.
3. **Formatting & layout:**
   - Preserve tables, item numbers, headings, clause numbers, units, symbols, and values.
   - Do NOT change numbers, limits, dates, test results, verdicts, or standard identifiers.
   - Keep IEC / EN / UL standard codes (e.g., "IEC 62368-1") in English.
   - If the input contains multiple paragraphs separated by special markers like "|||", preserve these markers in your output.

4. **What must remain in English:**
   - Standard names and numbers (IEC/EN/UL/CSA, etc.).
   - Trade names, model names, company names, PCB designators (R1, C2, T1, etc.).
   - Keep abbreviations like "CB", "ICT", "AV" if they are part of standard terminology in the report.

5. **Do NOT leave English untranslated**
   - Except for items listed above, **everything else must be translated into Traditional Chinese**.
   - If you must keep a term in English for technical accuracy, add a clear Traditional Chinese explanation on first occurrence.

### Terminology – MANDATORY glossary (English ➜ Traditional Chinese)
When these English terms or phrases appear, you MUST use EXACTLY the following translations.
Always match the **longest phrase first** (e.g., match "primary winding" before the single word "primary").

Parts / components:
- Bleeding resistor ➜ 洩放電阻
- Electrolytic capacitor ➜ 電解電容
- MOSFET ➜ 電晶體
- Current limit resistor ➜ 限流電阻
- Varistor / MOV ➜ 突波吸收器
- Primary wire ➜ 一次側引線
- Line choke / Line chock ➜ 電感
- Bobbin ➜ 線架
- Plug holder ➜ 刃片插座塑膠材質
- AC connector ➜ AC 連接器
- Fuse ➜ 保險絲
- Triple insulated wire ➜ 三層絕緣線
- Trace (PCB) ➜ 銅箔

Circuit sides & windings:
- primary winding ➜ 一次側繞線
- primary circuit ➜ 一次側電路
- primary (alone, referring to primary side) ➜ 一次側
- secondary ➜ 二次側
- Sec. (abbreviation) ➜ 二次側
- winding (general) ➜ 繞線
- core (magnetic core) ➜ 鐵芯

Test conditions, environment, status:
- Unit shutdown immediately ➜ 設備立即中斷
- Unit shutdown ➜ 設備中斷
- Ambient (temperature, condition) ➜ 室溫
- Plastic enclosure outside near ➜ 塑膠外殼內側靠近
- For model ➜ 適用型號
- Optional ➜ 可選
- Interchangeable ➜ 不限
- Minimum / at least ➜ 至少

Additional wording constraints:
- NEVER translate "primary" as "初級" or "一次測" or "一次"; always use **一次側**.
- NEVER translate "secondary" as "次級"; always use **二次側**.
- Use **Traditional Chinese** characters only.

### Table cell formatting rules
- Flammability rating cells: When you see "UL 94, UL 746C" or similar, output ONLY "UL 94" (remove UL 746C)
- Empty or blank cells: Keep them empty/blank, do NOT add any content
- Certification/approval cells with file numbers: Remove file numbers, keep ONLY the certification standard names
  Example: "VDE↓40029550↓UL E249609" → "VDE" (remove all file numbers like 40029550, E249609, E121562, etc.)

### Quality checks
Before finalizing each answer, mentally check:
1. All English technical content (except standard names, model names, etc.) has been translated into Traditional Chinese.
2. All glossary terms above are applied consistently, prioritizing the longest phrase match.
3. Numbers, units, limits, clause numbers, table structures, and verdicts are preserved exactly.
4. The overall tone is that of a professional safety/compliance report used in Taiwan.

Output ONLY the translated Traditional Chinese text (with the preserved structure), without explanations."""


# ============================================================
# 翻譯配置參數
# ============================================================
CHUNK_SIZE = 1500           # 每個 chunk 的最大字符數
MAX_RETRIES = 3             # API 呼叫最大重試次數
RETRY_DELAY = 2             # 重試間隔時間（秒）
MAX_CONCURRENT_CHUNKS = 20  # 每個文件最大併發 API 呼叫數
CACHE_TTL = 30 * 24 * 3600  # 快取過期時間：30 天（秒）
CACHE_KEY_PREFIX = "llm:translate:"  # Redis 快取 key 前綴
API_TIMEOUT = 60            # API 呼叫超時時間（秒）

# Token 價格（USD per 1M tokens）
TOKEN_PRICES = {
    'gpt-5.1': {'input': 1.25, 'cached_input': 0.125, 'output': 10.00},
    'gpt-5.2': {'input': 1.75, 'cached_input': 0.175, 'output': 14.00},
}


class LLMTranslator:
    """LLM 翻譯器 - Azure OpenAI（支援併發、Redis 快取持久化）"""

    def __init__(self, max_workers: int = MAX_CONCURRENT_CHUNKS):
        self.enabled = False
        self.client = None
        self.deployment = None
        self._memory_cache: Dict[str, str] = {}  # 記憶體快取（fallback）
        self._cache_lock = threading.Lock()
        self._redis_client = None  # Redis 快取
        self.max_workers = max_workers  # 併發數量

        # Token 統計
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.total_cached_tokens = 0
        self._stats_lock = threading.Lock()

        # 快取統計
        self.cache_hits = 0
        self.cache_misses = 0

        # 建立按長度排序的術語表（優先匹配最長詞組）
        self._sorted_glossary = sorted(
            MANDATORY_GLOSSARY.items(),
            key=lambda x: len(x[0]),
            reverse=True
        )

        # 初始化 Redis 連接
        self._init_redis()

        if not HAS_OPENAI:
            print("[LLM] openai 套件未安裝，LLM 翻譯功能停用")
            return

        # 從環境變數讀取設定
        endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "https://whaleforce-eastus2-resource.cognitiveservices.azure.com/")
        api_key = os.getenv("AZURE_OPENAI_API_KEY")
        self.deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5.1")
        api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")

        if not api_key:
            print("[LLM] 未設定 AZURE_OPENAI_API_KEY，LLM 翻譯功能停用")
            return

        try:
            self.client = AzureOpenAI(
                api_version=api_version,
                azure_endpoint=endpoint,
                api_key=api_key,
                timeout=API_TIMEOUT,  # 設定超時時間
            )
            self.enabled = True
            print(f"[LLM] Azure OpenAI 翻譯已啟用 (deployment: {self.deployment}, timeout: {API_TIMEOUT}s)")
        except Exception as e:
            print(f"[LLM] Azure OpenAI 初始化失敗: {e}")

    def _init_redis(self):
        """初始化 Redis 連接（用於快取持久化）"""
        if not HAS_REDIS:
            print("[LLM] redis 套件未安裝，使用記憶體快取")
            return

        redis_url = os.getenv("REDIS_URI") or os.getenv("REDIS_URL")
        if not redis_url:
            print("[LLM] 未設定 REDIS_URL，使用記憶體快取")
            return

        try:
            self._redis_client = redis.from_url(redis_url, decode_responses=True)
            self._redis_client.ping()  # 測試連接
            print(f"[LLM] Redis 快取已啟用")
        except Exception as e:
            print(f"[LLM] Redis 連接失敗，使用記憶體快取: {e}")
            self._redis_client = None

    def _get_cache_key(self, text: str) -> str:
        """生成快取 key（使用 hash 避免 key 過長）"""
        text_hash = hashlib.md5(text.encode('utf-8')).hexdigest()
        return f"{CACHE_KEY_PREFIX}{text_hash}"

    def _get_from_cache(self, text: str) -> Optional[str]:
        """從快取取得翻譯結果"""
        cache_key = self._get_cache_key(text)

        # 優先使用 Redis
        if self._redis_client:
            try:
                result = self._redis_client.get(cache_key)
                if result:
                    self.cache_hits += 1
                    return result
            except Exception:
                pass  # Redis 失敗時 fallback 到記憶體

        # Fallback 到記憶體快取
        with self._cache_lock:
            if text.strip() in self._memory_cache:
                self.cache_hits += 1
                return self._memory_cache[text.strip()]

        self.cache_misses += 1
        return None

    def _set_to_cache(self, text: str, translation: str):
        """將翻譯結果存入快取"""
        cache_key = self._get_cache_key(text)

        # 優先存入 Redis
        if self._redis_client:
            try:
                self._redis_client.setex(cache_key, CACHE_TTL, translation)
            except Exception:
                pass  # Redis 失敗時仍存入記憶體

        # 同時存入記憶體快取（加速本次執行）
        with self._cache_lock:
            self._memory_cache[text.strip()] = translation

    def get_cache_stats(self) -> dict:
        """取得快取統計資訊"""
        total = self.cache_hits + self.cache_misses
        hit_rate = (self.cache_hits / total * 100) if total > 0 else 0

        stats = {
            'cache_hits': self.cache_hits,
            'cache_misses': self.cache_misses,
            'hit_rate': f"{hit_rate:.1f}%",
            'redis_enabled': self._redis_client is not None,
            'memory_cache_size': len(self._memory_cache),
        }

        # 如果 Redis 可用，取得 Redis 快取數量
        if self._redis_client:
            try:
                # 使用 SCAN 計算符合 prefix 的 key 數量（避免 KEYS 命令阻塞）
                cursor = 0
                redis_count = 0
                while True:
                    cursor, keys = self._redis_client.scan(cursor, match=f"{CACHE_KEY_PREFIX}*", count=1000)
                    redis_count += len(keys)
                    if cursor == 0:
                        break
                stats['redis_cache_size'] = redis_count
            except Exception:
                stats['redis_cache_size'] = 'unknown'

        return stats

    def _is_chinese(self, text: str) -> bool:
        """檢查文本是否主要為中文"""
        if not text:
            return True
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        total_chars = len(re.sub(r'\s+', '', text))
        if total_chars == 0:
            return True
        return chinese_chars / total_chars > 0.3

    def _has_significant_english(self, text: str) -> bool:
        """
        檢查是否有需要翻譯的英文（排除常見縮寫）
        """
        if not text:
            return False

        # 移除數字和符號
        words_only = re.sub(r'[0-9\.\-\+\/%°℃Ω,;:()（）\[\]]+', ' ', text)
        # 提取英文單詞
        english_words = re.findall(r'\b[a-zA-Z]{2,}\b', words_only)

        # 過濾掉排除清單中的詞彙
        significant_words = [
            w for w in english_words
            if w.lower() not in ENGLISH_EXCLUDE_LIST
        ]

        return len(significant_words) > 0

    def _should_translate(self, text: str) -> bool:
        """判斷是否需要翻譯"""
        if not text or len(text.strip()) < 3:
            return False
        # 已經是中文為主
        if self._is_chinese(text):
            return False
        # 純數字或符號
        if re.match(r'^[\d\s\.\-\+\/%°℃Ω]+$', text):
            return False
        # 標準編號（如 IEC 60950-1）
        if re.match(r'^[A-Z]+\s*\d+[\-\d\.]*$', text.strip()):
            return False
        return True

    def _apply_special_translation(self, text: str) -> Optional[str]:
        """
        檢查特殊翻譯映射（P, N/A, -- 等）
        返回翻譯結果或 None（表示需要進一步處理）
        """
        text_stripped = text.strip()
        if text_stripped in SPECIAL_TRANSLATIONS:
            return SPECIAL_TRANSLATIONS[text_stripped]
        return None

    def _apply_glossary(self, text: str) -> str:
        """
        套用強制術語表進行預翻譯
        優先匹配最長的詞組
        """
        result = text
        for eng, chi in self._sorted_glossary:
            # 使用 word boundary 進行替換（忽略大小寫）
            pattern = re.compile(re.escape(eng), re.IGNORECASE)
            result = pattern.sub(chi, result)
        return result

    def _filter_refusal(self, text: str) -> str:
        """
        過濾 LLM 拒絕消息
        如果回覆以拒絕消息開頭，替換為 '--'
        """
        for prefix in LLM_REFUSAL_PREFIXES:
            if text.startswith(prefix):
                return '--'
        return text

    def _update_token_stats(self, usage):
        """更新 token 統計（thread-safe）"""
        with self._stats_lock:
            if hasattr(usage, 'prompt_tokens'):
                self.total_input_tokens += usage.prompt_tokens
            if hasattr(usage, 'completion_tokens'):
                self.total_output_tokens += usage.completion_tokens
            # Azure OpenAI 可能有 cached_tokens
            if hasattr(usage, 'prompt_tokens_details') and usage.prompt_tokens_details:
                if hasattr(usage.prompt_tokens_details, 'cached_tokens'):
                    self.total_cached_tokens += usage.prompt_tokens_details.cached_tokens

    def get_token_stats(self) -> Dict:
        """獲取 token 統計"""
        with self._stats_lock:
            return {
                'input_tokens': self.total_input_tokens,
                'output_tokens': self.total_output_tokens,
                'cached_tokens': self.total_cached_tokens,
                'total_tokens': self.total_input_tokens + self.total_output_tokens,
            }

    def get_cost_estimate(self) -> Dict:
        """計算成本估算（USD）"""
        stats = self.get_token_stats()
        model = self.deployment or 'gpt-5.1'
        prices = TOKEN_PRICES.get(model, TOKEN_PRICES['gpt-5.1'])

        # 計算成本（價格是 per 1M tokens）
        input_cost = (stats['input_tokens'] - stats['cached_tokens']) * prices['input'] / 1_000_000
        cached_cost = stats['cached_tokens'] * prices['cached_input'] / 1_000_000
        output_cost = stats['output_tokens'] * prices['output'] / 1_000_000
        total_cost = input_cost + cached_cost + output_cost

        # 加入快取統計
        cache_stats = self.get_cache_stats()

        return {
            'model': model,
            'input_tokens': stats['input_tokens'],
            'cached_tokens': stats['cached_tokens'],
            'output_tokens': stats['output_tokens'],
            'input_cost': round(input_cost, 4),
            'cached_cost': round(cached_cost, 4),
            'output_cost': round(output_cost, 4),
            'total_cost': round(total_cost, 4),
            'cache_hits': cache_stats['cache_hits'],
            'cache_misses': cache_stats['cache_misses'],
            'cache_hit_rate': cache_stats['hit_rate'],
            'redis_enabled': cache_stats['redis_enabled'],
        }

    def reset_stats(self):
        """重置統計"""
        with self._stats_lock:
            self.total_input_tokens = 0
            self.total_output_tokens = 0
            self.total_cached_tokens = 0
        # 重置快取統計（但不清除快取內容）
        self.cache_hits = 0
        self.cache_misses = 0

    def translate(self, text: str) -> str:
        """
        翻譯單個文本（純 LLM 模式，支援 Redis 快取持久化）

        流程:
        1. 檢查特殊翻譯映射 (P, N/A, --)
        2. 檢查快取（Redis → 記憶體）
        3. 如果 LLM 啟用且有英文，呼叫 LLM
        4. LLM 失敗或未啟用時，保留原文
        """
        if not text or len(text.strip()) < 1:
            return text

        # Step 1: 特殊翻譯映射（P, N/A, -- 等直接映射）
        special = self._apply_special_translation(text)
        if special is not None:
            return special

        # 如果已經是中文為主，直接返回
        if self._is_chinese(text):
            return text

        # Step 2: 檢查快取（Redis 優先，記憶體 fallback）
        cached = self._get_from_cache(text)
        if cached:
            return cached

        # Step 3: 純 LLM 翻譯（含重試機制）
        if self.enabled and self._has_significant_english(text):
            for retry in range(MAX_RETRIES):
                try:
                    response = self.client.chat.completions.create(
                        model=self.deployment,
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": f"翻譯以下內容：\n{text}"}
                        ],
                        max_completion_tokens=500,
                        temperature=0.1,  # 低溫度確保一致性
                    )
                    llm_result = response.choices[0].message.content.strip()

                    # 追蹤 token 使用量
                    if hasattr(response, 'usage') and response.usage:
                        self._update_token_stats(response.usage)

                    # 過濾拒絕消息
                    llm_result = self._filter_refusal(llm_result)

                    # 存入快取（Redis + 記憶體）
                    self._set_to_cache(text, llm_result)
                    return llm_result
                except Exception as e:
                    if retry < MAX_RETRIES - 1:
                        print(f"[LLM] 翻譯失敗 (重試 {retry + 1}/{MAX_RETRIES}): {e}")
                        time.sleep(RETRY_DELAY)
                    else:
                        print(f"[LLM] 翻譯失敗，保留原文: {e}")
                        # LLM 失敗時，保留原文（不使用字典）
                        return text

        # LLM 未啟用時，保留原文（不使用字典）
        return text

    def translate_no_cache(self, text: str) -> str:
        """
        翻譯單個文本（不使用/寫入快取）
        用於避免快取污染或需強制重翻的情境
        """
        if not text or len(text.strip()) < 1:
            return text

        special = self._apply_special_translation(text)
        if special is not None:
            return special

        if self._is_chinese(text):
            return text

        if self.enabled and self._has_significant_english(text):
            for retry in range(MAX_RETRIES):
                try:
                    response = self.client.chat.completions.create(
                        model=self.deployment,
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": f"翻譯以下內容：\n{text}"}
                        ],
                        max_completion_tokens=500,
                        temperature=0.1,
                    )
                    llm_result = response.choices[0].message.content.strip()
                    llm_result = self._filter_refusal(llm_result)
                    return llm_result
                except Exception as e:
                    if retry < MAX_RETRIES - 1:
                        print(f"[LLM] 翻譯失敗 (重試 {retry + 1}/{MAX_RETRIES}): {e}")
                        time.sleep(RETRY_DELAY)
                    else:
                        print(f"[LLM] 翻譯失敗，保留原文: {e}")
                        return text

        return text

    def _translate_single_for_batch(self, text: str, idx: int) -> tuple:
        """併發翻譯的單個任務"""
        try:
            result = self.translate(text)
            return (idx, result)
        except Exception as e:
            print(f"[LLM] 併發翻譯失敗 (idx={idx}): {e}")
            return (idx, text)

    def _create_chunks(self, texts: List[str], indices: List[int]) -> List[List[int]]:
        """
        將文本分組成 chunks 進行批次翻譯

        Args:
            texts: 需要翻譯的文本列表
            indices: 對應的原始索引列表

        Returns:
            chunks: 分組後的索引列表
        """
        chunks = []
        current_chunk = []
        current_length = 0

        for i, text in enumerate(texts):
            text_length = len(text)

            # 超過 CHUNK_SIZE 就新建一個 chunk
            if current_length + text_length > CHUNK_SIZE and current_chunk:
                chunks.append(current_chunk)
                current_chunk = []
                current_length = 0

            current_chunk.append(i)
            current_length += text_length

        if current_chunk:
            chunks.append(current_chunk)

        return chunks

    def _translate_chunk(self, texts: List[str], chunk_indices: List[int]) -> Dict[int, str]:
        """
        翻譯單個 chunk（合併文本用分隔符）

        Args:
            texts: 完整的文本列表
            chunk_indices: 此 chunk 包含的索引

        Returns:
            翻譯結果字典 {index: translated_text}
        """
        separator = " ||| "
        chunk_texts = [texts[i] for i in chunk_indices]
        combined_text = separator.join(chunk_texts)

        for retry in range(MAX_RETRIES):
            try:
                response = self.client.chat.completions.create(
                    model=self.deployment,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": f"翻譯以下內容（保持 ||| 分隔符）：\n{combined_text}"}
                    ],
                    max_completion_tokens=2000,
                    temperature=0.1,
                )
                translated_combined = response.choices[0].message.content.strip()

                # 追蹤 token 使用量
                if hasattr(response, 'usage') and response.usage:
                    self._update_token_stats(response.usage)

                translated_combined = self._filter_refusal(translated_combined)

                # 拆分翻譯結果
                translated_texts = translated_combined.split(separator)

                # 如果分隔符數量匹配
                if len(translated_texts) == len(chunk_indices):
                    result = {}
                    for i, idx in enumerate(chunk_indices):
                        result[idx] = translated_texts[i].strip()
                    return result

                # 分隔符不匹配，逐個翻譯
                print(f"[LLM] Chunk 分隔符不匹配，改用逐個翻譯")
                result = {}
                for idx in chunk_indices:
                    result[idx] = self.translate(texts[idx])
                return result

            except Exception as e:
                if retry < MAX_RETRIES - 1:
                    print(f"[LLM] Chunk 翻譯失敗 (重試 {retry + 1}/{MAX_RETRIES}): {e}")
                    time.sleep(RETRY_DELAY)
                else:
                    print(f"[LLM] Chunk 翻譯失敗，保留原文: {e}")
                    # 失敗時保留原文（不使用字典）
                    result = {}
                    for idx in chunk_indices:
                        result[idx] = texts[idx]
                    return result

        # 所有重試都失敗
        result = {}
        for idx in chunk_indices:
            result[idx] = texts[idx]
        return result

    def translate_batch(self, texts: List[str]) -> List[str]:
        """
        批次翻譯多個文本（分 chunk + 併發處理）

        採用二階段策略：
        1. 先將文本分成 chunks（每個 chunk ≤ CHUNK_SIZE 字符）
        2. 每個 chunk 內的文本合併用分隔符連接，一次 API 調用翻譯
        3. 使用 ThreadPoolExecutor 併發處理多個 chunks
        """
        if not self.enabled:
            # LLM 未啟用時，保留原文（不使用字典）
            return texts

        # 過濾出需要翻譯的文本（先檢查快取）
        to_translate = []
        indices = []
        results = list(texts)

        for i, text in enumerate(texts):
            if self._should_translate(text):
                # 檢查快取（Redis + 記憶體）
                cached = self._get_from_cache(text)
                if cached:
                    results[i] = cached
                    continue
                to_translate.append(text)
                indices.append(i)

        if not to_translate:
            return results

        # 分 chunk
        chunks = self._create_chunks(to_translate, indices)
        print(f"[LLM] 開始翻譯 {len(to_translate)} 個項目，分成 {len(chunks)} 個 chunks (max_workers={self.max_workers})...")

        # 使用 ThreadPoolExecutor 併發處理 chunks
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = {
                executor.submit(self._translate_chunk, to_translate, chunk): chunk_idx
                for chunk_idx, chunk in enumerate(chunks)
            }

            completed = 0
            for future in as_completed(futures):
                try:
                    chunk_results = future.result()
                    for batch_idx, translated in chunk_results.items():
                        original_idx = indices[batch_idx]
                        results[original_idx] = translated
                        # 更新快取（Redis + 記憶體）
                        self._set_to_cache(to_translate[batch_idx], translated)
                    completed += 1
                    if completed % 5 == 0 or completed == len(chunks):
                        print(f"[LLM] Chunk 進度: {completed}/{len(chunks)}")
                except Exception as e:
                    print(f"[LLM] Chunk 處理失敗: {e}")

        print(f"[LLM] 批次翻譯完成")
        return results

    def translate_with_glossary_only(self, text: str) -> str:
        """
        僅使用字典翻譯（不呼叫 LLM）
        適用於不需要 LLM 或 LLM 未啟用的情況
        """
        if not text or len(text.strip()) < 1:
            return text

        # 特殊翻譯映射
        special = self._apply_special_translation(text)
        if special is not None:
            return special

        # 套用強制術語表
        return self._apply_glossary(text)

    def final_review(self, texts: Dict[str, str]) -> Dict[str, str]:
        """
        最終審查 - 檢查並修正遺漏的英文（純 LLM 模式）
        """
        # 找出仍有英文的欄位
        to_review = {}
        for key, value in texts.items():
            if value and self._has_significant_english(value):
                to_review[key] = value

        if not to_review:
            return texts

        print(f"[翻譯] 最終審查：發現 {len(to_review)} 個含英文欄位")

        result = dict(texts)

        if self.enabled:
            # LLM 啟用時，批次翻譯
            keys = list(to_review.keys())
            values = list(to_review.values())
            translated = self.translate_batch(values)
            for i, key in enumerate(keys):
                result[key] = translated[i]
        # LLM 未啟用時，保留原文（不使用字典）

        return result

    def second_pass_translate(self, texts: List[str]) -> List[str]:
        """
        第二階段翻譯 - 掃描並翻譯殘留的英文

        用於文件翻譯後的細部檢查，找出仍含有英文的文本並重新翻譯。
        採用更嚴格的英文檢測，確保不遺漏任何需要翻譯的內容。

        Args:
            texts: 已翻譯的文本列表

        Returns:
            經過二次翻譯的文本列表
        """
        if not texts:
            return texts

        # 找出仍有殘留英文的文本
        to_retranslate = []
        indices = []
        results = list(texts)

        for i, text in enumerate(texts):
            if text and self._has_significant_english(text):
                to_retranslate.append(text)
                indices.append(i)

        if not to_retranslate:
            return results

        print(f"[二次翻譯] 發現 {len(to_retranslate)} 個殘留英文文本")

        if self.enabled:
            # LLM 啟用時，批次翻譯
            translated = self.translate_batch(to_retranslate)
            for i, idx in enumerate(indices):
                results[idx] = translated[i]
            print(f"[二次翻譯] LLM 完成 {len(to_retranslate)} 個文本")
        else:
            # LLM 未啟用時，保留原文（不使用字典）
            print(f"[二次翻譯] LLM 未啟用，保留原文")

        return results


# 全局翻譯器實例
_translator: Optional[LLMTranslator] = None


def get_translator() -> LLMTranslator:
    """獲取全局翻譯器"""
    global _translator
    if _translator is None:
        _translator = LLMTranslator()
    return _translator


def llm_translate(text: str) -> str:
    """便捷函數：翻譯單個文本（使用字典 + LLM）"""
    return get_translator().translate(text)


def glossary_translate(text: str) -> str:
    """便捷函數：僅使用字典翻譯（不呼叫 LLM）"""
    return get_translator().translate_with_glossary_only(text)


def llm_translate_batch(texts: List[str]) -> List[str]:
    """便捷函數：批次翻譯"""
    return get_translator().translate_batch(texts)


def llm_final_review(texts: Dict[str, str]) -> Dict[str, str]:
    """便捷函數：最終審查"""
    return get_translator().final_review(texts)


def llm_second_pass(texts: List[str]) -> List[str]:
    """便捷函數：第二階段翻譯（掃描殘留英文）"""
    return get_translator().second_pass_translate(texts)


def get_token_stats() -> Dict:
    """便捷函數：獲取 token 統計"""
    return get_translator().get_token_stats()


def get_cost_estimate() -> Dict:
    """便捷函數：獲取成本估算"""
    return get_translator().get_cost_estimate()


def reset_translator_stats():
    """便捷函數：重置翻譯器統計"""
    get_translator().reset_stats()


# 導出術語表供其他模組使用
def get_mandatory_glossary() -> Dict[str, str]:
    """獲取強制術語表"""
    return MANDATORY_GLOSSARY.copy()


def get_special_translations() -> Dict[str, str]:
    """獲取特殊翻譯映射"""
    return SPECIAL_TRANSLATIONS.copy()
