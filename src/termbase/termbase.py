"""
Termbase Module - 術語庫與 placeholder 強制機制

功能:
1. 載入術語表 (JSON/CSV)
2. protect_terms: 把英文術語替換為 placeholder (⟦TERM_xxxx⟧)
3. restore_terms: 把 placeholder 還原為中文術語
4. validate_terms: 檢查翻譯結果是否違反術語規則
"""
from __future__ import annotations
import re
import csv
import json
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class TermEntry:
    """術語條目"""
    source_en: str  # 英文術語
    target_zh: str  # 中文翻譯
    priority: int = 0  # 優先順序 (越高越優先)
    note: str = ""  # 備註
    count: int = 0  # 出現次數 (用於統計)

    @property
    def length(self) -> int:
        """術語長度 (用於最長匹配)"""
        return len(self.source_en)


@dataclass
class TermProtection:
    """術語保護結果"""
    protected_text: str  # 替換後的文字
    mapping: Dict[str, TermEntry]  # placeholder -> TermEntry 的映射
    original_text: str  # 原始文字


@dataclass
class TermViolation:
    """術語違規"""
    term: str  # 違規的術語
    expected: str  # 預期的翻譯
    found: str  # 發現的翻譯
    position: int  # 在文字中的位置
    violation_type: str  # 違規類型: unprotected_token, wrong_translation


class Termbase:
    """術語庫管理器"""

    # Placeholder 格式
    PLACEHOLDER_PATTERN = re.compile(r'⟦TERM_([0-9a-f]{4})⟧')
    PLACEHOLDER_TEMPLATE = "⟦TERM_{:04x}⟧"

    def __init__(self):
        self.entries: Dict[str, TermEntry] = {}  # source_en (lowercase) -> TermEntry
        self._sorted_entries: List[TermEntry] = []  # 按長度和優先順序排序
        self._counter = 0

    def add_entry(self, entry: TermEntry):
        """新增術語條目"""
        key = entry.source_en.lower().strip()
        if key in self.entries:
            # 更新現有條目 (如果新的優先順序更高)
            if entry.priority > self.entries[key].priority:
                self.entries[key] = entry
        else:
            self.entries[key] = entry

        # 重新排序
        self._sort_entries()

    def _sort_entries(self):
        """按長度降序、優先順序降序排序 (最長匹配優先)"""
        self._sorted_entries = sorted(
            self.entries.values(),
            key=lambda e: (-e.length, -e.priority)
        )

    def protect_terms(self, text: str) -> TermProtection:
        """
        保護術語：用 placeholder 替換英文術語

        Args:
            text: 要處理的英文文字

        Returns:
            TermProtection 包含替換後文字和映射表
        """
        if not text:
            return TermProtection(
                protected_text=text,
                mapping={},
                original_text=text
            )

        protected_text = text
        mapping: Dict[str, TermEntry] = {}
        used_positions: Set[Tuple[int, int]] = set()

        # 按照排序順序 (最長優先) 進行匹配
        for entry in self._sorted_entries:
            # 建立不區分大小寫的搜索模式
            pattern = re.compile(
                re.escape(entry.source_en),
                re.IGNORECASE
            )

            # 找出所有匹配位置
            for match in pattern.finditer(protected_text):
                start, end = match.start(), match.end()

                # 檢查是否已被其他術語佔用
                overlap = False
                for used_start, used_end in used_positions:
                    if not (end <= used_start or start >= used_end):
                        overlap = True
                        break

                if overlap:
                    continue

                # 檢查匹配的文字是否已是 placeholder
                matched_text = match.group()
                if matched_text.startswith("⟦TERM_"):
                    continue

                # 生成 placeholder
                placeholder = self.PLACEHOLDER_TEMPLATE.format(self._counter)
                self._counter = (self._counter + 1) % 0xFFFF

                # 記錄映射
                mapping[placeholder] = entry
                used_positions.add((start, end))

        # 批量替換 (從後往前，避免位置偏移)
        replacements = []
        for entry in self._sorted_entries:
            pattern = re.compile(re.escape(entry.source_en), re.IGNORECASE)
            for match in pattern.finditer(text):
                matched_text = match.group()
                if matched_text.startswith("⟦TERM_"):
                    continue

                # 找到對應的 placeholder
                for placeholder, term_entry in mapping.items():
                    if term_entry.source_en.lower() == entry.source_en.lower():
                        replacements.append((match.start(), match.end(), placeholder))
                        break

        # 去重並排序
        replacements = list(set(replacements))
        replacements.sort(key=lambda x: -x[0])  # 從後往前

        # 執行替換
        protected_text = text
        for start, end, placeholder in replacements:
            protected_text = protected_text[:start] + placeholder + protected_text[end:]

        return TermProtection(
            protected_text=protected_text,
            mapping=mapping,
            original_text=text
        )

    def restore_terms(self, text: str, mapping: Dict[str, TermEntry]) -> str:
        """
        還原術語：把 placeholder 替換為中文術語

        Args:
            text: 含有 placeholder 的翻譯文字
            mapping: placeholder -> TermEntry 的映射

        Returns:
            還原後的中文文字
        """
        if not text or not mapping:
            return text

        result = text
        for placeholder, entry in mapping.items():
            result = result.replace(placeholder, entry.target_zh)

        return result

    def validate_terms(self, text: str) -> List[TermViolation]:
        """
        驗證術語：檢查翻譯結果是否有術語違規

        Args:
            text: 翻譯後的文字

        Returns:
            違規列表
        """
        violations = []

        # 1. 檢查未替換的 placeholder
        for match in self.PLACEHOLDER_PATTERN.finditer(text):
            violations.append(TermViolation(
                term=match.group(),
                expected="應被替換為中文術語",
                found=match.group(),
                position=match.start(),
                violation_type="unprotected_token"
            ))

        # 2. 檢查是否有英文術語殘留 (這些本應被翻譯)
        for entry in self._sorted_entries:
            pattern = re.compile(
                r'\b' + re.escape(entry.source_en) + r'\b',
                re.IGNORECASE
            )
            for match in pattern.finditer(text):
                # 排除在 placeholder 內的情況
                if not self.PLACEHOLDER_PATTERN.match(text[max(0, match.start()-10):match.end()+10]):
                    violations.append(TermViolation(
                        term=entry.source_en,
                        expected=entry.target_zh,
                        found=match.group(),
                        position=match.start(),
                        violation_type="untranslated_term"
                    ))

        return violations

    def get_term(self, source_en: str) -> Optional[TermEntry]:
        """取得術語條目"""
        return self.entries.get(source_en.lower().strip())

    def __len__(self) -> int:
        return len(self.entries)

    def __contains__(self, term: str) -> bool:
        return term.lower().strip() in self.entries


def load_termbase_from_json(json_path: Path) -> Termbase:
    """
    從 JSON 檔案載入術語庫

    預期格式:
    [
        {"en_norm": "...", "zh_pref": "...", "count": N, ...},
        ...
    ]
    """
    termbase = Termbase()

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for item in data:
        entry = TermEntry(
            source_en=item.get("en_norm", ""),
            target_zh=item.get("zh_pref", ""),
            priority=item.get("count", 0),  # 使用出現次數作為優先順序
            count=item.get("count", 0)
        )
        if entry.source_en and entry.target_zh:
            termbase.add_entry(entry)

    logger.info(f"Loaded {len(termbase)} terms from {json_path}")
    return termbase


def load_termbase_from_csv(csv_path: Path) -> Termbase:
    """
    從 CSV 檔案載入術語庫 (翻譯記憶庫格式)

    預期格式:
    source, context, en_raw, zh_raw, en_norm, zh_norm
    """
    termbase = Termbase()

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            en_text = row.get("en_norm") or row.get("en_raw", "")
            zh_text = row.get("zh_norm") or row.get("zh_raw", "")

            if en_text and zh_text:
                # 檢查是否已存在
                existing = termbase.get_term(en_text)
                priority = existing.priority + 1 if existing else 1

                entry = TermEntry(
                    source_en=en_text.strip(),
                    target_zh=zh_text.strip(),
                    priority=priority,
                    count=1
                )
                termbase.add_entry(entry)

    logger.info(f"Loaded {len(termbase)} terms from {csv_path}")
    return termbase


def create_combined_termbase(
    glossary_path: Optional[Path] = None,
    tm_path: Optional[Path] = None
) -> Termbase:
    """
    建立合併的術語庫 (優先使用 glossary，再補充 TM)
    """
    termbase = Termbase()

    # 1. 載入 glossary (高優先順序)
    if glossary_path and glossary_path.exists():
        glossary_tb = load_termbase_from_json(glossary_path)
        for entry in glossary_tb._sorted_entries:
            entry.priority += 1000  # 提高 glossary 優先順序
            termbase.add_entry(entry)

    # 2. 載入 TM (低優先順序，填補空白)
    if tm_path and tm_path.exists():
        tm_tb = load_termbase_from_csv(tm_path)
        for entry in tm_tb._sorted_entries:
            if entry.source_en.lower() not in termbase.entries:
                termbase.add_entry(entry)

    logger.info(f"Combined termbase has {len(termbase)} terms")
    return termbase
