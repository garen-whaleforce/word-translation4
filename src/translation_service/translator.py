"""
Translation Service - 使用 LiteLLM 進行英中翻譯

功能:
1. Bulk pass: 快速翻譯 (gemini-2.5-flash)
2. Refinement pass: 精修低品質片段 (gpt-5.2)
3. 強制遵守 termbase (placeholder 機制)
4. QA 報告生成
"""
from __future__ import annotations
import re
import logging
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from pathlib import Path

# Lazy import litellm
litellm = None

from ..config import settings
from ..termbase import Termbase, TermEntry, TermProtection, load_termbase_from_json

logger = logging.getLogger(__name__)


@dataclass
class TranslationResult:
    """翻譯結果"""
    original_text: str
    translated_text: str
    model_used: str
    was_refined: bool = False
    refine_model: Optional[str] = None
    term_mapping: Dict[str, TermEntry] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    # Token 和成本統計
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    cost: float = 0.0  # USD


@dataclass
class TranslationQAReport:
    """翻譯品質報告"""
    total_segments: int = 0
    translated_segments: int = 0
    refined_segments: int = 0
    term_violations: List[Dict[str, Any]] = field(default_factory=list)
    number_mismatches: List[Dict[str, Any]] = field(default_factory=list)
    quality_score: float = 1.0  # 0-1, 1 為最佳
    # Token 和成本統計
    total_prompt_tokens: int = 0
    total_completion_tokens: int = 0
    total_tokens: int = 0
    total_cost: float = 0.0  # USD

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_segments": self.total_segments,
            "translated_segments": self.translated_segments,
            "refined_segments": self.refined_segments,
            "term_violations": self.term_violations,
            "number_mismatches": self.number_mismatches,
            "quality_score": self.quality_score,
            "total_prompt_tokens": self.total_prompt_tokens,
            "total_completion_tokens": self.total_completion_tokens,
            "total_tokens": self.total_tokens,
            "total_cost": self.total_cost
        }


class TranslationService:
    """翻譯服務"""

    # 翻譯提示詞
    TRANSLATION_PROMPT = """你是專業的 IEC 安規標準翻譯專家。請將以下英文翻譯成繁體中文。

重要規則:
1. 保留所有 ⟦TERM_xxxx⟧ 格式的標記，不要翻譯或修改它們
2. 保留所有數字、條款編號（如 4.1.1, G.7.3.2）、測試值（如 250V, 60Hz）
3. 保留所有單位（如 mm, kV, mA, °C）
4. 翻譯要專業、準確、符合台灣安規術語習慣
5. P = 通過, N/A = 不適用, Fail = 不通過

英文原文:
{text}

繁體中文翻譯:"""

    REFINEMENT_PROMPT = """你是專業的 IEC 安規標準翻譯校對專家。請修正以下翻譯中的問題。

原始英文:
{original}

目前翻譯:
{translated}

發現的問題:
{issues}

請只修正有問題的部分，保持其他內容不變。輸出修正後的完整翻譯:"""

    def __init__(
        self,
        termbase: Optional[Termbase] = None,
        bulk_model: str = None,
        refine_model: str = None,
        api_base: str = None,
        api_key: str = None,
        dry_run: bool = False
    ):
        """
        初始化翻譯服務

        Args:
            termbase: 術語庫
            bulk_model: Bulk 翻譯使用的模型
            refine_model: Refinement 使用的模型
            api_base: LiteLLM API base URL
            api_key: API key
            dry_run: 是否為乾跑模式 (不呼叫 LLM)
        """
        self.termbase = termbase or Termbase()
        self.bulk_model = bulk_model or settings.bulk_model
        self.refine_model = refine_model or settings.refine_model
        self.api_base = api_base or settings.litellm_api_base
        self.api_key = api_key or settings.litellm_api_key
        self.dry_run = dry_run
        self._litellm = None

    def _get_litellm(self):
        """Lazy load litellm module."""
        if self._litellm is None:
            try:
                import litellm as _litellm
                if self.api_base:
                    _litellm.api_base = self.api_base
                if self.api_key:
                    _litellm.api_key = self.api_key
                self._litellm = _litellm
            except ImportError:
                logger.warning("litellm not installed, using dry_run mode")
                self.dry_run = True
        return self._litellm

    def translate(
        self,
        text: str,
        enable_refinement: bool = True,
        model: str = None
    ) -> TranslationResult:
        """
        翻譯文字

        Args:
            text: 要翻譯的英文文字
            enable_refinement: 是否啟用精修
            model: 指定使用的模型 (覆蓋預設)

        Returns:
            TranslationResult 包含翻譯結果和 metadata
        """
        if not text or not text.strip():
            return TranslationResult(
                original_text=text,
                translated_text=text,
                model_used="none"
            )

        # 累計 token 和成本
        total_prompt_tokens = 0
        total_completion_tokens = 0
        total_cost = 0.0

        # 1. 保護術語
        protection = self.termbase.protect_terms(text)
        logger.debug(f"Protected text: {protection.protected_text}")

        # 2. Bulk 翻譯
        bulk_model = model or self.bulk_model
        translated, prompt_tokens, completion_tokens, cost = self._call_llm(
            protection.protected_text,
            bulk_model
        )
        total_prompt_tokens += prompt_tokens
        total_completion_tokens += completion_tokens
        total_cost += cost

        # 3. 還原術語
        translated = self.termbase.restore_terms(translated, protection.mapping)

        # 4. QA 檢查
        issues = self._check_quality(text, translated)

        # 5. Refinement (如果需要)
        was_refined = False
        refine_model_used = None

        if enable_refinement and issues:
            refined, ref_prompt, ref_completion, ref_cost = self._refine_translation(text, translated, issues)
            if refined != translated:
                translated = refined
                was_refined = True
                refine_model_used = self.refine_model
                total_prompt_tokens += ref_prompt
                total_completion_tokens += ref_completion
                total_cost += ref_cost

        return TranslationResult(
            original_text=text,
            translated_text=translated,
            model_used=bulk_model,
            was_refined=was_refined,
            refine_model=refine_model_used,
            term_mapping=protection.mapping,
            metadata={
                "issues_found": len(issues),
                "protected_terms": len(protection.mapping)
            },
            prompt_tokens=total_prompt_tokens,
            completion_tokens=total_completion_tokens,
            total_tokens=total_prompt_tokens + total_completion_tokens,
            cost=total_cost
        )

    def translate_batch(
        self,
        texts: List[str],
        enable_refinement: bool = True
    ) -> List[TranslationResult]:
        """
        批量翻譯

        Args:
            texts: 要翻譯的文字列表
            enable_refinement: 是否啟用精修

        Returns:
            翻譯結果列表
        """
        results = []
        for text in texts:
            result = self.translate(text, enable_refinement)
            results.append(result)
        return results

    def generate_qa_report(self, results: List[TranslationResult]) -> TranslationQAReport:
        """
        生成 QA 報告

        Args:
            results: 翻譯結果列表

        Returns:
            TranslationQAReport
        """
        report = TranslationQAReport(
            total_segments=len(results),
            translated_segments=len([r for r in results if r.translated_text]),
            refined_segments=len([r for r in results if r.was_refined])
        )

        # 檢查所有結果並累計 token/cost
        for i, result in enumerate(results):
            # 累計 token 統計
            report.total_prompt_tokens += result.prompt_tokens
            report.total_completion_tokens += result.completion_tokens
            report.total_tokens += result.total_tokens
            report.total_cost += result.cost

            # 術語違規
            violations = self.termbase.validate_terms(result.translated_text)
            for v in violations:
                report.term_violations.append({
                    "segment_index": i,
                    "term": v.term,
                    "type": v.violation_type
                })

            # 數字/條款不一致
            mismatches = self._check_numbers(
                result.original_text,
                result.translated_text
            )
            for m in mismatches:
                report.number_mismatches.append({
                    "segment_index": i,
                    **m
                })

        # 計算品質分數
        total_issues = len(report.term_violations) + len(report.number_mismatches)
        if report.total_segments > 0:
            report.quality_score = max(0, 1 - (total_issues / report.total_segments * 0.1))

        return report

    def _call_llm(self, text: str, model: str) -> tuple:
        """
        呼叫 LLM 進行翻譯

        Returns:
            tuple: (translated_text, prompt_tokens, completion_tokens, cost)
        """
        if self.dry_run:
            # 乾跑模式：返回模擬翻譯
            return f"[DRY_RUN:{model}] {text}", 0, 0, 0.0

        try:
            # 使用 OpenAI SDK (相容 LiteLLM 代理)
            from openai import OpenAI

            client = OpenAI(
                api_key=self.api_key,
                base_url=self.api_base
            )

            prompt = self.TRANSLATION_PROMPT.format(text=text)

            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=4096,
                temperature=0.3
            )

            # 提取 token 統計
            usage = response.usage
            prompt_tokens = usage.prompt_tokens if usage else 0
            completion_tokens = usage.completion_tokens if usage else 0

            # 從 LiteLLM 回應取得成本 (如果有)
            cost = getattr(usage, 'cost', 0.0) if usage else 0.0
            if cost == 0.0:
                # 估算成本 (gemini-2.5-flash 價格: $0.075/1M input, $0.30/1M output)
                cost = (prompt_tokens * 0.075 / 1_000_000) + (completion_tokens * 0.30 / 1_000_000)

            return response.choices[0].message.content.strip(), prompt_tokens, completion_tokens, cost

        except Exception as e:
            logger.error(f"LLM call failed: {e}")
            # 失敗時返回原文
            return text, 0, 0, 0.0

    def _check_quality(self, original: str, translated: str) -> List[Dict[str, Any]]:
        """檢查翻譯品質"""
        issues = []

        # 1. 檢查未還原的 placeholder
        for match in Termbase.PLACEHOLDER_PATTERN.finditer(translated):
            issues.append({
                "type": "unrestored_placeholder",
                "value": match.group()
            })

        # 2. 檢查數字/條款是否保留
        number_issues = self._check_numbers(original, translated)
        issues.extend(number_issues)

        return issues

    def _check_numbers(self, original: str, translated: str) -> List[Dict[str, Any]]:
        """檢查數字和條款編號是否保留"""
        issues = []

        # 條款編號 (如 4.1.1, G.7.3.2.1)
        clause_pattern = re.compile(r'\b([A-Z]\.)?(\d+)(\.\d+)+\b')
        original_clauses = set(clause_pattern.findall(original))
        translated_clauses = set(clause_pattern.findall(translated))

        # 重要數字 (如 250V, 60Hz, 25°C)
        number_pattern = re.compile(r'\b(\d+(?:\.\d+)?)\s*(V|A|W|Hz|°C|mm|kV|mA|kW)\b')
        original_numbers = set(number_pattern.findall(original))
        translated_numbers = set(number_pattern.findall(translated))

        # 檢查缺失
        for clause in original_clauses - translated_clauses:
            issues.append({
                "type": "missing_clause",
                "value": "".join(clause)
            })

        for num in original_numbers - translated_numbers:
            issues.append({
                "type": "missing_number",
                "value": f"{num[0]}{num[1]}"
            })

        return issues

    def _refine_translation(
        self,
        original: str,
        translated: str,
        issues: List[Dict[str, Any]]
    ) -> tuple:
        """
        精修翻譯

        Returns:
            tuple: (refined_text, prompt_tokens, completion_tokens, cost)
        """
        if self.dry_run:
            return f"[REFINED:{self.refine_model}] {translated}", 0, 0, 0.0

        try:
            from openai import OpenAI

            client = OpenAI(
                api_key=self.api_key,
                base_url=self.api_base
            )

            issues_text = "\n".join([
                f"- {issue['type']}: {issue.get('value', '')}"
                for issue in issues
            ])

            prompt = self.REFINEMENT_PROMPT.format(
                original=original,
                translated=translated,
                issues=issues_text
            )

            response = client.chat.completions.create(
                model=self.refine_model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=4096,
                temperature=0.1
            )

            # 提取 token 統計
            usage = response.usage
            prompt_tokens = usage.prompt_tokens if usage else 0
            completion_tokens = usage.completion_tokens if usage else 0

            # 從 LiteLLM 回應取得成本
            cost = getattr(usage, 'cost', 0.0) if usage else 0.0
            if cost == 0.0:
                cost = (prompt_tokens * 0.075 / 1_000_000) + (completion_tokens * 0.30 / 1_000_000)

            return response.choices[0].message.content.strip(), prompt_tokens, completion_tokens, cost

        except Exception as e:
            logger.error(f"Refinement failed: {e}")
            return translated, 0, 0, 0.0


def translate_text(
    text: str,
    termbase: Optional[Termbase] = None,
    dry_run: bool = False
) -> TranslationResult:
    """
    便捷函數：翻譯單一文字

    Args:
        text: 要翻譯的文字
        termbase: 術語庫 (可選)
        dry_run: 是否為乾跑模式

    Returns:
        TranslationResult
    """
    service = TranslationService(termbase=termbase, dry_run=dry_run)
    return service.translate(text)
