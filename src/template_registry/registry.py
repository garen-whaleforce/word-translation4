"""
Template Registry Module - 模板管理與自動選擇

功能:
1. 註冊/管理多個 Word 模板
2. 根據 PDF 內容自動選擇最適合的模板
3. 支援 signature 匹配 (TRF No, Clause 集合, 錨點)
"""
from __future__ import annotations
import json
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Set
from dataclasses import dataclass, field

from ..cb_parser import ParseResult

logger = logging.getLogger(__name__)


@dataclass
class TemplateInfo:
    """模板資訊"""
    id: str  # 唯一識別符
    name: str  # 顯示名稱
    path: Path  # 檔案路徑
    description: str = ""

    # Signature for matching
    trf_patterns: List[str] = field(default_factory=list)  # e.g., ["IEC62368"]
    expected_clauses: Set[str] = field(default_factory=set)  # e.g., {"4", "4.1.1", ...}
    anchors: List[str] = field(default_factory=list)  # e.g., ["安全防護總攬表"]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.id,
            "name": self.name,
            "path": str(self.path),
            "description": self.description,
            "trf_patterns": self.trf_patterns,
            "expected_clauses": list(self.expected_clauses),
            "anchors": self.anchors
        }


class TemplateRegistry:
    """模板註冊表"""

    def __init__(self, templates_dir: Optional[Path] = None):
        """
        初始化註冊表

        Args:
            templates_dir: 模板目錄
        """
        self.templates_dir = templates_dir or Path("templates")
        self.templates: Dict[str, TemplateInfo] = {}
        self._load_templates()

    def _load_templates(self):
        """掃描並載入模板"""
        if not self.templates_dir.exists():
            logger.warning(f"Templates directory not found: {self.templates_dir}")
            return

        # 載入配置檔 (如果存在)
        config_path = self.templates_dir / "templates.json"
        if config_path.exists():
            self._load_from_config(config_path)

        # 掃描 docx 檔案
        for docx_file in self.templates_dir.glob("*.docx"):
            if docx_file.stem not in self.templates:
                self._register_from_file(docx_file)

    def _load_from_config(self, config_path: Path):
        """從配置檔載入模板資訊"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            self.default_template_id = config.get("default_template", None)

            for item in config.get("templates", []):
                # 支援 filename 或 path
                template_file = item.get("filename") or item.get("path")
                template = TemplateInfo(
                    id=item["id"],
                    name=item.get("name", item["id"]),
                    path=self.templates_dir / template_file,
                    description=item.get("description", ""),
                    trf_patterns=item.get("trf_patterns", []),
                    expected_clauses=set(item.get("expected_clauses", [])),
                    anchors=item.get("anchors", [])
                )
                # 儲存 signature 資訊
                if "signature" in item:
                    template.signature = item["signature"]

                if template.path.exists():
                    self.templates[template.id] = template
                    logger.info(f"Loaded template: {template.id} ({template.path.name})")
                else:
                    logger.warning(f"Template file not found: {template.path}")

        except Exception as e:
            logger.error(f"Failed to load templates config: {e}")

    def _register_from_file(self, docx_path: Path):
        """從檔案註冊模板 (最小資訊)"""
        template = TemplateInfo(
            id=docx_path.stem,
            name=docx_path.stem.replace("_", " ").replace("-", " "),
            path=docx_path,
            description=f"Auto-detected template: {docx_path.name}"
        )
        self.templates[template.id] = template
        logger.debug(f"Auto-registered template: {template.id}")

    def register(self, template: TemplateInfo):
        """手動註冊模板"""
        self.templates[template.id] = template

    def get(self, template_id: str) -> Optional[TemplateInfo]:
        """取得模板"""
        return self.templates.get(template_id)

    def list_templates(self) -> List[TemplateInfo]:
        """列出所有模板"""
        return list(self.templates.values())

    def select_best_template(
        self,
        parse_result: ParseResult,
        threshold: float = 0.5
    ) -> Optional[TemplateInfo]:
        """
        根據 PDF 解析結果選擇最適合的模板

        Args:
            parse_result: PDF 解析結果
            threshold: 最低匹配分數 (0-1)

        Returns:
            最適合的模板，或 None 如果沒有達到門檻
        """
        if not self.templates:
            return None

        pdf_trf = parse_result.trf_no.upper() if parse_result.trf_no else ""
        pdf_clauses = {c.clause_id for c in parse_result.clauses}

        best_template = None
        best_score = 0

        for template in self.templates.values():
            score = self._calculate_match_score(
                template,
                pdf_trf,
                pdf_clauses
            )

            if score > best_score:
                best_score = score
                best_template = template

        if best_score >= threshold:
            logger.info(f"Selected template: {best_template.id} (score: {best_score:.2f})")
            return best_template
        else:
            logger.warning(f"No template matched above threshold. Best: {best_template.id if best_template else 'None'} ({best_score:.2f})")
            return None

    def _calculate_match_score(
        self,
        template: TemplateInfo,
        pdf_trf: str,
        pdf_clauses: Set[str]
    ) -> float:
        """計算模板匹配分數"""
        score = 0.0
        weights = {"trf": 0.4, "clauses": 0.6}

        # TRF 匹配
        if template.trf_patterns:
            for pattern in template.trf_patterns:
                if pattern.upper() in pdf_trf:
                    score += weights["trf"]
                    break
        else:
            # 無 TRF 限制，給一半分數
            score += weights["trf"] * 0.5

        # Clause 匹配 (Jaccard similarity)
        if template.expected_clauses and pdf_clauses:
            intersection = len(template.expected_clauses & pdf_clauses)
            union = len(template.expected_clauses | pdf_clauses)
            jaccard = intersection / union if union > 0 else 0
            score += weights["clauses"] * jaccard
        else:
            # 無 clause 限制，給一半分數
            score += weights["clauses"] * 0.5

        return score


def select_template(
    parse_result: ParseResult,
    templates_dir: Optional[Path] = None
) -> Optional[TemplateInfo]:
    """
    便捷函數：選擇最適合的模板

    Args:
        parse_result: PDF 解析結果
        templates_dir: 模板目錄

    Returns:
        最適合的模板
    """
    registry = TemplateRegistry(templates_dir)
    return registry.select_best_template(parse_result)
