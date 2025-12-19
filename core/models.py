# core/models.py
"""
Job 狀態資料結構
"""
from enum import Enum
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Optional, List, Dict, Any
import json


class JobStatus(str, Enum):
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    PASS = "PASS"
    FAIL = "FAIL"
    ERROR = "ERROR"


@dataclass
class QAResult:
    """QA 檢查結果"""
    gate_name: str
    status: str  # PASS, FAIL, WARN
    message: str = ""
    details: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class Job:
    """工作狀態"""
    job_id: str
    status: JobStatus = JobStatus.PENDING
    created_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())

    # 輸入
    original_pdf_key: str = ""
    pdf_filename: str = ""

    # 處理中間產物
    json_data_key: str = ""
    pdf_clause_rows_key: str = ""

    # QA 結果
    qa_results: List[QAResult] = field(default_factory=list)
    qa_report_key: str = ""

    # 輸出
    docx_key: str = ""  # FINAL 或 DRAFT
    docx_type: str = ""  # "FINAL" 或 "DRAFT"

    # 錯誤訊息
    error_message: str = ""

    # 封面欄位（用戶填入）
    cover_report_no: str = ""
    cover_applicant_name: str = ""
    cover_applicant_address: str = ""

    def update_status(self, status: JobStatus):
        self.status = status
        self.updated_at = datetime.utcnow().isoformat()

    def add_qa_result(self, gate_name: str, status: str, message: str = "", details: Dict = None):
        self.qa_results.append(QAResult(
            gate_name=gate_name,
            status=status,
            message=message,
            details=details or {}
        ))
        self.updated_at = datetime.utcnow().isoformat()

    def to_dict(self) -> dict:
        return {
            "job_id": self.job_id,
            "status": self.status.value,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "original_pdf_key": self.original_pdf_key,
            "pdf_filename": self.pdf_filename,
            "json_data_key": self.json_data_key,
            "pdf_clause_rows_key": self.pdf_clause_rows_key,
            "qa_results": [r.to_dict() for r in self.qa_results],
            "qa_report_key": self.qa_report_key,
            "docx_key": self.docx_key,
            "docx_type": self.docx_type,
            "error_message": self.error_message,
            "cover_report_no": self.cover_report_no,
            "cover_applicant_name": self.cover_applicant_name,
            "cover_applicant_address": self.cover_applicant_address,
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    @classmethod
    def from_dict(cls, data: dict) -> "Job":
        job = cls(
            job_id=data["job_id"],
            status=JobStatus(data.get("status", "PENDING")),
            created_at=data.get("created_at", ""),
            updated_at=data.get("updated_at", ""),
            original_pdf_key=data.get("original_pdf_key", ""),
            pdf_filename=data.get("pdf_filename", ""),
            json_data_key=data.get("json_data_key", ""),
            pdf_clause_rows_key=data.get("pdf_clause_rows_key", ""),
            qa_report_key=data.get("qa_report_key", ""),
            docx_key=data.get("docx_key", ""),
            docx_type=data.get("docx_type", ""),
            error_message=data.get("error_message", ""),
            cover_report_no=data.get("cover_report_no", ""),
            cover_applicant_name=data.get("cover_applicant_name", ""),
            cover_applicant_address=data.get("cover_applicant_address", ""),
        )
        for qa in data.get("qa_results", []):
            job.qa_results.append(QAResult(**qa))
        return job
