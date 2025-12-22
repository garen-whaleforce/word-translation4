# apps/api/main.py
"""
FastAPI 應用程式
提供 CB PDF 上傳和 CNS DOCX 下載 API
"""
import os
import uuid
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query, Depends
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.security import APIKeyQuery
from pydantic import BaseModel
import redis

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from core.models import Job, JobStatus
from core.storage import get_storage

# ===== 配置 =====
# 支援 Zeabur 自動注入的環境變數
REDIS_URL = os.getenv("REDIS_URI") or os.getenv("REDIS_URL", "redis://localhost:6379")
ALLOWED_ORIGINS = os.getenv("CORS_ORIGINS", "*").split(",")
SHARED_PASSWORD = os.getenv("SHARED_PASSWORD", "")  # 空字串表示不啟用認證

# ===== 認證 =====
api_key_query = APIKeyQuery(name="p", auto_error=False)


async def verify_password(api_key: str = Depends(api_key_query)) -> str:
    """
    驗證 API 密碼
    - 若 SHARED_PASSWORD 未設定或為空，則跳過認證
    - 若已設定，則必須提供正確的密碼
    """
    if not SHARED_PASSWORD:
        # 未設定密碼，跳過認證
        return ""

    if not api_key or api_key != SHARED_PASSWORD:
        raise HTTPException(
            status_code=403,
            detail="Invalid or missing password. Use ?p=<password> query parameter."
        )
    return api_key

# ===== App =====
app = FastAPI(
    title="CNS 15598-1 Report Generator",
    description="CB PDF → CNS DOCX 轉換服務",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== Redis =====
redis_client: Optional[redis.Redis] = None


def get_redis() -> redis.Redis:
    global redis_client
    if redis_client is None:
        redis_client = redis.from_url(REDIS_URL, decode_responses=True)
    return redis_client


def get_job_or_404(job_id: str) -> Job:
    """
    從 Redis 取得 Job，若不存在則拋出 404
    """
    r = get_redis()
    job_data = r.get(f"job:{job_id}")
    if not job_data:
        raise HTTPException(status_code=404, detail="Job not found")
    return Job.from_dict(json.loads(job_data))


def safe_filename(filename: str) -> str:
    """
    過濾檔案名稱，防止路徑遍歷攻擊
    """
    # 只取最後的檔名部分，移除任何路徑
    return Path(filename).name


# ===== Models =====
class JobResponse(BaseModel):
    job_id: str
    status: str
    created_at: str
    updated_at: str
    pdf_filename: str
    docx_type: Optional[str] = None
    error_message: Optional[str] = None


class JobDetailResponse(JobResponse):
    qa_results: list = []
    docx_download_url: Optional[str] = None
    qa_report_download_url: Optional[str] = None


# ===== Routes =====
@app.get("/api")
async def api_root():
    return {"status": "ok", "service": "CNS Report Generator API"}


@app.get("/health")
async def health():
    try:
        r = get_redis()
        r.ping()
        redis_status = "connected"
    except Exception as e:
        redis_status = f"error: {str(e)}"

    storage = get_storage()
    storage_status = "enabled" if storage.enabled else "local"

    return {
        "status": "healthy",
        "redis": redis_status,
        "storage": storage_status,
        "timestamp": datetime.utcnow().isoformat()
    }


MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB


@app.post("/api/jobs", response_model=JobResponse)
async def create_job(
    file: UploadFile = File(...),
    report_no: Optional[str] = Form(None),
    applicant_name: Optional[str] = Form(None),
    applicant_address: Optional[str] = Form(None),
    _: str = Depends(verify_password),
):
    """
    上傳 CB PDF 並建立轉換任務

    封面欄位（選填）：
    - report_no: 報告編號
    - applicant_name: 申請者名稱
    - applicant_address: 申請者地址

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    # 過濾檔案名稱，防止路徑遍歷
    original_filename = safe_filename(file.filename or "unknown.pdf")

    # 驗證副檔名
    if not original_filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    # 讀取並驗證檔案內容
    content = await file.read()

    # 驗證檔案大小
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large (max 50MB)")

    # 驗證 PDF 魔數（magic number）
    if not content.startswith(b'%PDF'):
        raise HTTPException(status_code=400, detail="Invalid PDF file format")

    # 建立 Job
    job_id = str(uuid.uuid4())[:8]
    job = Job(
        job_id=job_id,
        pdf_filename=original_filename,
        status=JobStatus.PENDING
    )

    # 設定封面欄位
    job.cover_report_no = report_no or ""
    job.cover_applicant_name = applicant_name or ""
    job.cover_applicant_address = applicant_address or ""

    # 上傳 PDF 到 Storage（content 已在前面讀取並驗證）
    storage = get_storage()
    pdf_key = f"jobs/{job_id}/original.pdf"
    storage.upload_bytes(content, pdf_key, content_type="application/pdf")
    job.original_pdf_key = pdf_key

    # 儲存 Job 到 Redis
    r = get_redis()
    r.set(f"job:{job_id}", job.to_json())
    r.lpush("job_queue", job_id)

    return JobResponse(
        job_id=job.job_id,
        status=job.status.value,
        created_at=job.created_at,
        updated_at=job.updated_at,
        pdf_filename=job.pdf_filename
    )


@app.get("/api/jobs/{job_id}", response_model=JobDetailResponse)
async def get_job(job_id: str, _: str = Depends(verify_password)):
    """
    查詢任務狀態

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    job = get_job_or_404(job_id)
    storage = get_storage()

    # 生成下載 URL
    docx_url = None
    qa_url = None

    if job.docx_key:
        docx_url = storage.get_presigned_url(job.docx_key, expires_in=3600)

    if job.qa_report_key:
        qa_url = storage.get_presigned_url(job.qa_report_key, expires_in=3600)

    return JobDetailResponse(
        job_id=job.job_id,
        status=job.status.value,
        created_at=job.created_at,
        updated_at=job.updated_at,
        pdf_filename=job.pdf_filename,
        docx_type=job.docx_type or None,
        error_message=job.error_message or None,
        qa_results=[r.to_dict() for r in job.qa_results],
        docx_download_url=docx_url,
        qa_report_download_url=qa_url
    )


@app.get("/api/jobs/{job_id}/download/docx")
async def download_docx(job_id: str, _: str = Depends(verify_password)):
    """
    下載 DOCX 檔案

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    job = get_job_or_404(job_id)

    if not job.docx_key:
        raise HTTPException(status_code=404, detail="DOCX not available")

    storage = get_storage()
    content = storage.download_bytes(job.docx_key)

    # 使用 safe_filename 防止路徑遍歷
    safe_pdf_name = safe_filename(job.pdf_filename)
    docx_type = (job.docx_type or "cns").lower()
    filename = f"{docx_type}_{safe_pdf_name.replace('.pdf', '.docx')}"

    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


@app.get("/api/jobs/{job_id}/download/qa")
async def download_qa_report(job_id: str, _: str = Depends(verify_password)):
    """
    下載 QA 報告

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    job = get_job_or_404(job_id)

    if not job.qa_report_key:
        raise HTTPException(status_code=404, detail="QA report not available")

    storage = get_storage()
    content = storage.download_bytes(job.qa_report_key)

    return StreamingResponse(
        iter([content]),
        media_type="application/json",
        headers={"Content-Disposition": f"attachment; filename=qa_report_{job_id}.json"}
    )


@app.get("/api/jobs/{job_id}/llm-stats")
async def get_llm_stats(job_id: str, _: str = Depends(verify_password)):
    """
    取得 LLM 翻譯統計（token 使用量與成本）

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    job = get_job_or_404(job_id)

    # 從 Job 物件取得 llm_stats
    if job.llm_stats:
        return job.llm_stats

    # 如果沒有，返回預設值
    return {
        "model": "gpt-5.1",
        "input_tokens": 0,
        "output_tokens": 0,
        "cached_tokens": 0,
        "total_cost": 0.0
    }


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str, _: str = Depends(verify_password)):
    """
    取消任務

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    job = get_job_or_404(job_id)

    # 只能取消 PENDING 或 RUNNING 的任務
    if job.status not in [JobStatus.PENDING, JobStatus.RUNNING]:
        raise HTTPException(
            status_code=400,
            detail=f"Cannot cancel job with status: {job.status.value}"
        )

    # 更新狀態為 CANCELLED
    job.update_status(JobStatus.CANCELLED)
    job.error_message = "任務已被使用者取消"

    # 儲存到 Redis
    r = get_redis()
    r.set(f"job:{job_id}", job.to_json())

    # 設定取消標記（讓 worker 檢測）
    r.set(f"job:{job_id}:cancel", "1", ex=3600)

    return {"status": "cancelled", "job_id": job_id}


@app.get("/api/jobs")
async def list_jobs(limit: int = Query(default=20, le=100), _: str = Depends(verify_password)):
    """
    列出最近的任務

    認證：若已設定 SHARED_PASSWORD，需在 query string 提供 ?p=<password>
    """
    r = get_redis()
    job_ids = []

    # 從 Redis 取得所有 job keys
    for key in r.scan_iter("job:*"):
        job_id = key.split(":")[1]
        job_ids.append(job_id)

    jobs = []
    for job_id in job_ids[:limit]:
        job_data = r.get(f"job:{job_id}")
        if job_data:
            job = Job.from_dict(json.loads(job_data))
            jobs.append({
                "job_id": job.job_id,
                "status": job.status.value,
                "pdf_filename": job.pdf_filename,
                "created_at": job.created_at,
                "docx_type": job.docx_type
            })

    # 按建立時間排序
    jobs.sort(key=lambda x: x["created_at"], reverse=True)

    return {"jobs": jobs[:limit]}


# ===== 靜態檔案 =====
# 如果存在 web 目錄，掛載靜態檔案
web_dir = Path(__file__).parent.parent / "web"
if web_dir.exists():
    app.mount("/", StaticFiles(directory=str(web_dir), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
