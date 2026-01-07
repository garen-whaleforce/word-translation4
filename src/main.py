"""FastAPI application entry point."""
import os
import uuid
import json
import logging
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, UploadFile, File, HTTPException, Query, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from .config import settings
from .cb_parser import CBParser
from .termbase import Termbase, load_termbase_from_json
from .translation_service import TranslationService, TranslationResult
from .pipeline import Pipeline, PipelineConfig, PipelineResult
from .template_registry import TemplateRegistry
from .pipeline_v2 import PipelineV2, run_pipeline_v2, PipelineV2Result

# Configure logging
logging.basicConfig(
    level=logging.DEBUG if settings.debug else logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

app = FastAPI(
    title=settings.app_name,
    description="CB PDF to Word Translation Service - 從 CB PDF 擷取結構化資料並翻譯回填至 Word 模板",
    version="0.1.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
static_dir = Path(__file__).parent.parent / "static"
if static_dir.exists():
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


# Response Models
class HealthResponse(BaseModel):
    status: str
    timestamp: str
    version: str


class UploadResponse(BaseModel):
    filename: str
    file_size: int
    file_id: str
    temp_path: str


class ErrorResponse(BaseModel):
    error: str
    detail: Optional[str] = None


class TranslateRequest(BaseModel):
    text: str
    model: Optional[str] = None
    enable_refinement: bool = True
    dry_run: bool = False


class TranslateResponse(BaseModel):
    original_text: str
    translated_text: str
    model_used: str
    was_refined: bool
    quality_issues: int


@app.get("/health", response_model=HealthResponse, tags=["System"])
async def health_check():
    """Health check endpoint."""
    return HealthResponse(
        status="healthy",
        timestamp=datetime.utcnow().isoformat(),
        version="0.1.0"
    )


@app.post("/upload", response_model=UploadResponse, tags=["Files"])
async def upload_file(file: UploadFile = File(...)):
    """
    Upload a CB PDF file for processing.

    Returns file metadata and temporary storage path.
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    # Check file size
    content = await file.read()
    file_size = len(content)
    max_size = settings.max_upload_size_mb * 1024 * 1024

    if file_size > max_size:
        raise HTTPException(
            status_code=400,
            detail=f"File too large. Maximum size is {settings.max_upload_size_mb}MB"
        )

    # Generate unique file ID and save
    file_id = str(uuid.uuid4())
    safe_filename = f"{file_id}_{file.filename}"
    temp_path = settings.upload_dir / safe_filename

    with open(temp_path, "wb") as f:
        f.write(content)

    logger.info(f"Uploaded file: {file.filename} ({file_size} bytes) -> {temp_path}")

    return UploadResponse(
        filename=file.filename,
        file_size=file_size,
        file_id=file_id,
        temp_path=str(temp_path)
    )


@app.post("/parse", tags=["Parse"])
async def parse_pdf(file: UploadFile = File(...)):
    """
    解析 CB PDF 並返回結構化 JSON

    擷取內容:
    - TRF No, Test Report No
    - Overview of Energy Sources and Safeguards (安全防護總攬表)
    - Energy Source Diagram (能量源圖)
    - Clause / Requirement / Result / Verdict (條款表)
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    # Save file temporarily
    content = await file.read()
    file_id = str(uuid.uuid4())
    temp_path = settings.upload_dir / f"{file_id}_{file.filename}"

    with open(temp_path, "wb") as f:
        f.write(content)

    try:
        # Parse PDF
        parser = CBParser(temp_path)
        result = parser.parse()

        logger.info(f"Parsed {file.filename}: {len(result.clauses)} clauses, {len(result.errors)} errors")

        return result.to_dict()

    except Exception as e:
        logger.error(f"Parse error: {e}")
        raise HTTPException(status_code=500, detail=f"PDF parsing failed: {str(e)}")

    finally:
        # Cleanup temp file
        if temp_path.exists():
            temp_path.unlink()


# Load termbase at startup
_termbase: Optional[Termbase] = None


def get_termbase() -> Termbase:
    """Get or load termbase."""
    global _termbase
    if _termbase is None:
        glossary_path = Path("rules/en_zh_glossary_preferred.json")
        if glossary_path.exists():
            _termbase = load_termbase_from_json(glossary_path)
            logger.info(f"Loaded termbase with {len(_termbase)} terms")
        else:
            _termbase = Termbase()
            logger.warning("Termbase file not found, using empty termbase")
    return _termbase


@app.post("/translate", response_model=TranslateResponse, tags=["Translation"])
async def translate_text(request: TranslateRequest):
    """
    翻譯英文文字為繁體中文

    功能:
    - 使用術語庫保護關鍵術語
    - Bulk 翻譯 + 可選 Refinement 精修
    - 支援乾跑模式 (不呼叫 LLM)
    """
    try:
        termbase = get_termbase()
        service = TranslationService(
            termbase=termbase,
            dry_run=request.dry_run
        )

        result = service.translate(
            text=request.text,
            enable_refinement=request.enable_refinement,
            model=request.model
        )

        return TranslateResponse(
            original_text=result.original_text,
            translated_text=result.translated_text,
            model_used=result.model_used,
            was_refined=result.was_refined,
            quality_issues=result.metadata.get("issues_found", 0)
        )

    except Exception as e:
        logger.error(f"Translation error: {e}")
        raise HTTPException(status_code=500, detail=f"Translation failed: {str(e)}")


@app.post("/generate", tags=["Generate"])
async def generate_word(
    file: UploadFile = File(...),
    template: Optional[str] = Query(None, description="模板 ID 或路徑"),
    dry_run: bool = Query(False, description="乾跑模式 (不呼叫 LLM)")
):
    """
    端到端生成 Word 文件

    流程:
    1. 解析 PDF → 結構化 JSON
    2. 翻譯 (使用術語庫保護)
    3. 填充 Word 模板
    4. 返回下載連結

    Returns:
        生成的 docx 檔案下載連結及處理報告
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    # Save PDF temporarily
    content = await file.read()
    file_id = str(uuid.uuid4())
    pdf_path = settings.upload_dir / f"{file_id}_{file.filename}"

    with open(pdf_path, "wb") as f:
        f.write(content)

    # Output directory
    output_dir = settings.output_dir / file_id

    # Determine template path
    template_path = None
    if template:
        # Try as path first
        if Path(template).exists():
            template_path = template
        else:
            # Try in templates directory
            templates_dir = Path("templates")
            for ext in ['.docx', '']:
                candidate = templates_dir / f"{template}{ext}"
                if candidate.exists():
                    template_path = candidate
                    break

    try:
        # Run pipeline
        config = PipelineConfig(
            template_path=Path(template_path) if template_path else None,
            dry_run=dry_run
        )
        pipeline = Pipeline(config)
        result = pipeline.process(pdf_path, output_dir, template_path)

        logger.info(f"Generated output for {file.filename}: {result.total_clauses} clauses")

        # Build response
        response_data = {
            "status": "success" if not result.errors else "partial",
            "file_id": file_id,
            "pdf_filename": result.pdf_filename,
            "template_used": result.template_used,
            "total_clauses": result.total_clauses,
            "translated_segments": result.translated_segments,
            "output_files": {
                "docx": result.output_docx if result.output_docx else None,
                "extracted_json": result.extracted_json,
                "qa_report": result.qa_report_json
            },
            "errors": result.errors,
            "warnings": result.warnings
        }

        # Add download URL if docx was generated
        if result.output_docx:
            response_data["download_url"] = f"/download/{file_id}"

        return response_data

    except Exception as e:
        logger.error(f"Generate error: {e}")
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")

    finally:
        # Cleanup PDF (keep output)
        if pdf_path.exists():
            pdf_path.unlink()


@app.get("/download/{file_id}", tags=["Files"])
async def download_file(file_id: str):
    """下載生成的 Word 文件"""
    output_dir = settings.output_dir / file_id

    if not output_dir.exists():
        raise HTTPException(status_code=404, detail="File not found")

    # Find the docx file
    docx_files = list(output_dir.glob("*_output.docx"))
    if not docx_files:
        raise HTTPException(status_code=404, detail="Output file not found")

    return FileResponse(
        docx_files[0],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=docx_files[0].name
    )


@app.get("/templates", tags=["Templates"])
async def list_templates():
    """列出可用的 Word 模板"""
    try:
        registry = TemplateRegistry(Path("templates"))
        templates = registry.list_templates()

        return {
            "templates": [t.to_dict() for t in templates],
            "total": len(templates)
        }

    except Exception as e:
        logger.error(f"Failed to list templates: {e}")
        return {"templates": [], "total": 0, "error": str(e)}


@app.get("/ui", response_class=HTMLResponse, tags=["UI"])
async def ui():
    """前端頁面"""
    html_path = Path(__file__).parent.parent / "static" / "index.html"
    if html_path.exists():
        return HTMLResponse(content=html_path.read_text(encoding='utf-8'))
    return HTMLResponse(content="<h1>UI not available</h1>", status_code=404)


# ============================================================
# V2 API Endpoints - 使用新的 Pipeline v2
# ============================================================

# 背景任務狀態儲存
_job_status: Dict[str, Dict[str, Any]] = {}


class JobStatus(BaseModel):
    job_id: str
    status: str  # pending, processing, completed, failed
    progress: int = 0
    message: str = ""
    result: Optional[Dict[str, Any]] = None
    created_at: str = ""
    completed_at: Optional[str] = None


def _run_pipeline_v2_job(job_id: str, pdf_path: Path, output_dir: Path, template_path: Optional[Path]):
    """背景執行 Pipeline v2"""
    try:
        _job_status[job_id]["status"] = "processing"
        _job_status[job_id]["progress"] = 10
        _job_status[job_id]["message"] = "解析 PDF..."

        result = run_pipeline_v2(
            pdf_path=pdf_path,
            output_dir=output_dir,
            template_path=template_path,
            translate_func=None,  # TODO: 加入翻譯函數
            dry_run=False
        )

        _job_status[job_id]["status"] = "completed"
        _job_status[job_id]["progress"] = 100
        _job_status[job_id]["message"] = "完成"
        _job_status[job_id]["completed_at"] = datetime.utcnow().isoformat()
        _job_status[job_id]["result"] = result.to_dict()

    except Exception as e:
        logger.error(f"Job {job_id} failed: {e}")
        _job_status[job_id]["status"] = "failed"
        _job_status[job_id]["message"] = str(e)
        _job_status[job_id]["completed_at"] = datetime.utcnow().isoformat()

    finally:
        # 清理上傳的 PDF
        if pdf_path.exists():
            try:
                pdf_path.unlink()
            except:
                pass


@app.post("/api/v2/process", tags=["V2 API"])
async def process_pdf_v2(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    template: Optional[str] = Query("AST-B", description="模板 ID (AST-B, Generic-CB)"),
    async_mode: bool = Query(False, description="非同步模式 (背景執行)")
):
    """
    V2 API: 處理 CB PDF

    整合功能:
    - PDF 區塊偵測 (Overview, Energy Diagram, Clause, Appended Tables)
    - Clause 按章節分群
    - 附表提取與回填
    - QA 報告生成

    Returns:
        同步模式: 直接返回處理結果
        非同步模式: 返回 job_id，可透過 /api/v2/job/{job_id} 查詢狀態
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted")

    # Save PDF temporarily
    content = await file.read()
    file_id = str(uuid.uuid4())[:8]
    pdf_path = settings.upload_dir / f"{file_id}_{file.filename}"

    with open(pdf_path, "wb") as f:
        f.write(content)

    # Output directory
    output_dir = settings.output_dir / file_id

    # Determine template path
    template_path = None
    if template:
        templates_dir = Path("templates")
        for name in [template, f"{template}.docx"]:
            candidate = templates_dir / name
            if candidate.exists():
                template_path = candidate
                break

    if async_mode:
        # 非同步模式 - 背景執行
        job_id = file_id
        _job_status[job_id] = {
            "job_id": job_id,
            "status": "pending",
            "progress": 0,
            "message": "排隊中...",
            "result": None,
            "created_at": datetime.utcnow().isoformat(),
            "completed_at": None,
            "pdf_filename": file.filename,
            "template": template
        }

        background_tasks.add_task(
            _run_pipeline_v2_job,
            job_id, pdf_path, output_dir, template_path
        )

        return {
            "mode": "async",
            "job_id": job_id,
            "status_url": f"/api/v2/job/{job_id}",
            "message": "Job submitted, check status_url for progress"
        }

    else:
        # 同步模式 - 直接執行
        try:
            result = run_pipeline_v2(
                pdf_path=pdf_path,
                output_dir=output_dir,
                template_path=template_path,
                translate_func=None,
                dry_run=False
            )

            response_data = {
                "status": "success" if not result.errors else "partial",
                "file_id": file_id,
                "pdf_filename": result.pdf_filename,
                "template_used": result.template_used,
                "total_clauses": result.total_clauses,
                "chapters_count": result.chapters_count,
                "appended_tables_count": result.appended_tables_count,
                "output_files": {
                    "docx": result.output_docx if result.output_docx else None,
                    "extracted_json": result.extracted_json,
                    "qa_report": result.qa_report_json,
                    "energy_diagram": result.energy_diagram_png if result.energy_diagram_png else None
                },
                "qa_report": result.qa_report.to_dict() if result.qa_report else None,
                "errors": result.errors,
                "warnings": result.warnings[:10]  # 限制警告數量
            }

            # Add download URL if docx was generated
            if result.output_docx:
                response_data["download_url"] = f"/download/{file_id}"

            return response_data

        except Exception as e:
            logger.error(f"Process error: {e}")
            raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")

        finally:
            # Cleanup PDF
            if pdf_path.exists():
                pdf_path.unlink()


@app.get("/api/v2/job/{job_id}", response_model=JobStatus, tags=["V2 API"])
async def get_job_status(job_id: str):
    """查詢背景任務狀態"""
    if job_id not in _job_status:
        raise HTTPException(status_code=404, detail="Job not found")

    job = _job_status[job_id]
    return JobStatus(
        job_id=job["job_id"],
        status=job["status"],
        progress=job.get("progress", 0),
        message=job.get("message", ""),
        result=job.get("result"),
        created_at=job.get("created_at", ""),
        completed_at=job.get("completed_at")
    )


@app.get("/api/v2/jobs", tags=["V2 API"])
async def list_jobs():
    """列出所有任務"""
    jobs = []
    for job_id, job in _job_status.items():
        jobs.append({
            "job_id": job_id,
            "status": job["status"],
            "pdf_filename": job.get("pdf_filename", ""),
            "created_at": job.get("created_at", ""),
            "completed_at": job.get("completed_at")
        })
    return {"jobs": jobs, "total": len(jobs)}


@app.get("/", tags=["System"])
async def root():
    """Root endpoint with API information."""
    return {
        "name": settings.app_name,
        "version": "2.0.0",
        "ui": "/ui",
        "endpoints": {
            "health": "/health",
            "upload": "/upload",
            "parse": "/parse (v1)",
            "translate": "/translate",
            "generate": "/generate (v1)",
            "templates": "/templates",
            "download": "/download/{file_id}",
            "v2_process": "/api/v2/process (v2 - 推薦)",
            "v2_job_status": "/api/v2/job/{job_id}",
            "v2_jobs": "/api/v2/jobs"
        }
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=settings.host, port=settings.port)
