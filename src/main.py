"""FastAPI application entry point."""
import os
import uuid
import logging
from pathlib import Path
from typing import Optional
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
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


@app.get("/", tags=["System"])
async def root():
    """Root endpoint with API information."""
    return {
        "name": settings.app_name,
        "version": "0.1.0",
        "ui": "/ui",
        "endpoints": {
            "health": "/health",
            "upload": "/upload",
            "parse": "/parse (v2)",
            "translate": "/translate (v4)",
            "generate": "/generate (v6)",
            "templates": "/templates (v7)",
            "download": "/download/{file_id}"
        }
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=settings.host, port=settings.port)
