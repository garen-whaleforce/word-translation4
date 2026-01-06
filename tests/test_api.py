"""Tests for the FastAPI endpoints."""
import pytest
from fastapi.testclient import TestClient
from io import BytesIO

from src.main import app


@pytest.fixture
def client():
    """Create test client."""
    return TestClient(app)


def test_health_check(client):
    """Test health endpoint returns healthy status."""
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "healthy"
    assert "timestamp" in data
    assert data["version"] == "0.1.0"


def test_root_endpoint(client):
    """Test root endpoint returns API info."""
    response = client.get("/")
    assert response.status_code == 200
    data = response.json()
    assert "name" in data
    assert "endpoints" in data


def test_upload_pdf_success(client, tmp_path):
    """Test successful PDF upload."""
    # Create a minimal PDF-like content
    pdf_content = b"%PDF-1.4\ntest content"

    response = client.post(
        "/upload",
        files={"file": ("test.pdf", BytesIO(pdf_content), "application/pdf")}
    )

    assert response.status_code == 200
    data = response.json()
    assert data["filename"] == "test.pdf"
    assert data["file_size"] > 0
    assert "file_id" in data
    assert "temp_path" in data


def test_upload_non_pdf_rejected(client):
    """Test that non-PDF files are rejected."""
    response = client.post(
        "/upload",
        files={"file": ("test.txt", BytesIO(b"test content"), "text/plain")}
    )

    assert response.status_code == 400
    assert "PDF" in response.json()["detail"]


def test_upload_empty_file(client):
    """Test uploading empty file."""
    response = client.post(
        "/upload",
        files={"file": ("empty.pdf", BytesIO(b""), "application/pdf")}
    )

    # Empty file should still be accepted (validation happens later)
    assert response.status_code == 200
