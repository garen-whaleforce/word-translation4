"""Application configuration using pydantic-settings."""
from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore"
    )

    # API Settings
    app_name: str = "CB PDF to Word Translation Service"
    debug: bool = False
    host: str = "0.0.0.0"
    port: int = 8000

    # File Storage
    upload_dir: Path = Path("/tmp/cb-uploads")
    output_dir: Path = Path("/tmp/cb-outputs")
    max_upload_size_mb: int = 50

    # LiteLLM Settings
    litellm_api_base: str = "https://litellm.whaleforce.dev"
    litellm_api_key: str = ""
    bulk_model: str = "gemini-2.5-flash"  # 快速翻譯
    refine_model: str = "gpt-5.2"  # 精修翻譯

    # Template Settings
    templates_dir: Path = Path("templates")
    default_template: str = "AST-B"

    # Processing Settings
    enable_refinement: bool = True
    start_from_overview: bool = True  # Start from "安全防護總攬表"


settings = Settings()

# Ensure directories exist
settings.upload_dir.mkdir(parents=True, exist_ok=True)
settings.output_dir.mkdir(parents=True, exist_ok=True)
