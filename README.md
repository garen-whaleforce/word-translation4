# CB PDF to Word Translation Service

將 CB (Certification Body) PDF 報告自動擷取、翻譯並回填至 Word 模板的服務。

## 功能特色

- **PDF 結構化擷取**：自動解析 CB PDF 中的 Overview 表、Energy Source Diagram、Clause 主表
- **術語庫約束翻譯**：使用 placeholder 機制強制遵守術語表
- **雙階段翻譯**：Bulk 快速翻譯 + Refinement 精修
- **Word 模板回填**：保持頁首/格式不變，表格可自動擴列
- **一致性驗證**：PDF vs Word 結構比對報告

## 快速開始

### 本機執行

```bash
# 1. 建立虛擬環境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 設定環境變數
cp .env.example .env
# 編輯 .env 設定 LiteLLM API Key 等

# 4. 啟動服務
uvicorn src.main:app --reload --port 8000
```

### Docker 執行

```bash
# 建置映像
docker build -t cb-translator .

# 執行容器
docker run -p 8000:8000 \
  -e LITELLM_API_KEY=your-key \
  -v $(pwd)/templates:/app/templates \
  cb-translator
```

### Zeabur 部署

1. 連接 Git Repository
2. 設定環境變數 (參考 `.env.example`)
3. 部署即可

## API 端點

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | 健康檢查 |
| `/upload` | POST | 上傳 PDF 檔案 |
| `/parse` | POST | 解析 PDF 結構 (v2) |
| `/translate` | POST | 翻譯文字 (v4) |
| `/generate` | POST | 端到端生成 Word (v6) |
| `/templates` | GET | 列出可用模板 (v7) |

## 使用範例

### 上傳 PDF

```bash
curl -X POST "http://localhost:8000/upload" \
  -F "file=@CB-Report.pdf"
```

### 解析 PDF (v2+)

```bash
curl -X POST "http://localhost:8000/parse" \
  -F "file=@CB-Report.pdf"
```

### 生成 Word (v6+)

```bash
curl -X POST "http://localhost:8000/generate" \
  -F "file=@CB-Report.pdf" \
  -F "template=AST-B"
```

## 專案結構

```
.
├── src/
│   ├── main.py              # FastAPI 應用入口
│   ├── config.py            # 設定管理
│   ├── cb_parser/           # PDF 解析模組 (v2)
│   ├── termbase/            # 術語庫模組 (v3)
│   ├── translation_service/ # 翻譯服務 (v4)
│   ├── word_filler/         # Word 回填模組 (v5)
│   ├── template_registry/   # 模板管理 (v7)
│   └── validator/           # 驗證模組 (v8)
├── rules/
│   ├── en_zh_glossary_preferred.json  # 術語表
│   ├── en_zh_translation_memory.csv   # 翻譯記憶庫
│   └── IEC62368_EN2ZH_translation_guideline.md
├── templates/               # Word 模板
├── tests/                   # 測試
├── scripts/                 # CLI 工具
├── Dockerfile
├── requirements.txt
└── README.md
```

## 翻譯規則

系統使用以下資源確保翻譯品質：

1. **術語表** (`rules/en_zh_glossary_preferred.json`)：強制術語對照
2. **翻譯記憶庫** (`rules/en_zh_translation_memory.csv`)：歷史翻譯參考
3. **翻譯指南** (`rules/IEC62368_EN2ZH_translation_guideline.md`)：IEC 標準翻譯規範

## 開發

### 執行測試

```bash
pytest tests/ -v
```

### 程式碼檢查

```bash
ruff check src/
```

## 版本歷程

- **v1**: 專案骨架 (FastAPI + Dockerfile)
- **v2**: PDF 擷取 (Overview/Energy Diagram/Clause 表)
- **v3**: 術語庫 + placeholder 機制
- **v4**: LiteLLM 翻譯服務 (Bulk + Refinement)
- **v5**: Word 模板回填
- **v6**: 端到端管線
- **v7**: 模板自動選擇
- **v8**: 一致性驗證報告
- **v9**: 前端頁面
- **v10**: 部署強化

## License

Internal Use Only
