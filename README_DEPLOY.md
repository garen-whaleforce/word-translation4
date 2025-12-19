# CNS 15598-1 報告產生器 - 部署指南

## 架構

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│   Frontend  │────▶│   FastAPI   │────▶│   Redis     │
│  (Static)   │     │    (API)    │     │   Queue     │
└─────────────┘     └─────────────┘     └──────┬──────┘
                                               │
                    ┌─────────────┐             │
                    │   Worker    │◀────────────┘
                    │  (Python)   │
                    └──────┬──────┘
                           │
                    ┌──────▼──────┐
                    │   S3/R2     │
                    │  Storage    │
                    └─────────────┘
```

## Zeabur 部署步驟

### 1. 準備工作

1. 建立 GitHub Repository
2. 推送程式碼到 GitHub

```bash
git init
git add .
git commit -m "Initial commit: CNS Report Generator"
git remote add origin https://github.com/YOUR_USERNAME/word-translation3.git
git push -u origin main
```

### 2. 在 Zeabur 建立專案

1. 登入 [Zeabur](https://zeabur.com)
2. 點擊 "New Project"
3. 選擇 Region（建議選擇離使用者最近的區域）

### 3. 部署 Redis

1. 在專案中點擊 "Add Service"
2. 選擇 "Marketplace"
3. 搜尋並選擇 "Redis"
4. 等待部署完成，記下 Redis 連接 URL

### 4. 部署 API 服務

1. 點擊 "Add Service" > "Git"
2. 連接您的 GitHub Repository
3. 選擇 Repository
4. 設定環境變數：

```
SERVICE_TYPE=api
PORT=8000
REDIS_URL=<從 Redis 服務複製>
SHARED_PASSWORD=<您的密碼>
S3_ENDPOINT_URL=<Cloudflare R2 Endpoint>
S3_ACCESS_KEY=<R2 Access Key>
S3_SECRET_KEY=<R2 Secret Key>
S3_BUCKET_NAME=cns-reports
```

5. 等待建置和部署完成
6. 綁定自訂網域或使用 Zeabur 提供的網域

### 5. 部署 Worker 服務

1. 再次點擊 "Add Service" > "Git"
2. 選擇相同的 Repository
3. 設定環境變數：

```
SERVICE_TYPE=worker
REDIS_URL=<與 API 相同>
S3_ENDPOINT_URL=<與 API 相同>
S3_ACCESS_KEY=<與 API 相同>
S3_SECRET_KEY=<與 API 相同>
S3_BUCKET_NAME=<與 API 相同>
```

### 6. 設定 Cloudflare R2（可選但建議）

如果不設定 S3/R2，系統會使用本地存儲，但在 Container 重啟後資料會遺失。

1. 登入 [Cloudflare Dashboard](https://dash.cloudflare.com)
2. 進入 R2
3. 建立 Bucket（如 `cns-reports`）
4. 建立 API Token：
   - 進入 "Manage R2 API Tokens"
   - 建立新 Token，選擇 "Object Read & Write"
   - 記下 Access Key ID 和 Secret Access Key
5. Endpoint URL 格式：`https://<account-id>.r2.cloudflarestorage.com`

## 本地開發

### 使用 Docker Compose

```bash
# 啟動所有服務
docker-compose up -d

# 查看日誌
docker-compose logs -f

# 停止服務
docker-compose down
```

### 手動啟動

```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 啟動 Redis（需要先安裝 Redis）
redis-server

# 3. 啟動 API
python -m uvicorn apps.api.main:app --reload

# 4. 啟動 Worker（在另一個終端）
python -m apps.worker.run
```

### 測試

```bash
# 上傳 PDF
curl -X POST "http://localhost:8000/api/jobs?p=dev123" \
  -F "file=@templates/CB MC-601.pdf"

# 查詢狀態
curl "http://localhost:8000/api/jobs/{job_id}?p=dev123"

# 下載結果
curl -O "http://localhost:8000/api/jobs/{job_id}/download/docx?p=dev123"
```

## 環境變數說明

| 變數名 | 說明 | 預設值 |
|--------|------|--------|
| `REDIS_URL` | Redis 連接 URL | `redis://localhost:6379` |
| `SHARED_PASSWORD` | 共用存取密碼 | `cns2024` |
| `S3_ENDPOINT_URL` | S3/R2 Endpoint | - |
| `S3_ACCESS_KEY` | S3/R2 Access Key | - |
| `S3_SECRET_KEY` | S3/R2 Secret Key | - |
| `S3_BUCKET_NAME` | S3/R2 Bucket 名稱 | `cns-reports` |
| `LOCAL_STORAGE_PATH` | 本地存儲路徑 | `/tmp/cns-storage` |
| `SERVICE_TYPE` | 服務類型 (api/worker) | `api` |
| `PORT` | API 服務 Port | `8000` |
| `CORS_ORIGINS` | CORS 允許來源 | `*` |

## QA Gate 說明

系統包含多層 QA 驗證：

1. **extract_overview** - Overview 表格抽取驗證
2. **extract_clause_rows** - 條款列抽取驗證
3. **qa_overview_match** - PDF vs Word Overview 比對
4. **qa_clause_match** - PDF vs Word 條款表比對
5. **final_qa** - 最終合規檢查

只有所有 Gate 通過，才會產出 FINAL 報告；否則產出 DRAFT 報告供人工審閱。

## 疑難排解

### Worker 沒有處理任務

1. 確認 Redis 連接正常
2. 檢查 Worker 日誌
3. 確認 `job_queue` 有待處理任務

```bash
# 連接 Redis 檢查
redis-cli
> LLEN job_queue
> KEYS job:*
```

### PDF 抽取失敗

1. 確認 PDF 格式正確（CB TRF 報告）
2. 檢查 pdfplumber 是否正確安裝
3. 查看 Worker 日誌中的詳細錯誤

### 下載連結失效

1. 確認 S3/R2 設定正確
2. 檢查 Presigned URL 是否過期（預設 1 小時）
3. 確認 Bucket 存取權限
