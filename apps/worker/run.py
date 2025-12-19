# apps/worker/run.py
"""
Redis Queue Worker
從 Redis 佇列取得任務並執行 CB PDF → CNS DOCX 轉換
"""
import os
import sys
import json
import time
import signal
from pathlib import Path
from datetime import datetime

import redis

# 添加專案根目錄到 path
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.models import Job, JobStatus
from core.pipeline import process_job
from core.storage import get_storage

# ===== 配置 =====
REDIS_URL = os.getenv("REDIS_URL", "redis://localhost:6379")
QUEUE_NAME = "job_queue"
POLL_INTERVAL = 2  # 秒
MAX_RETRIES = 3

# ===== 全局變數 =====
running = True


def signal_handler(signum, frame):
    global running
    print(f"\n[Worker] 收到信號 {signum}，準備停止...")
    running = False


def get_redis_client() -> redis.Redis:
    return redis.from_url(REDIS_URL, decode_responses=True)


def process_queue_item(r: redis.Redis, job_id: str) -> bool:
    """
    處理單個佇列項目

    Returns:
        bool: 是否處理成功
    """
    print(f"\n[Worker] 處理任務: {job_id}")

    try:
        # 讀取 Job
        job_data = r.get(f"job:{job_id}")
        if not job_data:
            print(f"[Worker] 任務 {job_id} 不存在，跳過")
            return False

        job = Job.from_dict(json.loads(job_data))

        # 檢查狀態
        if job.status != JobStatus.PENDING:
            print(f"[Worker] 任務 {job_id} 狀態為 {job.status.value}，跳過")
            return False

        # 更新狀態為 RUNNING
        job.update_status(JobStatus.RUNNING)
        r.set(f"job:{job_id}", job.to_json())

        print(f"[Worker] 開始處理 PDF: {job.pdf_filename}")

        # 執行 Pipeline
        storage = get_storage()
        job = process_job(job, storage)

        # 儲存最終狀態
        r.set(f"job:{job_id}", job.to_json())

        print(f"[Worker] 任務 {job_id} 完成，狀態: {job.status.value}")

        if job.status == JobStatus.PASS:
            print(f"[Worker] FINAL DOCX 已生成: {job.docx_key}")
        elif job.status == JobStatus.FAIL:
            print(f"[Worker] DRAFT DOCX 已生成: {job.docx_key}")
            print(f"[Worker] QA 問題: {len(job.qa_results)} 項")
        elif job.status == JobStatus.ERROR:
            print(f"[Worker] 錯誤: {job.error_message}")

        return True

    except Exception as e:
        import traceback
        print(f"[Worker] 處理任務 {job_id} 時發生錯誤: {e}")
        traceback.print_exc()

        # 嘗試更新 Job 狀態
        try:
            job_data = r.get(f"job:{job_id}")
            if job_data:
                job = Job.from_dict(json.loads(job_data))
                job.update_status(JobStatus.ERROR)
                job.error_message = str(e)
                r.set(f"job:{job_id}", job.to_json())
        except:
            pass

        return False


def run_worker():
    """
    主 Worker 迴圈
    """
    global running

    print("=" * 50)
    print("CNS Report Generator Worker")
    print("=" * 50)
    print(f"Redis URL: {REDIS_URL}")
    print(f"Queue: {QUEUE_NAME}")
    print(f"Poll Interval: {POLL_INTERVAL}s")
    print("=" * 50)

    # 設定信號處理
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    # 連接 Redis
    r = get_redis_client()

    try:
        r.ping()
        print("[Worker] Redis 連接成功")
    except Exception as e:
        print(f"[Worker] Redis 連接失敗: {e}")
        return

    # 檢查 Storage
    storage = get_storage()
    print(f"[Worker] Storage: {'S3/R2' if storage.enabled else 'Local'}")

    print("\n[Worker] 開始監聽佇列...")

    processed = 0
    errors = 0

    while running:
        try:
            # 從佇列取得任務（阻塞 POLL_INTERVAL 秒）
            result = r.brpop(QUEUE_NAME, timeout=POLL_INTERVAL)

            if result:
                _, job_id = result
                success = process_queue_item(r, job_id)
                if success:
                    processed += 1
                else:
                    errors += 1

        except redis.ConnectionError as e:
            print(f"[Worker] Redis 連接中斷: {e}")
            time.sleep(5)

        except Exception as e:
            print(f"[Worker] 未預期錯誤: {e}")
            errors += 1
            time.sleep(1)

    print("\n" + "=" * 50)
    print("[Worker] Worker 已停止")
    print(f"[Worker] 處理: {processed} 成功, {errors} 錯誤")
    print("=" * 50)


if __name__ == "__main__":
    run_worker()
