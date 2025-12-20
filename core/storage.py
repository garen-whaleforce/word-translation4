# core/storage.py
"""
Storage 抽象層
支援 S3/R2、Redis、本地存儲三種模式
"""
import os
import json
import base64
from pathlib import Path
from typing import Optional
import redis

# 嘗試導入 boto3（S3 模式需要）
try:
    import boto3
    from botocore.config import Config
    HAS_BOTO3 = True
except ImportError:
    HAS_BOTO3 = False


class StorageClient:
    """Storage 客戶端 - 支援 S3/R2、Redis、本地存儲"""

    def __init__(
        self,
        endpoint_url: Optional[str] = None,
        access_key: Optional[str] = None,
        secret_key: Optional[str] = None,
        bucket_name: Optional[str] = None,
        region: str = "auto",
    ):
        self.endpoint_url = endpoint_url or os.getenv("S3_ENDPOINT_URL")
        self.access_key = access_key or os.getenv("S3_ACCESS_KEY")
        self.secret_key = secret_key or os.getenv("S3_SECRET_KEY")
        self.bucket_name = bucket_name or os.getenv("S3_BUCKET_NAME", "cns-reports")
        self.region = region

        # 存儲模式: s3, redis, local
        self.mode = "local"
        self.client = None
        self.redis_client = None

        # 優先使用 S3/R2
        if HAS_BOTO3 and self.endpoint_url and self.access_key and self.secret_key:
            self.client = boto3.client(
                "s3",
                endpoint_url=self.endpoint_url,
                aws_access_key_id=self.access_key,
                aws_secret_access_key=self.secret_key,
                region_name=self.region,
                config=Config(signature_version="s3v4"),
            )
            self.mode = "s3"
            print("[Storage] 使用 S3/R2 存儲")
        else:
            # 嘗試使用 Redis 存儲
            redis_url = os.getenv("REDIS_URI") or os.getenv("REDIS_URL")
            if redis_url:
                try:
                    self.redis_client = redis.from_url(redis_url, decode_responses=False)
                    self.redis_client.ping()
                    self.mode = "redis"
                    print("[Storage] 使用 Redis 存儲")
                except Exception as e:
                    print(f"[Storage] Redis 連接失敗: {e}，降級為本地存儲")
                    self.mode = "local"
            else:
                print("[Storage] 使用本地文件存儲")

    @property
    def enabled(self) -> bool:
        """兼容舊代碼 - 判斷是否啟用雲存儲"""
        return self.mode == "s3"

    def _local_path(self, key: str) -> Path:
        """生成本地存儲路徑"""
        base = Path(os.getenv("LOCAL_STORAGE_PATH", "/tmp/cns-storage"))
        path = base / key
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def _redis_key(self, key: str) -> str:
        """生成 Redis key"""
        return f"file:{key}"

    def upload_file(self, local_path: str, key: str) -> str:
        """上傳文件到存儲"""
        data = Path(local_path).read_bytes()
        return self.upload_bytes(data, key)

    def upload_bytes(self, data: bytes, key: str, content_type: str = "application/octet-stream") -> str:
        """上傳二進制數據"""
        if self.mode == "s3":
            self.client.put_object(
                Bucket=self.bucket_name,
                Key=key,
                Body=data,
                ContentType=content_type,
            )
            return f"s3://{self.bucket_name}/{key}"
        elif self.mode == "redis":
            # 存儲到 Redis（設定 24 小時過期）
            self.redis_client.setex(
                self._redis_key(key),
                86400,  # 24 小時
                data
            )
            return f"redis://{key}"
        else:
            dest = self._local_path(key)
            dest.write_bytes(data)
            return str(dest)

    def upload_json(self, data: dict, key: str) -> str:
        """上傳 JSON 數據"""
        json_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        return self.upload_bytes(json_bytes, key, content_type="application/json")

    def download_file(self, key: str, local_path: str) -> str:
        """下載文件到本地"""
        data = self.download_bytes(key)
        Path(local_path).parent.mkdir(parents=True, exist_ok=True)
        Path(local_path).write_bytes(data)
        return local_path

    def download_bytes(self, key: str) -> bytes:
        """下載二進制數據"""
        if self.mode == "s3":
            response = self.client.get_object(Bucket=self.bucket_name, Key=key)
            return response["Body"].read()
        elif self.mode == "redis":
            data = self.redis_client.get(self._redis_key(key))
            if data is None:
                raise FileNotFoundError(f"File not found in Redis: {key}")
            return data
        else:
            return self._local_path(key).read_bytes()

    def download_json(self, key: str) -> dict:
        """下載 JSON 數據"""
        data = self.download_bytes(key)
        return json.loads(data.decode("utf-8"))

    def get_presigned_url(self, key: str, expires_in: int = 3600) -> str:
        """生成預簽名下載 URL"""
        if self.mode == "s3":
            return self.client.generate_presigned_url(
                "get_object",
                Params={"Bucket": self.bucket_name, "Key": key},
                ExpiresIn=expires_in,
            )
        else:
            # Redis 和本地存儲不支援預簽名 URL，返回 API 下載路徑
            return f"/api/jobs/{key.split('/')[1]}/download/docx"

    def exists(self, key: str) -> bool:
        """檢查文件是否存在"""
        if self.mode == "s3":
            try:
                self.client.head_object(Bucket=self.bucket_name, Key=key)
                return True
            except:
                return False
        elif self.mode == "redis":
            return self.redis_client.exists(self._redis_key(key)) > 0
        else:
            return self._local_path(key).exists()

    def delete(self, key: str) -> bool:
        """刪除文件"""
        if self.mode == "s3":
            try:
                self.client.delete_object(Bucket=self.bucket_name, Key=key)
                return True
            except:
                return False
        elif self.mode == "redis":
            return self.redis_client.delete(self._redis_key(key)) > 0
        else:
            path = self._local_path(key)
            if path.exists():
                path.unlink()
                return True
            return False

    def list_keys(self, prefix: str) -> list:
        """列出指定前綴的所有 key"""
        if self.mode == "s3":
            response = self.client.list_objects_v2(
                Bucket=self.bucket_name,
                Prefix=prefix,
            )
            return [obj["Key"] for obj in response.get("Contents", [])]
        elif self.mode == "redis":
            pattern = f"file:{prefix}*"
            keys = self.redis_client.keys(pattern)
            return [k.decode().replace("file:", "") for k in keys]
        else:
            base = self._local_path(prefix).parent
            if not base.exists():
                return []
            return [str(p.relative_to(self._local_path("").parent)) for p in base.glob("**/*") if p.is_file()]


# 全局存儲客戶端實例
_storage_client: Optional[StorageClient] = None


def get_storage() -> StorageClient:
    """獲取全局存儲客戶端"""
    global _storage_client
    if _storage_client is None:
        _storage_client = StorageClient()
    return _storage_client
