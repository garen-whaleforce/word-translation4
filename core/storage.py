# core/storage.py
"""
S3/R2 Storage 抽象層
支援 Cloudflare R2 和 AWS S3 相容存儲
"""
import os
import json
from pathlib import Path
from typing import Optional, BinaryIO
import boto3
from botocore.config import Config


class StorageClient:
    """S3/R2 存儲客戶端"""

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

        if self.endpoint_url and self.access_key and self.secret_key:
            self.client = boto3.client(
                "s3",
                endpoint_url=self.endpoint_url,
                aws_access_key_id=self.access_key,
                aws_secret_access_key=self.secret_key,
                region_name=self.region,
                config=Config(signature_version="s3v4"),
            )
            self.enabled = True
        else:
            self.client = None
            self.enabled = False
            print("[Storage] S3/R2 未配置，使用本地文件存儲")

    def _local_path(self, key: str) -> Path:
        """生成本地存儲路徑"""
        base = Path(os.getenv("LOCAL_STORAGE_PATH", "/tmp/cns-storage"))
        path = base / key
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def upload_file(self, local_path: str, key: str) -> str:
        """上傳文件到存儲"""
        if self.enabled:
            self.client.upload_file(local_path, self.bucket_name, key)
            return f"s3://{self.bucket_name}/{key}"
        else:
            dest = self._local_path(key)
            dest.write_bytes(Path(local_path).read_bytes())
            return str(dest)

    def upload_bytes(self, data: bytes, key: str, content_type: str = "application/octet-stream") -> str:
        """上傳二進制數據"""
        if self.enabled:
            self.client.put_object(
                Bucket=self.bucket_name,
                Key=key,
                Body=data,
                ContentType=content_type,
            )
            return f"s3://{self.bucket_name}/{key}"
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
        if self.enabled:
            self.client.download_file(self.bucket_name, key, local_path)
        else:
            src = self._local_path(key)
            Path(local_path).write_bytes(src.read_bytes())
        return local_path

    def download_bytes(self, key: str) -> bytes:
        """下載二進制數據"""
        if self.enabled:
            response = self.client.get_object(Bucket=self.bucket_name, Key=key)
            return response["Body"].read()
        else:
            return self._local_path(key).read_bytes()

    def download_json(self, key: str) -> dict:
        """下載 JSON 數據"""
        data = self.download_bytes(key)
        return json.loads(data.decode("utf-8"))

    def get_presigned_url(self, key: str, expires_in: int = 3600) -> str:
        """生成預簽名下載 URL"""
        if self.enabled:
            return self.client.generate_presigned_url(
                "get_object",
                Params={"Bucket": self.bucket_name, "Key": key},
                ExpiresIn=expires_in,
            )
        else:
            # 本地存儲返回文件路徑
            return str(self._local_path(key))

    def exists(self, key: str) -> bool:
        """檢查文件是否存在"""
        if self.enabled:
            try:
                self.client.head_object(Bucket=self.bucket_name, Key=key)
                return True
            except:
                return False
        else:
            return self._local_path(key).exists()

    def delete(self, key: str) -> bool:
        """刪除文件"""
        if self.enabled:
            try:
                self.client.delete_object(Bucket=self.bucket_name, Key=key)
                return True
            except:
                return False
        else:
            path = self._local_path(key)
            if path.exists():
                path.unlink()
                return True
            return False

    def list_keys(self, prefix: str) -> list:
        """列出指定前綴的所有 key"""
        if self.enabled:
            response = self.client.list_objects_v2(
                Bucket=self.bucket_name,
                Prefix=prefix,
            )
            return [obj["Key"] for obj in response.get("Contents", [])]
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
