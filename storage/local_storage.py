"""
本地文件系统存储实现
"""
import os
import time
import uuid
from typing import Optional, Dict, List
from datetime import datetime

from storage.base_storage import BaseStorage


class LocalStorage(BaseStorage):
    """本地文件系统存储实现"""
    
    def __init__(self, storage_path: str, retention_seconds: int = 600):
        """
        初始化本地存储
        
        Args:
            storage_path: 存储文件的目录路径
            retention_seconds: 文件保留时间（秒），默认10分钟
        """
        self.storage_path = storage_path
        self.retention_seconds = retention_seconds
        self.file_metadata: Dict[str, Dict] = {}
        
        # 确保存储目录存在
        os.makedirs(self.storage_path, exist_ok=True)
    
    def save(self, file_content: bytes, file_name: Optional[str] = None) -> str:
        """
        保存文件到本地存储
        
        Args:
            file_content: 文件二进制内容
            file_name: 可选的文件名，如不提供则自动生成
            
        Returns:
            文件ID（UUID）
        """
        # 生成唯一文件ID和文件名
        file_id = str(uuid.uuid4())
        if file_name is None:
            file_name = f"{file_id}.docx"
        
        # 构建完整文件路径并保存
        file_path = os.path.join(self.storage_path, file_name)
        with open(file_path, 'wb') as f:
            f.write(file_content)
        
        # 记录文件元数据
        self.file_metadata[file_id] = {
            'file_path': file_path,
            'created_at': time.time(),
            'expires_at': time.time() + self.retention_seconds
        }
        
        return file_id
    
    def get(self, file_id: str) -> Optional[bytes]:
        """
        从本地存储获取文件
        
        Args:
            file_id: 文件标识符
            
        Returns:
            文件二进制内容，如文件不存在则返回None
        """
        metadata = self.file_metadata.get(file_id)
        if metadata is None:
            return None
        
        file_path = metadata['file_path']
        if not os.path.exists(file_path):
            return None
        
        with open(file_path, 'rb') as f:
            return f.read()
    
    def delete(self, file_id: str) -> bool:
        """
        从本地存储删除文件
        
        Args:
            file_id: 文件标识符
            
        Returns:
            删除是否成功
        """
        metadata = self.file_metadata.get(file_id)
        if metadata is None:
            return False
        
        file_path = metadata['file_path']
        if os.path.exists(file_path):
            os.remove(file_path)
        
        # 移除元数据
        self.file_metadata.pop(file_id, None)
        return True
    
    def cleanup_expired(self) -> int:
        """
        清理过期文件
        
        Returns:
            清理的文件数量
        """
        now = time.time()
        expired_ids = [
            file_id for file_id, metadata in self.file_metadata.items()
            if metadata['expires_at'] < now
        ]
        
        count = 0
        for file_id in expired_ids:
            if self.delete(file_id):
                count += 1
        
        return count 