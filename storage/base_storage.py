"""
存储服务抽象基类，提供文件存储和管理功能
"""
from abc import ABC, abstractmethod
from typing import Optional, List


class BaseStorage(ABC):
    """存储服务抽象基类"""

    @abstractmethod
    def save(self, file_content: bytes, file_name: Optional[str] = None) -> str:
        """
        保存文件到存储
        
        Args:
            file_content: 文件二进制内容
            file_name: 可选的文件名，如不提供则自动生成
            
        Returns:
            存储的文件标识符
        """
        pass
    
    @abstractmethod
    def get(self, file_id: str) -> Optional[bytes]:
        """
        从存储获取文件
        
        Args:
            file_id: 文件标识符
            
        Returns:
            文件二进制内容，如文件不存在则返回None
        """
        pass
    
    @abstractmethod
    def delete(self, file_id: str) -> bool:
        """
        从存储删除文件
        
        Args:
            file_id: 文件标识符
            
        Returns:
            删除是否成功
        """
        pass
    
    @abstractmethod
    def cleanup_expired(self) -> int:
        """
        清理过期文件
        
        Returns:
            清理的文件数量
        """
        pass
    
    @abstractmethod
    def get_all_file_ids(self) -> List[str]:
        """
        获取所有存储的文件ID
        
        Returns:
            文件ID列表
        """
        pass 