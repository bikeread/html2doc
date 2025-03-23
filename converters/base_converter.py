"""
转换器基类实现，提供HTML到DOCX转换的抽象接口。
遵循适配器模式设计，便于扩展不同的转换实现。
"""
from abc import ABC, abstractmethod
import os
from typing import Union, BinaryIO


class BaseConverter(ABC):
    """HTML转DOCX转换器的抽象基类"""
    
    @abstractmethod
    def convert(self, html_content: str) -> Union[BinaryIO, bytes]:
        """
        将HTML内容转换为DOCX格式
        
        Args:
            html_content: HTML字符串内容
            
        Returns:
            文件对象或二进制数据
        """
        pass 