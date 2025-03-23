"""
令牌服务，处理临时下载链接的生成与验证
"""
import time
import jwt
from typing import Dict, Any, Optional, Union


class TokenService:
    """令牌服务类"""
    
    def __init__(self, secret_key: str, default_expires: int = 300, max_expires: int = 3600):
        """
        初始化令牌服务
        
        Args:
            secret_key: 加密密钥
            default_expires: 默认过期时间（秒），默认5分钟
            max_expires: 最大过期时间（秒），默认1小时
        """
        self.secret_key = secret_key
        self.default_expires = int(default_expires)
        self.max_expires = int(max_expires)
    
    def generate_download_token(self, file_id: str, expires_in: Optional[Union[int, str]] = None) -> str:
        """
        生成下载令牌
        
        Args:
            file_id: 文件ID
            expires_in: 过期时间（秒），如不提供则使用默认值
            
        Returns:
            JWT令牌字符串
        """
        # 确保expires_in是整数
        try:
            expires_in_int = int(expires_in) if expires_in is not None else None
        except (ValueError, TypeError):
            expires_in_int = None
            
        # 使用默认过期时间，如果没有提供或超出最大值
        if expires_in_int is None or expires_in_int <= 0:
            expires_in_int = self.default_expires
        elif expires_in_int > self.max_expires:
            expires_in_int = self.max_expires
        
        # 创建JWT载荷
        payload = {
            'file_id': file_id,
            'exp': int(time.time()) + expires_in_int,
            'iat': int(time.time())
        }
        
        # 生成令牌
        token = jwt.encode(payload, self.secret_key, algorithm='HS256')
        return token
    
    def verify_download_token(self, token: str) -> Optional[str]:
        """
        验证下载令牌
        
        Args:
            token: JWT令牌字符串
            
        Returns:
            如果令牌有效，返回文件ID；否则返回None
        """
        try:
            # 解码并验证令牌
            payload = jwt.decode(token, self.secret_key, algorithms=['HS256'])
            return payload.get('file_id')
        except jwt.PyJWTError:
            return None 