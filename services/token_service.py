"""
令牌服务，处理临时下载链接的生成与验证
"""
import time
import jwt
import os
import base64
import hashlib
from typing import Dict, Any, Optional, Union
import sys
from importlib import import_module


class TokenService:
    """令牌服务类"""
    
    def __init__(self, secret_key: str, default_expires: int = 300):
        """
        初始化令牌服务
        
        Args:
            secret_key: 加密密钥
            default_expires: 默认过期时间（秒），默认5分钟
            max_expires: 最大过期时间（秒），默认1小时
        """
        self.secret_key = secret_key
        self.default_expires = int(default_expires)

    def generate_short_token(self, file_id: str, expires_in: Optional[Union[int, str]] = None) -> str:
        """
        生成短下载令牌
        
        Args:
            file_id: 文件ID
            expires_in: 过期时间（秒），如不提供则使用默认值
            
        Returns:
            短令牌字符串
        """
        # 确保expires_in是整数
        try:
            expires_in_int = int(expires_in) if expires_in is not None else None
        except (ValueError, TypeError):
            expires_in_int = None
            
        # 创建简化的数据结构
        exp_time = int(time.time()) + expires_in_int
        
        # 生成唯一随机值加入混淆
        random_bytes = os.urandom(3)  # 减少到3个字节
        
        # 对文件ID进行哈希处理，保证安全性
        file_id_hash = hashlib.md5(file_id.encode()).digest()[:5]  # 取MD5前5个字节
        
        # 将过期时间编码到token中
        exp_bytes = int(exp_time).to_bytes(4, 'big')
        
        # 组合最终token成分: 随机字节+过期时间+文件ID哈希
        final_bytes = random_bytes + exp_bytes + file_id_hash
        
        # 使用base64url编码，并移除填充字符
        token = base64.urlsafe_b64encode(final_bytes).decode('utf-8').rstrip('=')
        
        # 加密签名，保证token不被篡改
        token_signature = hashlib.sha256((token + self.secret_key).encode()).digest()[:3]
        signature_b64 = base64.urlsafe_b64encode(token_signature).decode('utf-8').rstrip('=')
        
        # 组合最终token
        return f"{token}{signature_b64}"
    
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
            
    def verify_short_token(self, token: str) -> Optional[str]:
        """
        验证短下载令牌
        
        Args:
            token: 短令牌字符串
            
        Returns:
            如果令牌有效返回文件ID；否则返回None
        """
        try:
            # 添加调试信息
            print(f"开始验证短链接令牌: {token}")
            
            # 获取签名部分（最后6个字符左右）
            signature_length = 4  # base64(3 bytes) = 4 chars
            main_token = token[:-signature_length]
            signature = token[-signature_length:]
            
            print(f"主令牌部分: {main_token}, 签名部分: {signature}")
            
            # 验证签名
            expected_signature = base64.urlsafe_b64encode(
                hashlib.sha256((main_token + self.secret_key).encode()).digest()[:3]
            ).decode('utf-8').rstrip('=')
            
            print(f"计算的期望签名: {expected_signature}")
            
            if signature != expected_signature:
                print("签名验证失败")
                return None
            
            # 添加回可能被移除的填充字符
            padding = '=' * (4 - len(main_token) % 4) if len(main_token) % 4 else ''
            decoded = base64.urlsafe_b64decode(main_token + padding)
            
            # 解析token成分
            random_bytes = decoded[:3]
            exp_time = int.from_bytes(decoded[3:7], 'big')
            file_id_hash = decoded[7:]
            
            print(f"令牌解码: 随机值={random_bytes.hex()}, 过期时间={exp_time}, 文件ID哈希={file_id_hash.hex()}")
            
            # 验证是否过期
            current_time = int(time.time())
            print(f"当前时间: {current_time}, 过期时间: {exp_time}, 剩余时间: {exp_time - current_time}秒")
            
            if exp_time < current_time:
                print("令牌已过期")
                return None
            
            # 简化实现：直接使用全局存储实例
            # 由于我们确认在app.py中已经有全局的storage变量
            from app import storage
            print(f"获取文件ID列表: {storage.__class__.__name__}")
            
            all_file_ids = storage.get_all_file_ids()
            print(f"找到{len(all_file_ids)}个文件ID")
            
            for candidate_id in all_file_ids:
                candidate_hash = hashlib.md5(candidate_id.encode()).digest()[:5]
                if candidate_hash == file_id_hash:
                    print(f"找到匹配的文件ID: {candidate_id}")
                    return candidate_id
            
            print("未找到匹配的文件ID")
            return None
        except Exception as e:
            print(f"短链接验证出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return None 