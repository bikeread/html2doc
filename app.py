"""
HTML2DOC转换服务 - 主程序
提供REST API接口，将HTML转换为DOCX文档
"""
import os
import threading
import time
from typing import Dict, Any, Tuple
from flask import Flask, request, jsonify, send_file, Response, make_response
import io
from dotenv import load_dotenv

from converters.converter_factory import ConverterFactory
from storage.local_storage import LocalStorage
from services.token_service import TokenService

# 加载环境变量
load_dotenv()

# 安全解析环境变量值
def safe_get_int(env_name, default_value):
    """安全地从环境变量获取整数值，处理可能包含注释的情况"""
    value = os.getenv(env_name, str(default_value))
    # 移除可能存在的注释部分
    value = value.split('#')[0].strip()
    try:
        return int(value)
    except ValueError:
        return default_value

# 配置
SECRET_KEY = os.getenv('SECRET_KEY', 'dev_secret_key')
STORAGE_PATH = os.getenv('STORAGE_PATH', 'tmp/storage')
BASE_URL = os.getenv('BASE_URL', 'http://localhost:5000')
CONVERTER_TYPE = os.getenv('CONVERTER_TYPE', 'docx')
LINK_EXPIRES_DEFAULT = safe_get_int('LINK_EXPIRES_DEFAULT', 300)
LINK_EXPIRES_MAX = safe_get_int('LINK_EXPIRES_MAX', 3600)
FILE_RETENTION = safe_get_int('FILE_RETENTION', 600)

# 初始化应用
app = Flask(__name__)
app.config['SECRET_KEY'] = SECRET_KEY

# 初始化服务
converter = ConverterFactory.get_converter(CONVERTER_TYPE)
storage = LocalStorage(STORAGE_PATH, FILE_RETENTION)
token_service = TokenService(SECRET_KEY, LINK_EXPIRES_DEFAULT, LINK_EXPIRES_MAX)

# 定期清理过期文件的后台任务
def cleanup_task() -> None:
    """定期清理过期文件的后台任务"""
    while True:
        try:
            count = storage.cleanup_expired()
            print(f"清理了 {count} 个过期文件")
        except Exception as e:
            print(f"清理文件时出错: {e}")
        time.sleep(60)  # 每分钟检查一次

# 启动清理线程
cleanup_thread = threading.Thread(target=cleanup_task, daemon=True)
cleanup_thread.start()

@app.route('/api/convert', methods=['POST'])
def convert_html() -> tuple[Response, int] | Response:
    """
    HTML转换为DOCX API接口
    
    接收JSON格式的请求，包含HTML内容和可选的链接有效期
    返回包含下载链接的JSON响应
    """
    # 验证请求格式
    if not request.is_json:
        return jsonify({'error': '请求必须是JSON格式'}), 400
    
    data = request.json
    html_content = data.get('html')
    expires_in = data.get('expires_in', LINK_EXPIRES_DEFAULT)
    
    # 验证参数
    if not html_content:
        return jsonify({'error': 'HTML内容不能为空'}), 400
    
    try:
        # 转换HTML为DOCX
        docx_content = converter.convert(html_content)
        
        # 保存文件
        file_id = storage.save(docx_content)
        
        # 生成临时下载链接
        token = token_service.generate_short_token(file_id, expires_in)  # 使用短链接
        download_url = f"{BASE_URL}/d/{token}"
        
        return jsonify({
            'download_url': download_url,
            'expires_in': expires_in
        })
    except Exception as e:
        return jsonify({'error': f'转换失败: {str(e)}'}), 500

@app.route('/download/<token>', methods=['GET'])
def download_file(token: str) -> Response:
    """
    文件下载接口
    
    验证令牌并返回对应的DOCX文件
    """
    # 验证下载令牌
    file_id = token_service.verify_download_token(token)
    if not file_id:
        return jsonify({'error': '无效或已过期的下载链接'}), 403
    
    # 获取文件内容
    file_content = storage.get(file_id)
    if not file_content:
        return jsonify({'error': '文件不存在或已被删除'}), 404
    
    # 将文件内容转换为文件对象
    file_obj = io.BytesIO(file_content)
    
    # 为内嵌场景创建响应
    response = make_response(file_content)
    response.headers['Content-Type'] = 'application/octet-stream'
    response.headers['Content-Disposition'] = f'inline; filename="{file_id}.docx"'
    
    return response

@app.route('/d/<token>', methods=['GET'])
def download_short_link(token: str) -> Response:
    """
    短链接文件下载接口
    
    验证令牌并返回对应的DOCX文件
    """
    # 验证短链接令牌并获取文件ID
    file_id = token_service.verify_short_token(token)
    if not file_id:
        return jsonify({'error': '无效或已过期的下载链接'}), 403
    
    # 获取文件内容
    file_content = storage.get(file_id)
    if not file_content:
        return jsonify({'error': '文件不存在或已被删除'}), 404
    
    # 为内嵌场景创建响应
    response = make_response(file_content)
    response.headers['Content-Type'] = 'application/octet-stream'
    response.headers['Content-Disposition'] = f'inline; filename="{file_id}.docx"'
    
    return response

@app.route('/health', methods=['GET'])
def health_check() -> Response:
    """健康检查接口"""
    return jsonify({'status': 'ok', 'converter_type': CONVERTER_TYPE})

if __name__ == '__main__':
    # 确保存储目录存在
    os.makedirs(STORAGE_PATH, exist_ok=True)
    # 启动应用
    app.run(debug=True) 