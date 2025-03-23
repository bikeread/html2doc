# HTML2DOC 转换服务

轻量级HTML转DOCX文档转换服务，提供REST API接口，支持临时文件下载链接生成。

## 功能特点

- 接收JSON格式的HTML内容并转换为DOCX文档
- 生成临时下载链接（默认5分钟有效期）
- 自动清理过期文件
- 使用python-docx + BeautifulSoup实现HTML到DOCX的转换
- 轻量级部署，无需外部依赖

## 架构设计

本项目采用**适配器+门面模式**设计，便于扩展不同的转换器实现和存储方案：

1. **转换器层**：负责将HTML转换为DOCX
   - 抽象基类 `BaseConverter`
   - 实现类 `DocxConverter` 
   - 工厂类 `ConverterFactory` 创建具体转换器实例
x
2. **存储层**：负责文件存储与管理
   - 抽象基类 `BaseStorage`
   - 实现类 `LocalStorage`（可扩展其他存储如S3等）

3. **服务层**：对外提供REST API接口
   - Flask应用提供HTTP接口
   - 令牌服务处理临时链接生成与验证

## 安装与配置

### 依赖安装

```bash
pip install -r requirements.txt
```

### 环境变量配置

创建`.env`文件或设置以下环境变量：

```
SECRET_KEY=your_secret_key
STORAGE_PATH=tmp/storage
BASE_URL=http://localhost:5000
CONVERTER_TYPE=docx
LINK_EXPIRES_DEFAULT=300  # 5分钟
LINK_EXPIRES_MAX=3600  # 最长1小时
FILE_RETENTION=600  # 文件保留10分钟
```

## 启动服务

```bash
python app.py
```

生产环境部署建议使用WSGI服务器：

```bash
gunicorn app:app
```

## API使用说明

### 转换HTML为DOCX

**请求**:

```
POST /api/convert
Content-Type: application/json

{
    "html": "<h1>文档标题</h1><p>这是一段文本内容</p>",
    "expires_in": 300  # 可选，链接有效期（秒）
}
```

**响应**:

```json
{
    "download_url": "http://localhost:5000/download/eyJ0eXAi...",
    "expires_in": 300
}
```

### 下载文件

使用生成的链接即可下载文件：

```
GET /download/eyJ0eXAi...
```

### 健康检查

```
GET /health
```

## 扩展指南

### 添加新的转换器

1. 在 `converters` 目录下创建新的转换器类，继承 `BaseConverter`
2. 实现 `convert` 方法
3. 在 `ConverterFactory` 中添加新的转换器类型支持

### 添加新的存储方式

1. 在 `storage` 目录下创建新的存储类，继承 `BaseStorage`
2. 实现所有抽象方法
3. 在 `app.py` 中使用新的存储类

## 注意事项

- 默认下载链接有效期为5分钟，最长不超过1小时
- 文件默认保留10分钟后自动清理
- 在生产环境中请设置强密钥和适当的BASE_URL 