# 文档阅读器 MCP 服务器

一个支持多种文档格式的 Model Context Protocol (MCP) 服务器，可以读取 Word、PDF、文本文件、RTF 等格式，并提供图片提取与分析功能。

## 🚀 快速开始

### 1. 安装依赖

```bash
# 方式1：使用安装脚本（推荐）
python install_deps.py

# 方式2：手动安装
pip install -r requirements.txt
```

### 2. 启动服务器

```bash
python server.py
```

### 3. 配置 Kiro IDE

在 Kiro IDE 中创建或编辑 `.kiro/settings/mcp.json` 文件：

**方式1：简化配置（推荐）**
```json
{
  "mcpServers": {
    "document-reader": {
      "command": "python",
      "args": ["server.py"],
      "cwd": "D:\\your-path\\document-reader-mcp",
      "env": {
        "PYTHONIOENCODING": "utf-8"
      },
      "disabled": false,
      "autoApprove": [
        "list_supported_formats",
        "get_document_info",
        "read_document",
        "read_document_with_media"
      ]
    }
  }
}
```

**方式2：完整路径配置**
```json
{
  "mcpServers": {
    "document-reader": {
      "command": "python",
      "args": ["D:\\devolp\\code1\\document-reader-mcp-master\\document-reader-mcp\\server.py"],
      "cwd": "D:\\devolp\\code1\\document-reader-mcp-master\\document-reader-mcp",
      "env": {
        "PYTHONPATH": "D:\\devolp\\code1\\document-reader-mcp-master\\document-reader-mcp",
        "PYTHONIOENCODING": "utf-8"
      },
      "disabled": false,
      "autoApprove": [
        "list_supported_formats",
        "get_document_info",
        "read_document",
        "read_document_with_media"
      ]
    }
  }
}
```

**配置说明：**
- `cwd`: 项目根目录路径
- `PYTHONPATH`: 确保模块导入正常（方式2需要）
- `PYTHONIOENCODING`: 确保中文字符正确显示
- `autoApprove`: 自动批准的安全工具列表，减少确认步骤

### 4. 测试功能

```bash
# 测试图表分析
python simple_diagram_reader.py

# 测试核心功能
python test_core_features.py
```

## ✨ 核心功能

- **多格式文档读取**: 支持 Word (.docx)、PDF、Excel (.xlsx/.xls)、文本文件、RTF 等格式
- **图片提取与分析**: 自动提取文档中的图片并进行结构分析
- **图表内容理解**: 基于 OpenCV 分析流程图、架构图等技术图表
- **媒体信息提取**: 提取文档中的图片和链接信息
- **链接验证**: 自动检查文档中链接的有效性
- **页面范围选择**: PDF 文档支持指定页面范围读取

## 🛠️ MCP 工具

### `read_document_
增强阅读文档，同时提取文字和图片内容
- `file_path`: 文
息

### `extract_documen`
提取文档中的图片和链接
- `file_path`: 文档路径
- `save_images`: 是否保存图片到本地

### `
无需复杂配置
- `image_path`: 图片路径

## 📊 图表分析能力

- **结构识别**: 检测矩形、圆形、线条等形元素
- *构图、网络图等
特征
- **技术理解**: 基

## 💡 使用场景

### 增强文档阅读
```
cx
```
自动提取文字内容和所有图片，并进行结构分析。

### 开发时
当 AI-Agent 需要理解业
```
 文件夹
analyze_s")


结构

```
├── server.py  主程序
析核心模块
├── ins 依赖安装脚本
├hon 依赖列表
└── extracted_ima
```

## 🔧 依赖说明

**核心依赖**:
- `mcp`: MCP 协议框架
- `pytho
- `PyPDF2`: PDF 文档处理
- `opencv-python`: 图像分析
- `numpy`: 数值计算

**可选依赖*
- `st
- `

## 设计理念

这个项目专注于**实用性**和***：
 OCR 配置
- 基于图像结构分析理解图表
-传
- 为 AI-Agent 开发提供

## 📄 许可证

MIT License: {
        "FASTMCP_LOG_LEVEL": "ERROR"
      },
      "disabled": false,
      "autoApprove": ["list_supported_formats"]
    }
  }
}
```

### 配置说明

- `autoApprove`: 自动批准的安全工具列表
- `disabled`: 设为 `false` 启用服务器
- `cwd`: 工作目录（本地安装时需要）
- `PYTHONIOENCODING`: 确保中文字符正确显示

## 使用方法

### 🚀 快速开始

1. **配置MCP服务器**（见上方配置说明）
2. **重启Kiro IDE** 或重新连接MCP服务器
3. **开始使用工具**

### 📖 基础文档阅读

#### 读取Word文档
```json
{
  "tool": "read_document",
  "arguments": {
    "file_path": "report.docx"
  }
}
```

#### 读取PDF特定页面
```json
{
  "tool": "read_document", 
  "arguments": {
    "file_path": "document.pdf",
    "page_range": "1-5"
  }
}
```

#### 读取Excel文档
```json
{
  "tool": "read_document",
  "arguments": {
    "file_path": "data.xlsx"
  }
}
```

#### 读取Excel特定工作表
```json
{
  "tool": "read_document",
  "arguments": {
    "file_path": "data.xlsx",
    "sheet_name": "Sheet1"
  }
}
```

#### 获取文档信息
```json
{
  "tool": "get_document_info",
  "arguments": {
    "file_path": "example.docx"
  }
}
```

### 🖼️ 图片自动解析功能

#### 方式1：增强文档阅读（推荐）
```json
{
  "tool": "read_document_with_media",
  "arguments": {
    "file_path": "document.docx",
    "include_media_info": true
  }
}
```
**特点**：
- ✅ 自动解析图片和链接
- 📊 在文档内容后显示媒体信息
- 🎯 适合完整文档分析

#### 方式2：专门媒体提取
```json
{
  "tool": "extract_document_media",
  "arguments": {
    "file_path": "document.docx",
    "extract_images": true,
    "extract_links": true,
    "save_images": false
  }
}
```
**特点**：
- 🎯 专门用于媒体分析
- 💾 可选择保存图片到本地
- 📈 提供详细统计信息

#### 方式3：普通阅读（不解析图片）
```json
{
  "tool": "read_document",
  "arguments": {
    "file_path": "document.docx"
  }
}
```
**特点**：
- 🚀 速度最快
- 📝 只获取文本内容
- 🎯 适合纯文本需求

### 🔍 图片解析能力展示

当使用图片解析功能时，你会得到：

```
文档媒体信息 (example.docx):

摘要:
- 图片总数: 2
- 链接总数: 3
- 图片处理错误: 0
- 链接处理错误: 0

图片信息:
1. image1.png
   - 格式: PNG
   - 尺寸: (670, 346)
   - 大小: 174900 字节
   - 流程图分析:
     - 文本: 检测到的文字内容
     - 节点数: 5
     - 连线数: 4

2. chart.jpg
   - 格式: JPEG
   - 尺寸: (800, 600)
   - 大小: 245600 字节

链接信息:
1. https://www.example.com
   - 域名: www.example.com
   - 协议: https
   - 状态: 可访问 (HTTP 200)
```

### 🛠️ 实用工具

#### 查看支持格式
```json
{
  "tool": "list_supported_formats",
  "arguments": {}
}
```

#### 批量处理示例
```python
# 在Kiro中，你可以这样批量处理文档
documents = ["doc1.docx", "doc2.pdf", "doc3.txt"]

for doc in documents:
    # 使用增强阅读获取完整信息
    result = mcp_call("read_document_with_media", {
        "file_path": doc,
        "include_media_info": True
    })
    print(f"处理完成: {doc}")
```

## 页面范围格式

PDF文档支持灵活的页面范围选择：

- `"all"` - 所有页面（默认）
- `"1-5"` - 第1到5页
- `"1,3,5"` - 第1、3、5页
- `"1-3,7,10-12"` - 第1-3页、第7页、第10-12页

## 错误处理

- 自动检测文本文件编码 (UTF-8, GBK, GB2312, Latin-1)
- 优雅处理缺失的依赖库
- 详细的错误信息和建议

## 媒体处理功能 🆕

### 图片处理
- **支持格式**: Word文档中的嵌入图片
- **提取信息**: 文件名、格式、尺寸、大小
- **保存功能**: 可选择将图片保存到本地
- **错误处理**: 优雅处理损坏或不支持的图片

### 链接处理
- **支持格式**: Word、PDF、文本文件中的HTTP/HTTPS链接
- **提取信息**: URL、域名、协议
- **有效性验证**: 自动检查链接是否可访问
- **状态码**: 显示HTTP响应状态码

### 使用场景
- 文档内容分析和审计
- 媒体资源清单生成
- 链接有效性批量检查
- 文档迁移前的资源盘点

## 在 Kiro IDE 中使用

### 🎯 配置步骤

1. **打开Kiro IDE**
2. **创建MCP配置文件**：
   - 路径：`.kiro/settings/mcp.json`
   - 如果文件不存在，创建一个新文件

3. **添加配置内容**：
```json
{
  "mcpServers": {
    "document-reader": {
      "command": "python",
      "args": ["server.py"],
      "cwd": "D:\\your-path\\document-reader-mcp",
      "env": {
        "PYTHONIOENCODING": "utf-8"
      },
      "disabled": false,
      "autoApprove": [
        "list_supported_formats",
        "get_document_info"
      ]
    }
  }
}
```

4. **重启Kiro** 或使用命令面板搜索 "MCP" 重新连接服务器

### 💡 使用技巧

#### 在聊天中使用
```
请帮我读取这个Word文档的内容，并分析其中的图片信息：
文件路径：./reports/monthly-report.docx
```

Kiro会自动调用 `read_document_with_media` 工具。

#### 批量文档分析
```
请分析这个文件夹中所有Word文档的媒体信息：
- ./docs/report1.docx
- ./docs/report2.docx  
- ./docs/presentation.docx
```

#### 文档格式转换准备
```
我需要将这些PDF文档转换为Markdown，
请先帮我分析文档结构和媒体内容：
./pdfs/technical-guide.pdf
```

### 🔧 故障排除

#### 常见问题

1. **MCP服务器连接失败**
   - 检查 `cwd` 路径是否正确
   - 确认 `server.py` 文件存在
   - 检查Python环境是否正确

2. **中文字符显示异常**
   - 确保配置中包含 `"PYTHONIOENCODING": "utf-8"`

3. **图片解析失败**
   - 检查是否安装了 `Pillow` 库：`pip install Pillow`
   - 确认Word文档中确实包含图片

4. **链接验证不工作**
   - 安装 `requests` 库：`pip install requests`
   - 检查网络连接

#### 调试方法

1. **查看MCP服务器状态**
   - 在Kiro中打开命令面板
   - 搜索 "MCP Server" 查看连接状态

2. **测试基础功能**
```json
{
  "tool": "list_supported_formats",
  "arguments": {}
}
```

3. **检查依赖库**
```json
{
  "tool": "list_supported_formats", 
  "arguments": {}
}
```
返回结果会显示各个依赖库的安装状态。

## 性能与最佳实践

### ⚡ 性能优化

#### 选择合适的工具
- **纯文本需求**：使用 `read_document`（最快）
- **需要媒体信息**：使用 `read_document_with_media`
- **专门媒体分析**：使用 `extract_document_media`

#### 大文件处理
```json
{
  "tool": "read_document",
  "arguments": {
    "file_path": "large-document.pdf",
    "page_range": "1-10"
  }
}
```
对于大型PDF，使用页面范围限制可以显著提升性能。

#### 批量处理建议
- 避免同时处理过多大文件
- 优先处理小文件，再处理大文件
- 使用 `get_document_info` 先了解文件大小

### 🎯 最佳实践

#### 1. 文档分析工作流
```
1. get_document_info - 了解文档基本信息
2. read_document - 快速获取文本内容  
3. read_document_with_media - 深入分析（如需要）
```

#### 2. 媒体资源管理
```json
{
  "tool": "extract_document_media",
  "arguments": {
    "file_path": "document.docx",
    "extract_images": true,
    "extract_links": true,
    "save_images": true
  }
}
```
设置 `save_images: true` 可以将图片保存到本地，便于后续处理。

#### 3. 错误处理
- 始终检查返回结果中的错误信息
- 对于批量处理，建议逐个处理并记录失败的文件
- 使用 `list_supported_formats` 确认依赖库状态

### 📊 性能数据

| 操作 | 平均耗时 | 内存使用 |
|------|----------|----------|
| 读取文本文件 (1MB) | < 0.1秒 | 低 |
| 读取Word文档 (5MB) | < 0.5秒 | 中等 |
| 读取PDF文档 (10MB) | < 1秒 | 中等 |
| 图片提取 (含10张图) | < 2秒 | 较高 |
| 链接验证 (10个链接) | 2-5秒 | 低 |

## 依赖库

### 核心依赖
- `mcp` - MCP协议支持
- `python-docx` - Word文档处理
- `PyPDF2` - PDF文档处理
- `striprtf` - RTF文档处理

### 媒体处理依赖 🆕
- `Pillow` - 图片处理和分析
- `requests` - 链接验证（可选）

### 安装命令
```bash
# 安装所有依赖
pip install -r requirements.txt

# 或单独安装
pip install mcp python-docx PyPDF2 striprtf Pillow requests openpyxl pandas
```

## 快速参考

### 🔧 工具对比表

| 工具名称 | 读取文本 | 解析图片 | 提取链接 | 性能 | 适用场景 |
|----------|----------|----------|----------|------|----------|
| `read_document` | ✅ | ❌ | ❌ | 🚀🚀🚀 | 快速文本阅读 |
| `read_document_with_media` | ✅ | ✅ | ✅ | 🚀🚀 | 完整文档分析 |
| `extract_document_media` | ❌ | ✅ | ✅ | 🚀 | 专门媒体提取 |
| `get_document_info` | ❌ | ❌ | ❌ | 🚀🚀🚀 | 文档信息查看 |
| `list_supported_formats` | ❌ | ❌ | ❌ | 🚀🚀🚀 | 功能状态检查 |

### 📋 支持格式一览

| 格式 | 扩展名 | 文本读取 | 图片提取 | 链接提取 | 页面范围 |
|------|--------|----------|----------|----------|----------|
| Word文档 | .docx | ✅ | ✅ | ✅ | ❌ |
| PDF文档 | .pdf | ✅ | ❌ | ✅ | ✅ |
| Excel文档 | .xlsx, .xls | ✅ | ✅ | ✅ | ❌ |
| 纯文本 | .txt, .md | ✅ | ❌ | ✅ | ❌ |
| RTF文档 | .rtf | ✅ | ❌ | ❌ | ❌ |
| 代码文件 | .py, .js, .html, .css | ✅ | ❌ | ✅ | ❌ |

### 🎯 使用场景推荐

| 场景 | 推荐工具 | 配置建议 |
|------|----------|----------|
| 快速浏览文档内容 | `read_document` | 无特殊配置 |
| 文档内容+媒体分析 | `read_document_with_media` | `include_media_info: true` |
| 媒体资源清单 | `extract_document_media` | `save_images: true` |
| 链接有效性检查 | `extract_document_media` | 安装 `requests` 库 |
| 大文件处理 | `read_document` | 使用 `page_range` |
| 批量文档处理 | 组合使用 | 先用 `get_document_info` |

## 更新日志

### v2.1.0 🆕
- ✅ 新增Excel文档支持 (.xlsx/.xls)
- ✅ Excel图片提取功能
- ✅ Excel链接提取和验证功能
- ✅ Excel工作表信息获取
- ✅ 支持指定工作表读取

### v2.0.0
- ✅ 新增图片自动解析功能
- ✅ 新增链接提取和验证功能
- ✅ 新增 `read_document_with_media` 工具
- ✅ 新增 `extract_document_media` 工具
- ✅ 支持图片保存到本地
- ✅ 支持流程图基础分析
- ✅ 完善错误处理机制

### v1.0.0
- ✅ 基础文档读取功能
- ✅ 多格式支持
- ✅ PDF页面范围选择
- ✅ 文档信息获取

## 许可证

MIT License