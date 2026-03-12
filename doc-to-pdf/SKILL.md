---
name: doc-to-pdf
description: 将 md、txt、docx 等文档格式转换为专业的 PDF 文件。使用场景：(1) 用户说"转成 PDF"、"转为 pdf"、"convert to PDF" (2) 用户提供 .md/.txt/.docx 文件并要求转为 PDF (3) 用户说"帮我把 xxx 文件转成 pdf"。输出文件保存在用户指定的目录或当前工作目录。自动支持中文、多图文档，输出专业的 A4 排版效果。
---

# Doc to PDF Converter

将 Markdown、纯文本、Word 文档转换为专业的 PDF 文件。支持中文显示、图片嵌入，输出符合出版规范的视觉排版。

## 支持格式

| 格式 | 说明 | 输出样式 |
|------|------|----------|
| `.md` | Markdown 文件 | 技术文档风格，支持代码高亮、目录 |
| `.txt` | 纯文本文件 | 简洁阅读版，适合小说/文章 |
| `.docx` | Word 文档 | 保留格式，支持图片嵌入 |

## 使用方式

直接提供要转换的文件路径：

```
转换: /Users/xxx/document.md
转换: /path/to/report.docx
转换: notes.txt
```

示例对话：
- "把 readme.md 转成 PDF"
- "convert /Users/xxx/report.docx to PDF"
- "帮我把这个 markdown 文件转 pdf"
- "把这个 txt 文件转为 pdf"

## 输出说明

- **默认输出位置**：源文件同目录
- **文件名**：`原文件名.pdf`
- **页面尺寸**：A4 (210mm × 297mm)
- **页边距**：上下左右各 2cm
- **页码**：底部居中，自动编号

## PDF 排版规范

### 字体规范

| 元素 | 字体 | 大小 | 颜色 |
|------|------|------|------|
| 正文 | PingFang SC, Microsoft YaHei, Helvetica Neue | 14px | #333 |
| 一级标题 | PingFang SC, Microsoft YaHei | 28px | #1a1a1a |
| 二级标题 | PingFang SC, Microsoft YaHei | 22px | #333 |
| 三级标题 | PingFang SC, Microsoft YaHei | 18px | #555 |
| 代码 | SF Mono, Monaco, Consolas | 13px | #333 (背景 #f5f5f5) |
| 引用 | PingFang SC | 11px 斜体 | #666 |
| 链接 | - | - | #0066cc |

### 行距与间距

- **正文行高**：1.8 倍字体大小
- **段落间距**：12px
- **标题上间距**：25-30px
- **标题下间距**：15-20px
- **代码块内边距**：15px
- **列表缩进**：25px

### 标题样式

- **一级标题**：28px 粗体，底部 2px 实线分隔线 (#333)，上下边距较大
- **二级标题**：22px 粗体，无分隔线
- **三级标题**：18px 粗体
- 标题避免分页，控制在页面顶部

### 代码块

- 背景色：#f5f5f5
- 圆角：6px
- 内边距：15px
- 等宽字体
- 避免代码块内部分页

### 表格

- 表头背景：#f5f5f5
- 单元格边框：1px solid #ddd
- 内边距：12px
- 斑马纹：偶数行 #fafafa
- 避免表格内部分页

### 图片

- 最大宽度：100%
- 保持原始宽高比
- 上下边距：15px
- 避免图片分页
- 无边框、无背景

### 列表

- 无序列表：实心圆点符号
- 有序列表：数字序号
- 缩进：25px
- 列表项间距：6px

### 引用块

- 左边框：4px solid #ddd
- 背景色：#fafafa
- 内边距：10px 20px
- 斜体文字

## 转换流程

1. **读取源文件**：识别文件格式 (md/txt/docx)
2. **格式转换**：使用 pandoc 转为 HTML
3. **图片处理**（仅 docx）：
   - 提取文档中的图片
   - 转换为 base64 嵌入 HTML
4. **样式注入**：应用专业排版样式
5. **PDF 生成**：使用 Playwright 渲染为 PDF
6. **输出文件**：保存到指定目录

## 依赖说明

- **pandoc**：文档格式转换
- **playwright**：浏览器渲染引擎
- **python-docx**：Word 文档图片提取
- **Pillow**：图像处理

首次使用时会自动检查依赖，如缺少会有提示。

## 输出示例

输入：`/Users/xxx/document.md`
输出：`/Users/xxx/document.pdf`

页面效果：
- 封面：标题样式
- 目录：自动生成（如果源文档有目录）
- 正文：专业排版
- 页脚：页码
