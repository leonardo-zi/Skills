#!/usr/bin/env python3
"""
Doc to PDF Converter
Supported formats: .md, .txt, .docx
Uses playwright + pandoc for PDF generation (supports Chinese + images)
"""

import sys
import os
import subprocess
import tempfile
import re


def check_dependencies():
    """检查必要的依赖"""
    missing = []

    try:
        subprocess.run(["pandoc", "--version"], capture_output=True, check=True)
    except:
        missing.append("pandoc")

    try:
        from playwright.sync_api import sync_playwright
    except:
        missing.append("playwright")

    return missing


def install_dependencies():
    """检查依赖"""
    print("正在检查依赖...")

    missing = check_dependencies()
    if not missing:
        print("✓ 所有依赖已安装")
        return True

    print(f"✗ 缺少依赖: {', '.join(missing)}")
    if "playwright" in missing:
        print("请运行: pip install playwright && playwright install chromium")
    return False


def convert_to_html(input_path):
    """使用 pandoc 转换为 HTML，支持 docx 图片"""

    ext = os.path.splitext(input_path)[1].lower()

    # 如果是 docx，先提取图片
    if ext == ".docx":
        return convert_docx_with_images(input_path)

    # txt 文件使用智能标题识别
    if ext == ".txt":
        return convert_txt_with_heading_detection(input_path)

    # md 文件直接转换
    cmd = [
        "pandoc",
        input_path,
        "-o",
        "-",
        "--wrap=none",
        "--standalone",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    return result.stdout


def convert_txt_with_heading_detection(txt_path):
    """转换纯文本文件，智能识别标题和正文"""

    with open(txt_path, "r", encoding="utf-8") as f:
        content = f.read()

    lines = content.split("\n")
    html_parts = []

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # 跳过空行
        if not line:
            i += 1
            continue

        # 检测标题样式1: ======= 包围 (一级标题)
        if line.startswith("=") and "=" in line and line.strip("=").strip():
            # 上一行是标题内容
            if i > 0 and lines[i - 1].strip():
                title = lines[i - 1].strip()
                html_parts.append(f"<h1>{title}</h1>\n")
            i += 1
            continue

        # 检测标题样式2: ------- 包围 (二级标题)
        if line.startswith("-") and "-" in line and line.strip("-").strip():
            if i > 0 and lines[i - 1].strip():
                title = lines[i - 1].strip()
                html_parts.append(f"<h2>{title}</h2>\n")
            i += 1
            continue

        # 检测纯分隔线
        if re.match(r"^[=\-]{4,}$", line):
            i += 1
            continue

        # 检测标题样式3: 全大写短行 (可能是标题)
        # 条件：全大写，长度 3-80 字符，且不是纯数字
        if (
            line.isupper()
            and len(line) >= 3
            and len(line) <= 80
            and not line.isdigit()
            and not re.match(r"^[0-9\.\,\-\s]+$", line)
        ):
            html_parts.append(f"<h2>{line}</h2>\n")
            i += 1
            continue

        # 检测标题样式4: 首字母大写短行，前面有空格缩进
        # 条件：首字母大写，其他小写，长度 3-60，且前面有足够的空白行
        if (
            line[0].isupper()
            and len(line) >= 3
            and len(line) <= 60
            and not line.isdigit()
        ):
            # 检查前面是否有空行
            if i == 0 or (i > 0 and not lines[i - 1].strip()):
                html_parts.append(f"<h3>{line}</h3>\n")
                i += 1
                continue

        # 检测标题样式5: 数字序号标题 (如 "1. 第一章" 或 "1.1 概述")
        if re.match(r"^\d+(\.\d+)*\.?\s+.{2,60}$", line):
            if re.match(r"^\d+(\.\d+)*\.\s+", line):
                level = line.count(".")
                if level == 1:
                    html_parts.append(f"<h2>{line}</h2>\n")
                else:
                    html_parts.append(f"<h3>{line}</h3>\n")
                i += 1
                continue

        # 检测标题样式6: 带中文标题符号 (#、※、◆)
        if re.match(r"^[#※◆●■★☆]+\s*.{1,50}$", line):
            title = re.sub(r"^[#※◆●■★☆]+\s*", "", line)
            html_parts.append(f"<h3>{title}</h3>\n")
            i += 1
            continue

        # 默认作为正文处理
        # 合并连续的非标题行
        para_lines = [line]
        j = i + 1
        while j < len(lines):
            next_line = lines[j].strip()
            if not next_line:
                break
            # 如果遇到可能的标题，停止
            if (
                (next_line.isupper() and len(next_line) >= 3 and len(next_line) <= 80)
                or re.match(r"^\d+(\.\d+)*\.?\s+", next_line)
                or re.match(r"^[#※◆●■★☆]+\s*.{1,50}$", next_line)
            ):
                break
            # 检查是否是分隔线
            if re.match(r"^[=\-]{4,}$", next_line):
                break
            para_lines.append(next_line)
            j += 1

        if para_lines:
            # 检查是否是列表项
            if re.match(r"^[\-\*\•]\s+", para_lines[0]):
                for p in para_lines:
                    text = re.sub(r"^[\-\*\•]\s+", "", p)
                    html_parts.append(f"<li>{text}</li>\n")
                html_parts.append("</ul>\n")
            elif re.match(r"^\d+\.\s+", para_lines[0]):
                for p in para_lines:
                    text = re.sub(r"^\d+\.\s+", "", p)
                    html_parts.append(f"<li>{text}</li>\n")
                html_parts.append("</ol>\n")
            else:
                # 合并为段落
                text = " ".join(para_lines)
                html_parts.append(f"<p>{text}</p>\n")

        i = j

    return "\n".join(html_parts)


def convert_docx_with_images(docx_path):
    """转换 docx 并提取图片"""
    from docx import Document
    import base64
    import shutil

    doc = Document(docx_path)

    # 提取所有图片为 base64
    images = {}  # rel_id -> base64_data

    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            try:
                image_part = doc.part.related_parts[rel_id]
                image_data = image_part.blob
                content_type = image_part.content_type

                # 获取文件扩展名
                ext = content_type.split("/")[-1]
                if ext == "jpeg":
                    ext = "jpg"

                # 转为 base64
                b64 = base64.b64encode(image_data).decode()
                images[rel_id] = {
                    "data": f"data:image/{ext};base64,{b64}",
                    "ext": ext,
                }
                print(f"  提取图片: {rel.target_ref}")
            except Exception as e:
                print(f"  无法提取图片: {e}")

    # 如果没有提取到图片，使用 pandoc 简单转换
    if not images:
        cmd = ["pandoc", docx_path, "-o", "-", "--wrap=none", "--standalone"]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        return result.stdout

    # 使用 pandoc 转换为 HTML
    cmd = ["pandoc", docx_path, "-o", "-", "--wrap=none", "--standalone"]
    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
    html_content = result.stdout

    # 替换图片占位符为实际 base64 数据
    for rel_id, img_data in images.items():
        # 查找对应的 media 文件名
        for rel in doc.part.rels.values():
            if rel.rId == rel_id:
                filename = rel.target_ref.split("/")[-1]
                # 替换 HTML 中的图片引用
                html_content = html_content.replace(
                    f"media/{filename}", img_data["data"]
                )
                # 也处理不带 media 前缀的
                html_content = html_content.replace(
                    f'"{filename}"', f'"{img_data["data"]}"'
                )
                break

    return html_content


def generate_html_with_styles(html_content):
    """生成带完整样式的 HTML"""
    css = """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        @page { size: A4; margin: 2cm; @bottom-right { content: counter(page); font-size: 10pt; color: #666; } }
        body { font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; font-size: 14px; line-height: 1.8; color: #333; padding: 20px; max-width: 100%; }
        h1 { font-size: 28px; color: #1a1a1a; border-bottom: 2px solid #333; padding-bottom: 12px; margin: 30px 0 20px 0; page-break-after: avoid; }
        h2 { font-size: 22px; color: #333; margin: 25px 0 15px 0; page-break-after: avoid; }
        h3 { font-size: 18px; color: #555; margin: 20px 0 10px 0; page-break-after: avoid; }
        h4, h5, h6 { font-size: 16px; color: #666; margin: 15px 0 10px 0; }
        p { margin: 12px 0; text-align: justify; word-wrap: break-word; }
        code { font-family: "SF Mono", Monaco, "Cascadia Code", Consolas, monospace; font-size: 13px; background-color: #f5f5f5; padding: 2px 6px; border-radius: 4px; color: #e83e8c; }
        pre { background-color: #f5f5f5; padding: 15px; border-radius: 6px; overflow-x: auto; font-size: 13px; margin: 15px 0; page-break-inside: avoid; }
        pre code { background: none; padding: 0; color: #333; }
        blockquote { border-left: 4px solid #ddd; margin: 15px 0; padding: 10px 20px; color: #666; background-color: #fafafa; font-style: italic; }
        img { max-width: 100%; max-height: 100%; width: auto; height: auto; margin: 15px 0; page-break-inside: avoid; object-fit: contain; display: block; border: none; outline: none; background: transparent; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; page-break-inside: avoid; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #f5f5f5; font-weight: bold; }
        tr:nth-child(even) { background-color: #fafafa; }
        ul, ol { margin: 12px 0; padding-left: 25px; }
        li { margin: 6px 0; }
        a { color: #0066cc; text-decoration: none; }
        a:hover { text-decoration: underline; }
        hr { border: none; border-top: 1px solid #ddd; margin: 20px 0; }
    """

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>{css}
    </style>
</head>
<body>
{html_content}
</body>
</html>
"""
    return html


def convert_file(input_path, output_dir=None):
    """主转换函数"""

    if not os.path.exists(input_path):
        print(f"✗ 文件不存在: {input_path}")
        return None

    if not install_dependencies():
        return None

    # 确定输出目录
    if output_dir is None:
        output_dir = os.path.dirname(input_path) or "."

    # 生成输出文件名
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.pdf")

    # 获取文件扩展名
    ext = os.path.splitext(input_path)[1].lower()

    if ext not in [".md", ".txt", ".docx"]:
        print(f"✗ 不支持的格式: {ext}")
        print("支持的格式: .md, .txt, .docx")
        return None

    print(f"转换: {input_path}")
    print(f"输出: {output_path}")

    try:
        # 转换为 HTML
        html_content = convert_to_html(input_path)

        # 添加样式
        full_html = generate_html_with_styles(html_content)

        # 写入临时 HTML 文件
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".html", delete=False, encoding="utf-8"
        ) as tmp:
            tmp.write(full_html)
            html_path = tmp.name

        try:
            # 使用 playwright 生成 PDF
            from playwright.sync_api import sync_playwright

            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()

                # 加载 HTML 文件
                page.goto(f"file://{html_path}", wait_until="networkidle")

                # 等待图片加载
                page.wait_for_timeout(2000)

                # 生成 PDF
                page.pdf(
                    path=output_path,
                    format="A4",
                    print_background=True,
                    margin={
                        "top": "2cm",
                        "bottom": "2cm",
                        "left": "2cm",
                        "right": "2cm",
                    },
                )

                browser.close()

            print(f"✓ 转换完成: {output_path}")
            return output_path

        finally:
            if os.path.exists(html_path):
                os.unlink(html_path)

    except Exception as e:
        import traceback

        print(f"✗ 转换失败: {e}")
        traceback.print_exc()
        return None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert.py <input_file> [output_directory]")
        print("示例: python convert.py /path/to/document.md")
        print("       python convert.py report.docx /output/path")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    result = convert_file(input_file, output_dir)
    if result:
        print(f"\n输出文件: {result}")
    else:
        sys.exit(1)
