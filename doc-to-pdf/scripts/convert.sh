#!/bin/bash
# Doc to PDF Converter Wrapper

INPUT_FILE="$1"
OUTPUT_DIR="$2"

if [ -z "$INPUT_FILE" ]; then
    echo "用法: ./convert.sh <input_file> [output_directory]"
    echo "示例: ./convert.sh /path/to/document.md"
    echo "       ./convert.sh report.docx /output/path"
    exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python3 "$SCRIPT_DIR/convert.py" "$INPUT_FILE" "$OUTPUT_DIR"
