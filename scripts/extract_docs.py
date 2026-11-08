#!/usr/bin/env python3
"""提取多种格式文件的文字内容，供备课提炼使用。

支持:
    .docx  纯标准库解析（无需第三方依赖）
    .pdf   优先 pypdf，其次 PyPDF2（缺失时提示安装）
    .txt / .md 直接读取
    .doc   旧格式无法解析，提示转存为 .docx
    .pptx  请用 extract_ppt_text.py
    图片   文字/题目需要视觉识别，由模型直接看图转写，不在本脚本范围

用法:
    python extract_docs.py 讲义.docx 要求.pdf 笔记.md

依赖: .docx/.txt/.md 无需依赖；.pdf 需要 pip install pypdf
"""
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import zipfile
from pathlib import Path
from xml.etree import ElementTree

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def extract_docx(path):
    """纯标准库提取 docx 正文（含表格），按段落输出。"""
    out = []
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = ElementTree.fromstring(xml)
    body = root.find(f"{W_NS}body")

    def para_text(p):
        return "".join(t.text or "" for t in p.iter(f"{W_NS}t")).strip()

    def walk(el):
        for child in el:
            tag = child.tag
            if tag == f"{W_NS}p":
                t = para_text(child)
                if t:
                    out.append(t)
            elif tag == f"{W_NS}tbl":
                for tr in child.iter(f"{W_NS}tr"):
                    cells = [para_text(tc) or " " for tc in tr.iter(f"{W_NS}tc")]
                    out.append(" | ".join(cells))
            else:
                walk(child)

    walk(body)
    return out


def extract_pdf(path):
    try:
        try:
            from pypdf import PdfReader
        except ImportError:
            from PyPDF2 import PdfReader
    except ImportError:
        return None  # 依赖缺失
    reader = PdfReader(path)
    out = []
    for i, page in enumerate(reader.pages, 1):
        text = (page.extract_text() or "").strip()
        out.append(f"--- 第 {i} 页 ---\n{text}" if text else f"--- 第 {i} 页 ---（无文字，可能是扫描件，需要视觉识别）")
    return out


def extract_lines(path):
    return Path(path).read_text(encoding="utf-8", errors="replace").splitlines()


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for arg in sys.argv[1:]:
        p = Path(arg)
        ext = p.suffix.lower()
        print(f"\n===== {p.name} =====")
        if not p.exists():
            print("（文件不存在，请确认路径）")
            continue
        if ext == ".docx":
            lines = extract_docx(str(p))
            print("\n".join(lines) if lines else "（未提取到文字）")
        elif ext == ".pdf":
            lines = extract_pdf(str(p))
            if lines is None:
                print("（缺少依赖，请先运行: pip install pypdf）")
            else:
                print("\n".join(lines))
        elif ext in (".txt", ".md", ".markdown"):
            print("\n".join(extract_lines(str(p))))
        elif ext == ".doc":
            print("（旧版 .doc 无法直接解析，请先另存为 .docx 再提取）")
        elif ext == ".pptx":
            print("（PPT 请用: python extract_ppt_text.py）")
        elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"):
            print("（图片内容由模型直接看图识别：若是题目/笔记请转写进大纲）")
        else:
            print(f"（暂不支持 {ext} 格式，请让用户提供可读格式）")


if __name__ == "__main__":
    main()
