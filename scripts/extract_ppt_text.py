#!/usr/bin/env python3
"""提取 PPT 每一页的全部文字，用于提炼知识点和生成后核对。

用法:
    python extract_ppt_text.py 课件1.pptx [课件2.pptx ...]

依赖: pip install python-pptx
"""
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def iter_shapes(shapes):
    """递归展开组合形状。"""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def slide_texts(slide):
    lines = []
    for sh in iter_shapes(slide.shapes):
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if t:
                lines.append(t)
        elif getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                cells = [c.text.strip() for c in row.cells]
                lines.append(" | ".join(cells))
    return lines


def main():
    if len(sys.argv) < 2:
        print("用法: python extract_ppt_text.py <文件.pptx> [更多.pptx ...]")
        sys.exit(1)
    for path in sys.argv[1:]:
        prs = Presentation(path)
        print(f"\n===== {path}（共 {len(prs.slides)} 页）=====")
        empty = []
        for i, slide in enumerate(prs.slides, 1):
            lines = slide_texts(slide)
            if lines:
                print(f"\n--- 第 {i} 页 ---")
                print("\n".join(lines))
            else:
                empty.append(i)
        if empty:
            print(f"\n（无文字页：{', '.join(map(str, empty))}）")


if __name__ == "__main__":
    main()
