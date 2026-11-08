#!/usr/bin/env python3
"""把已有 PPT 的每一页背景全部替换为指定背景图。

- 默认先删除每页原有的大面积背景图片（覆盖 ≥70% 页面的图片），再铺新背景；
  原背景在母版里时新背景会盖在上面，无需处理母版。
- 背景按 cover 模式铺满（等比缩放居中裁边，logo 不变形）。

用法:
    python apply_background.py 输入.pptx -b 背景.png -o 输出.pptx [--keep-old-bg] [--stretch]

依赖: pip install python-pptx
"""
import argparse
import copy

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu


def iter_pictures(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_pictures(sh.shapes)
        elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield sh


def apply_bg(prs, bg_path, keep_old=False, stretch=False):
    sw, sh = prs.slide_width, prs.slide_height
    removed = 0
    for slide in prs.slides:
        if not keep_old:
            for pic in list(iter_pictures(slide.shapes)):
                area = pic.width * pic.height
                if area >= 0.7 * sw * sh:
                    pic._element.getparent().remove(pic._element)
                    removed += 1

        new_pic = slide.shapes.add_picture(bg_path, 0, 0)
        iw, ih = new_pic.image.size
        img_ar, slide_ar = iw / ih, sw / sh
        if stretch:
            new_pic.width, new_pic.height, new_pic.left, new_pic.top = sw, sh, 0, 0
        else:
            if img_ar >= slide_ar:
                w, h = int(sh * img_ar), sh
            else:
                w, h = sw, int(sw / img_ar)
            new_pic.width, new_pic.height = Emu(w), Emu(h)
            new_pic.left, new_pic.top = Emu((sw - w) // 2), Emu((sh - h) // 2)
        # 置于最底层
        sp_tree = slide.shapes._spTree
        sp_tree.remove(new_pic._element)
        sp_tree.insert(2, new_pic._element)
    return removed


def main():
    ap = argparse.ArgumentParser(description="为 PPT 每一页替换背景图")
    ap.add_argument("input", help="输入 pptx")
    ap.add_argument("-b", "--background", required=True, help="背景图路径")
    ap.add_argument("-o", "--output", required=True, help="输出 pptx 路径")
    ap.add_argument("--keep-old-bg", action="store_true", help="保留原有背景图片（默认删除大面积旧背景）")
    ap.add_argument("--stretch", action="store_true", help="拉伸铺满（默认等比 cover 裁边，不变形）")
    args = ap.parse_args()

    prs = Presentation(args.input)
    removed = apply_bg(prs, args.background, keep_old=args.keep_old_bg, stretch=args.stretch)
    prs.save(args.output)
    print(f"已生成: {args.output}（共 {len(prs.slides)} 页，删除旧背景图 {removed} 张）")
    print("注意: 若原背景设置在母版中，已被新背景覆盖；请回读或打开检查文字是否清晰。")


if __name__ == "__main__":
    main()
