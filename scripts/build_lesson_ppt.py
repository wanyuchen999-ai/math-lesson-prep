#!/usr/bin/env python3
"""从 Markdown 大纲生成"知识点罗列式"教学 PPT，每页自动铺满背景图。

大纲格式（UTF-8）:
    # PPT 主标题
    ## 第一个板块标题
    - 要点1
    - 要点2
    ![几何图形或题目截图](图片路径.png)   ← 图片会插入本页，自动缩放
    ## 第二个板块标题
    - 要点1

特点: 16:9、无动画、每页一个板块标题 + 要点列表，适合线上课投屏。
用法:
    python build_lesson_ppt.py outline.md -o 课件.pptx -b 背景.png

依赖: pip install python-pptx
"""
import argparse
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

FONT = "微软雅黑"
TITLE_COLOR = RGBColor(0x1F, 0x38, 0x64)   # 深蓝，与背景 logo 同色系
BODY_COLOR = RGBColor(0x33, 0x33, 0x33)

SLIDE_W = Emu(12192000)  # 13.333 in
SLIDE_H = Emu(6858000)   # 7.5 in

IMG_RE = re.compile(r"^!\[[^\]]*\]\((.+?)\)\s*$")


def set_run_font(run, size, bold=False, color=BODY_COLOR):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    # 同时设置东亚字体，保证中文按微软雅黑渲染
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", FONT)


def add_background(slide, bg_path):
    """背景图按 cover 模式铺满整页并置于最底层。"""
    pic = slide.shapes.add_picture(bg_path, 0, 0)
    iw, ih = pic.image.size
    img_ar, slide_ar = iw / ih, SLIDE_W / SLIDE_H
    if img_ar >= slide_ar:
        pic.height = SLIDE_H
        pic.width = Emu(int(SLIDE_H * img_ar))
    else:
        pic.width = SLIDE_W
        pic.height = Emu(int(SLIDE_W / img_ar))
    pic.left = Emu(int((SLIDE_W - pic.width) / 2))
    pic.top = Emu(int((SLIDE_H - pic.height) / 2))
    sp_tree = slide.shapes._spTree
    sp_tree.remove(pic._element)
    sp_tree.insert(2, pic._element)  # 2 = 形状层最底


def parse_outline(path):
    """解析 Markdown 大纲 → (主标题, [(板块标题, [("text",行) | ("image",路径), ...]), ...])"""
    title = None
    sections = []
    cur = None
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.rstrip()
            if line.startswith("# "):
                if title is None:
                    title = line[2:].strip()
                else:
                    sections.append((line[2:].strip(), []))
                    cur = sections[-1][1]
            elif line.startswith("## "):
                sections.append((line[3:].strip(), []))
                cur = sections[-1][1]
            elif cur is None:
                continue
            else:
                m = IMG_RE.match(line.strip())
                if m:
                    cur.append(("image", m.group(1).strip()))
                    continue
                if line.lstrip().startswith(("-", "*")):
                    bullet = line.lstrip()[1:].strip()
                    if bullet:
                        cur.append(("text", bullet))
                elif line.strip():
                    cur.append(("text", line.strip()))
    if not title or not sections:
        raise SystemExit("大纲格式不对：需要 1 个 '# 主标题'，至少 1 个 '## 板块' 及其 '- 要点'")
    return title, sections


def wrap_bullet(text, limit=40):
    """过长要点按显示宽度手动折行，避免自动换行顶到页面边缘。"""
    import unicodedata
    lines, buf, width = [], "", 0.0
    for ch in text:
        w = 1 if unicodedata.east_asian_width(ch) in "FW" else 0.5
        if width + w > limit:
            lines.append(buf)
            buf, width = "", 0.0
        buf += ch
        width += w
    if buf:
        lines.append(buf)
    return lines


def add_images(slide, image_paths, top_start):
    """把图片依次放在 top_start 之下，每张等比缩放、水平居中。返回成功张数。"""
    n = len(image_paths)
    if n == 0:
        return 0
    available = SLIDE_H - top_start - Emu(400000)
    slot = Emu(int(available / n))
    max_w = Emu(6800000)
    placed = 0
    top = top_start
    for path in image_paths:
        try:
            pic = slide.shapes.add_picture(path, 0, 0)
        except Exception as e:
            print(f"警告: 图片插入失败，已跳过 {path}（{e}）", file=sys.stderr)
            continue
        iw, ih = pic.image.size
        ar = iw / ih
        w, h = max_w, Emu(int(max_w / ar))
        if h > slot:
            h = slot
            w = Emu(int(slot * ar))
        pic.width, pic.height = w, h
        pic.left = Emu(int((SLIDE_W - w) / 2))
        pic.top = top + Emu(int((slot - h) / 2))
        top = Emu(top + slot)
        placed += 1
    return placed


def build(title, sections, bg_path, out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]

    # 封面
    slide = prs.slides.add_slide(blank)
    add_background(slide, bg_path)
    box = slide.shapes.add_textbox(Emu(1000000), Emu(2800000), Emu(10192000), Emu(1400000))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title
    set_run_font(run, 40, bold=True, color=TITLE_COLOR)
    sub = slide.shapes.add_textbox(Emu(1000000), Emu(4300000), Emu(10192000), Emu(600000))
    p2 = sub.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "授课老师：陈老师"
    set_run_font(run2, 18, color=TITLE_COLOR)

    # 板块页
    for heading, items in sections:
        slide = prs.slides.add_slide(blank)
        add_background(slide, bg_path)
        head = slide.shapes.add_textbox(Emu(800000), Emu(500000), Emu(10600000), Emu(900000))
        tf = head.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = heading
        set_run_font(run, 30, bold=True, color=TITLE_COLOR)

        texts = [v for kind, v in items if kind == "text"]
        images = [v for kind, v in items if kind == "image"]
        body = slide.shapes.add_textbox(Emu(1000000), Emu(1600000), Emu(10200000), Emu(4800000))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        n_lines = 0
        for b in texts:
            for j, seg in enumerate(wrap_bullet(b)):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.space_after = Pt(10)
                run = p.add_run()
                run.text = ("• " if j == 0 else "   ") + seg
                set_run_font(run, 20)
                n_lines += 1
        # 有图片时按行数压缩正文框，图片放正文下方（每行约 38 万 EMU，留 15 万间距）
        img_top = Emu(1600000) + Emu(n_lines * 380000 + 150000) if images else None
        if images and img_top < Emu(5200000):
            add_images(slide, images, img_top)
        elif images:
            print(f"警告: 板块『{heading}』文字过多放不下图片，图片已省略，请拆页重试", file=sys.stderr)

    prs.save(out_path)
    print(f"已生成: {out_path}（封面 1 页 + 板块 {len(sections)} 页，共 {len(sections) + 1} 页）")
    print("请用 extract_ppt_text.py 回读核对知识点是否齐全。")


def main():
    ap = argparse.ArgumentParser(description="Markdown 大纲 → 知识点罗列式教学 PPT（支持插图）")
    ap.add_argument("outline", help="Markdown 大纲文件（UTF-8）")
    ap.add_argument("-o", "--output", required=True, help="输出 pptx 路径")
    ap.add_argument("-b", "--background", required=True, help="背景图路径")
    args = ap.parse_args()
    title, sections = parse_outline(args.outline)
    build(title, sections, args.background, args.output)


if __name__ == "__main__":
    main()
