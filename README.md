# math-lesson-prep · 数学备课全流程 Skill

一个可直接安装到各类 AI 编程助手（ZCode / Claude Code 等）的教师备课技能：发来 **PPT、Word、PDF、题目图片、零散文字** 中的任意组合，按固定流程产出一节课的全部备课材料。

## 它做什么

输入任意材料组合后，按以下流程工作：

```
① 多格式识别（PPT/Word/PDF/图片/文字）→ ② 生成或修改 PPT → ③ 全部页面换机构背景图
→ ④ 知识点清单 → ⑤ 课后反馈 → ⑥ 学习规划模板 → ⑦ 开课话术 → ⑧ 汇总交付
```

**PPT 三条处理路径**（模型自动判断，拿不准会问你）：

| 你说 | 它做 |
|---|---|
| "这个 PPT 直接能用" | 只把每页背景换成同益文化底图 |
| "基于这个 PPT 改" | 在原 PPT 上增删修改后换背景 |
| 内容零散 / 多份材料 / 要求重做 | 提炼合并知识点 → 重新生成新 PPT |

**多格式输入**：PPT（.pptx）、Word（.docx）、PDF、图片（题目截图会转写进课件，几何图形会直接插图进对应页）、零散文字要求。

**两种学习规划**：说"全局规划"→ 联网搜集课件/大纲写跨课次全局表；默认 → 只生成本节课规划表。

**课后反馈**：模板套话全部写好，课堂上才知道的内容留【待填写】人工补写。

## 安装

### ZCode

```bash
# 用户级（所有项目可用）
git clone https://github.com/wanyuchen999-ai/math-lesson-prep.git ~/.agents/skills/math-lesson-prep

# 或项目级
git clone https://github.com/wanyuchen999-ai/math-lesson-prep.git <你的项目>/.agents/skills/math-lesson-prep
```

### Claude Code

```bash
git clone https://github.com/wanyuchen999-ai/math-lesson-prep.git ~/.claude/skills/math-lesson-prep
```

### 其他智能体 / 大模型

把整个文件夹放到该工具的技能发现目录（如 `.agents/skills/`、`skills/`），或直接把 `SKILL.md` 的内容粘贴进系统提示词/自定义指令，并保留 `assets/`、`references/`、`scripts/` 的相对路径文件。

## 使用

安装后直接对话即可触发，例如：

- 「发你两个 PPT，帮我备课：平行四边形面积」
- 「这个 PPT 直接用，换个背景就行」
- 「这张图里的例题加到 PPT 上，然后做本节规划」
- 「写课后反馈 + 开课话术」
- 「做一份暑期全局规划模板」

## 依赖

| 功能 | 依赖 |
|---|---|
| PPT 生成/提取/换背景 | `pip install python-pptx` |
| PDF 提取 | `pip install pypdf`（可选） |
| Word 提取 | 无需依赖 |

## 目录结构

```
math-lesson-prep/
├── SKILL.md                      # 技能主文件（流程与规则）
├── assets/
│   └── tongyi-background.png     # 同益文化默认背景图
├── references/
│   ├── feedback-template.md      # 课后反馈模板
│   ├── plan-lesson-template.md   # 本节课学习规划模板
│   ├── plan-global-template.md   # 全局学习规划模板
│   └── opening-script.md         # 正式课/体验课开课话术
└── scripts/
    ├── extract_ppt_text.py       # PPT 逐页文字提取
    ├── extract_docs.py           # docx/pdf/md 提取
    ├── build_lesson_ppt.py       # Markdown 大纲 → 知识点罗列式 PPT（支持插图）
    └── apply_background.py       # 整本 PPT 换背景图
```

## License

MIT
