# 科研场景 PPT Skill

[English](#english) | 中文

把「按我们组那份答辩稿的样式，做一份科研汇报 PPT」这件事，从**每次重新摸索**变成**一条可复用的流水线**。

这是给 AI agent 用的技能包（Hermes Agent / Claude Code 之类的编码助手都能加载）：`SKILL.md` 是给 agent 看的操作手册，`scripts/` 是真正干活的 Python 脚本。

---

## 它到底是什么

**做的是风格迁移，不是套模板。** 给它一份参考稿（答辩 PPT、组会 PPT、别人的好稿子都行），它会把参考稿的**字号、色号、版面坐标、信息密度**量出来，然后用**原生可编辑的 PPTX**（不是图片拼出来的）复现同一套视觉语言，内容换成你的研究。

**它不做什么**：不提供模板下载，不帮你想研究内容，不把内容外包给 AI。内容永远来自你给的材料。

## 能产出什么

| 产物 | 说明 |
|---|---|
| 原生可编辑 PPTX | 每个文本框、表格、图形都能在 PowerPoint 里直接改 |
| Word 逐字稿 | 每页一段口语化讲稿，含该页要讲的重点与过渡 |
| 渲染预览 + 联系表 + PDF | 一次性看全稿，用于自查和转发 |
| 校验报告 | 溢出、孤行、拆词、行末标点、小字号、分区留白，逐条列出 |

## 踩出来的硬规则（都在 `SKILL.md` 里）

- **红色是强调色**：一页最多 1 处，只落在"判断句"上；表格里不用红，术语名不标红。
- **字号下限**：投影正文 ≥16pt、子点 12–13pt、**表格 ≥13pt**（11/12pt 投影看不清）。
- **每页只服务一个判断**，核心要点 ≤4 条；一页说两遍的内容删掉一处。
- **目录页在每个部分开头重复出现**，高亮当前部分（参考稿的既有结构）。
- **成品稿不出现任何免责或自证文字**（"AI 生成示意""逐字核对""依据材料""待补"）——成品不是草稿。
- **一页一张 16:9 的图最大约 9 英寸宽**（受版面区高限制）；素材/录屏位要占主体，别塞半栏。
- **配图尺寸用脚本量，不靠感觉**：半栏放 16:9 的图，`h ≈ 区宽 × 9/16`；交付前打印每张图的实际英寸 + 表格字号（≥13pt）+ 图表坐标字号（≥12pt）核对一遍。图不够大就抢高度、移表格、换版式，**永远不靠缩图腾地方**。
- **交付前做「质感对照」**：数一下这份稿子里 `chart`/`formula`/`table`/`figure` 各有几个（10 页级汇报：原生图表 ≥3、公式 ≥1）；再把新稿和已有稿子的联系表并排同屏看——一页都不带图的稿子必然看着薄。
- **案例小节先放演示、再讲原理**：录屏页放在小节第一页，别压在小节最后。
- **录屏位要真内嵌视频**（`scripts/embed_video.py`）：mp4 进 `ppt/media/`，载体图挂 `a:videoFile` + `p14:media` + `ppaction://media`，点击封面即可播放；只放截图或只挂文件链接都不算数。
- **密度有实测数值**：某个栏位一行放多少汉字、结论条多少字会折行、图注多少字会溢出，都在 `SKILL.md`，能省掉好几轮返工。

## 安装

**Hermes Agent**（推荐）：

```bash
git clone https://github.com/d2dmh/research-ppt-skill.git \
  ~/.hermes/skills/productivity/research-ppt-skill
```

**其他 agent**：把整个目录放到它能读到技能/知识库的地方；`SKILL.md` 就是全部说明，脚本是纯 Python，可以直接调用。

**依赖**：

```bash
pip install python-pptx python-docx lxml Pillow
# 可选：jieba（中文断词检查）、PyMuPDF（渲染成图）
```

渲染预览用 LibreOffice（macOS / Linux）；有 PowerPoint 时最终断行以 PowerPoint 为准。

## 怎么用

1. 让 agent 加载这个技能，把**参考稿**和**主题**给它；
2. 它会先问你要材料（场合与听众、时长/页数、核心内容、署名与 logo、必须讲的重点），再出**逐页大纲**给你确认；
3. 确认后生成 PPTX、讲稿与预览，按校验报告改——改内容只动内容 JSON，不手改 PPTX。

## 目录结构

```
SKILL.md              # 给 agent 的操作手册（规则、流程、实测数值、坑）
scripts/
  dense_deck.py       # 内容 JSON → PPTX
  layouts.py          # 版面骨架与分区摆放
  styled_deck.py      # 文本框、表格、流程图、图位等元素
  lint.py             # 溢出 / 孤行 / 拆词 / 小字号 / 留白 校验
  render.py           # 渲染成图 + 联系表 + PDF
  notes_docx.py       # 逐页讲稿 → Word
  insert_image.py     # 往图位里插真实图片
  embed_video.py      # 把 mp4 真正内嵌进 PPTX（录屏页用）
  prep_figs.py        # 生成前处理图位：图缺失则降级为占位框
  fix_typography.py   # 中英混排与标点修正
references/           # 风格系统、版式对照、密度数值、讲稿写法、踩坑记录
assets/               # 演示用内容 JSON、机构标识、字体配置
```

## 已知边界

- 默认风格来自一份**中文科研答辩 PPT**；换机构要换掉 `assets/branding/` 里的标识。
- 仓库**不包含任何参考稿文件**——用别人的稿子提取风格时，请自行取得授权，不要把原稿提交上来。
- `assets/branding/xmu-*.png` 是作者的默认机构素材，**使用前请替换成你自己的**。

## 许可

代码与文档 MIT。`assets/branding/` 下的机构标识版权归原机构，不在 MIT 授权范围内，仅作默认占位，请自行替换。

---

## English

**research-ppt-skill** — a skill package for AI coding agents that turns a sample deck (e.g. a thesis defense) into a reusable pipeline for producing **native, editable** research presentation decks in the same visual language: it measures the reference's type scale, colors, layout coordinates and information density, then regenerates the deck with your own content.

Contains an agent-facing `SKILL.md` plus Python scripts for deck generation (`python-pptx`), linting (overflow / orphan lines / word splitting / font-size floors), rendering previews, and producing per-slide speaker notes as Word.

Hard rules baked in from practice: red as a single emphasis per slide, minimum type sizes (body ≥16pt, tables ≥13pt), one judgment per slide, repeated agenda slides, and no defensive/disclaimer text on finished slides.

MIT licensed. No reference decks are included in this repository. Branding assets belong to their owner and should be replaced.
