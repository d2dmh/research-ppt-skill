# 科研场景 PPT Skill · 蓝白信息图式

给 AI agent 用的一份 PPT 技能包。你把主题和材料给它，它交回来三样东西：能直接打开改的 PPTX、一份口语讲稿、一份全稿预览。

生成的 PPTX 里，文字框、表格、流程条、箭头都是真对象，不是整页截图，PowerPoint 里点开就能改字。

`MIT` · `python-pptx` · 16 个推荐页型 · 22 种版式 · 逐页渲染核对

---

## 为什么写这个

我一年要做七八次汇报：开题、中期、组会、文献分享、技术交流。每次内容不同，排版却要从头再来一遍，光对齐全和调字号就要两三天，做出来的几份稿子长相还都不一样。

后来我把学院那份蓝白答辩模板的色号和字号一个个量下来（深蓝 `#073DA5`，页标题 32 pt，要点 14.4 pt），写成脚本，再把常见的页面结构拆成 16 个页型。现在换主题只换内容，版式不用重新想。

---

## 先看产出

### 一份 21 页的成品

微电网技术分享：封面、4 次目录过渡、18 个内容页、结束页。整份稿子从一份内容 JSON 生成，没有手工排版。

![全稿缩略图](docs/preview-deck.png)

### 四页细节，点开看原尺寸

**① 大图解读页。** 一张接线图占七成版面，右边三小节写"看哪里、关键设备、为什么这样接"。图的尺寸由脚本按格子算，不会被压成邮票。

![大图解读页](docs/preview-page-big.png)

**② 信息密度页。** 场景图、典型日曲线、分时电价、三条判断放在一页。留白靠内容填，没有一处是靠放大字号撑起来的。

![信息密度页](docs/preview-page-dense.png)

**③ 四栏并列页。** 四个动作各配一张插画、一个判据框、两条要点。上面那条蓝带只写一个判断。

![四栏并列页](docs/preview-page-four.png)

**④ 对比表页。** 表头深蓝，首列浅蓝，关键数字标红，整行加红框。表下面压一条四张小图的证据带，不然表格单独一页会显得空。

![对比表页](docs/preview-page-table.png)

### 16 个页型长什么样

![16 个推荐页型](docs/preview-page-types.jpg)

---

## 六种场合，各有各的走法

| 场合 | 一般页数 | 常用页序 | 重点 |
|---|---|---|---|
| 开题 / 答辩 | 20–28 | P01 背景 → P02 瓶颈 → P05 流程或 P06 技术路线 → P07 架构 / P08 闭环 → P04 核心结果 → P11 计划 → P12 结论 | 一页一个判断，创新点单独给页，不留"待补" |
| 中期检查 | 15–20 | P14 进展总表 → P09 三项任务 → P04 关键图 → P13 卡点 → P11 后续计划 | 已完成写产物，没做完写原因，不写"正在推进"四个字 |
| 组会进展 | 10–15 | P14 进展总表 → P09 实验对照 → P13 当前卡点 → 三栏计划 → 结束页 | 用产物说话，卡点拆成可验证的假设，末尾留一个要讨论的决定 |
| 文献汇报 | 12–18 | P15 论文信息卡 → P02 背景数据 → P04 论文框架 → P16 三条结论 → P04 重点图解读 → P02 我的评价 | 每张图说清"看哪里、看到什么、对我们有什么用" |
| 技术分享 / 入门讲 | 18–24 | P01 为什么需要 → P02 两种状态 → P04 一张总图 → P07 分层 → P03 方案对比 → P09 三环节 → P12 结论 | 用一个例子从头贯穿，别堆名词 |
| 竞选 / 活动策划与复盘 | 8–15 | 策划：信息条 → 任务表 → 备用方案；复盘：目标与实际 → 差异原因 → 改进动作 | 承诺要能核实，复盘只讲动作和检查点 |

页型不是必须照搬。塞不下就换版式，`references/info-style.md` 里有 22 种，还附了"要讲的事 → 选哪个版式"的对照表。

---

## 装之前要知道的三件事

**一、这是给 agent 用的技能包，不是独立软件。** 它靠 `SKILL.md` 指导 agent 干活，脚本负责排版。你自己不用读代码，但要有一个能加载技能的 agent。

**二、默认搭档是"网页端 ChatGPT 协作"那套。** 就是 `chatgpt-web-planning` 这个前置技能，仓库在 [hermes-chatgpt-web](https://github.com/d2dmh/hermes-chatgpt-web)，我平时叫它 Hermes with GPT。流水线里有两步交给它：出逐页大纲、画概念插画。

没装这套前置也能用，只是这两步要自己补：

| 有前置 | 没有前置 |
|---|---|
| 大纲交给网页端 GPT，用 `references/outline.md` 里的提示词 | 自己按 16 个页型写大纲，`assets/page-types.json` 里有现成骨架 |
| 概念插画让网页端画，取回后核对再贴 | 用本机 matplotlib 画示意，或者用手头的设备照片、论文原图 |

**三、实测范围。** 我只在 Hermes Agent 加网页端 ChatGPT 这条链上完整跑过，答辩稿、组会稿、技术分享稿各一份。Claude Code、Codex 这些能读 `SKILL.md` 的 agent 理论上都能用，但我没有实测记录，遇到问题开 issue。

另外，渲染预览要 LibreOffice（`soffice` 在 PATH 里）。Windows 上也能跑，但预览字体和 PowerPoint 有差别，终稿建议在 PowerPoint 里过一遍。

到底要装哪些东西，按你的环境看这张表：

| 你的环境 | 要做的 |
|---|---|
| Hermes Agent | 把仓库 clone 到 `~/.hermes/skills/productivity/`；建议同时装 `chatgpt-web-planning`，大纲和插画交给网页端 |
| Claude Code / Codex | 把目录放进仓库，在会话里让它读 `SKILL.md`；网页端协作那两步看它自己有没有对应的能力，没有就手工 |
| 不想用 agent | 也能跑：自己写 `content.json`，直接调 `scripts/dense_deck.py`，只是大纲、配图和看图修改都得自己来 |

Python 3.10 以上就够，我这台是 3.12 + python-pptx 1.0.2 实测的。

---

## 一分钟上手

```bash
git clone https://github.com/d2dmh/research-ppt-skill.git \
  ~/.hermes/skills/productivity/research-ppt-skill
```

然后跟 agent 说人话就行，比如：

> 帮我做下周三组会汇报的 PPT，微电网这个方向，材料在 `~/docs/组会材料.docx`，讲 15 分钟。

> 用 research-ppt-skill 把这篇论文做成文献汇报，12 页左右，重点讲它的方法和那张 Fig.3。

> 做个开题答辩稿，20 分钟，先把大纲给我看一眼。

它会先问场合、听众、时长、署名和素材，然后出逐页大纲（每页标页型和大中小图位），再生成、自检、逐页看图，最后交 PPTX、讲稿和预览。改内容只动 `content.json`，不手改 PPTX。

---

## 一页是怎么搭的

每页从上到下五层，所以二十几页读下来是同一个节奏。

| 层 | 长什么样 | 干什么用 |
|---|---|---|
| 页标题 | `2.3 ｜ 孤岛能不能撑住看三条` | 告诉听众这页讲什么 |
| 引导句 | 前面一个 ✓，一句话交代背景 | 把人带进这页的语境 |
| 蓝底标题带 | 白字，20.4 pt，一句判断 | 这页只讲这一件事 |
| 图文区 | 图、浅蓝小标题条、■ 要点、矩阵、流程、表格 | 摆证据 |
| 结论条 | 前面一个 ■，红字点睛 | 听众带走的那一句 |

红字一页最多四处，只给判断、关键数字、核心提法。组织性词语不标红。

字号是量出来的，不是估的：页标题 32、引导句 16.8、标题带 20.4、小标题条 15.6、要点 14.4、图注 13.2。放不下就往下收到 12，从不往上放大。这条是被逼出来的：留白多的时候第一反应都是放大字号，那样只会更难看，正确做法是往里填图、填小标题要点、换更紧凑的版式。

---

## 16 个页型

大纲阶段先定页型，再拿骨架起步，不要每页现编。骨架在 `assets/page-types.json`，按编号取 `skeleton`，换掉文字和图。

| 页型 | 用在哪 | 图位 |
|---|---|---|
| P01 背景需求页 | 研究背景、现实需求、研究对象 | 中图 1 + 中图 1 + 小图 1 |
| P02 双栏对比页 | 两个瓶颈、两种现状、前后对比 | 中图 2 |
| P03 方案优缺点页 | 两种方法选一个 | 小图 2 |
| P04 大图解读页 | 论文框架、技术总图、核心结果 | 大图 1（`role: big`） |
| P05 流程步骤页 | 研究流程、方法步骤、实施过程 | 小图 4 |
| P06 技术路线矩阵页 | 数据层、方法层、验证层 | 小图 9 |
| P07 分层架构页 | 系统架构、平台架构、模块关系 | 无图 |
| P08 四环闭环页 | 四个环节的循环、迭代机制 | 小图 4 |
| P09 三栏并列页 | 三个方向、三项任务、三条结论 | 中图 3 |
| P10 对比表页 | 多维对比、方案选择、效益汇总 | 小图 4（表下图带） |
| P11 时间轴计划页 | 研究计划、阶段安排 | 无图 |
| P12 结论行页 | 主要结论、创新点、预期成果 | 小图 3 |
| P13 问题拆解页 | 当前卡点、讨论重点 | 无图 |
| P14 进展总表页 | 组会进展、任务完成情况 | 无图 |
| P15 论文信息卡页 | 文献汇报的出发点 | 小图 1 |
| P16 分点结论页 | 文献或结果的三条结论 | 中图 3 |

`references/outline.md` 里有一段可以直接复制的提示词。连同 `assets/page-types.jpg` 一起发给网页端 GPT，它就会按页型编号和大中小图位给你排大纲。

---

## 风格

现在只有一种：科研风，蓝白信息图。是从学院那份答辩模板反推出来的，适合答辩、开题、中期、组会、文献汇报这类场合。

后面想加的：

| 风格 | 状态 | 打算长什么样 | 给谁用 |
|---|---|---|---|
| 科研风 · 蓝白信息图 | 已发布 | 深蓝标题带、✓ 引导句、浅蓝小标题条、■ 结论条 | 答辩、开题、组会、文献汇报 |
| 商务极简风 | 规划中 | 还没定 | 给甲方讲方案、项目评审 |
| 深色技术风 | 规划中 | 还没定 | 技术分享、产品发布、大屏 |

每种风格是一套独立的配色、字号和版式实现，内容 JSON 和流程不动，只换皮。当前这套写在 `scripts/info_style.py`。

想加某一种，开个 issue 说清楚场景和听众（"给甲方讲方案，30 页，要稳"），比凭空设计有用。自己动手也行：照 `info_style.py` 新写一个 style 脚本，在 `SKILL.md` 里加一节，页面骨架和组件接口可以直接复用。

---

## 图从哪来

| 图的用途 | 谁画 | 为什么 |
|---|---|---|
| 概念、流程、角色、状态（检测、隔离、切换、重联、边界、三层、数据流、充放电这些） | 网页端 ChatGPT，蓝白扁平风，一次要八九张，每张独立 | 比脚本画的示意图好看，而且不承载具体数值 |
| 定量对比（成本高低、SOC 轨迹、误差分布、日负荷曲线） | 本机 matplotlib，图注写"示意数据" | GPT 画不准定量关系，硬画就是编数据 |
| 系统架构、单线接线、分层控制 | 网页端 ChatGPT，一页一张大图 | 关系复杂，交给它自己判断构图 |

贴图有三个坑，都踩过：插画里要读的标签超过三个，格子宽度得给到 3.5 in 以上，放不下就换三栏四栏，别缩图；网页端给的插画常自带标题条，会和页里的标题带重复，得先裁掉；交付前把插画降到 1200 px 宽并量化 200 色，16 MB 的稿子能压到 11 MB。

---

## 交给你哪些文件

| 文件 | 内容 |
|---|---|
| `.pptx` | 主件，所有元素可编辑 |
| 逐字稿 `.docx` | 每页一段口语讲稿，带预计时长 |
| `.pdf` 加联系表 | 手机上翻着看，一眼扫全稿 |
| 检查报告 | lint 输出的问题清单 |
| 交付说明 | 图文来源、哪些是示意数据、后续怎么改（md 和 docx 各一份） |

---

## 检查在查什么

生成时脚本会打印问题：分区溢出（`over-full`）、内容下方空了一块、小图被塞进太小的格子、一页放了两张大图、高亮词在正文里找不到。这些问题先清零，再往下走。

然后跑 `lint.py`，它查溢出、孤行、拆词、红字数量、编号跳号、AI 免责声明、残留的图片占位框。目标 0 错误 0 警告。

最后逐页看渲染图。脚本查不出"分区里空了一块""图在格子里太小"，这两类只能靠眼睛，所以交付前每一页我都看过。

---

## 安装与依赖

```bash
# Hermes Agent
mkdir -p ~/.hermes/skills/productivity
cp -R research-ppt-skill ~/.hermes/skills/productivity/

# Python 依赖
pip install python-pptx python-docx lxml Pillow
# 可选
pip install jieba            # 断词检查
pip install PyMuPDF          # 渲染预览
# 渲染还要 LibreOffice，soffice 在 PATH 里
```

各平台的细节和坑写在 [`INSTALL.md`](INSTALL.md)。

---

## 目录结构

```
SKILL.md                 # 给 agent 的操作手册：流程、规则、踩坑清单、自检
INSTALL.md               # 安装与依赖
references/
  outline.md             # 交给网页端 GPT 的大纲提示词
  info-style.md          # 视觉规范、组件、22 种版式、按场合选版式的对照
  visuals.md             # 配图来源、提示词写法、取回核对
  speaker-notes.md       # 讲稿规则
  embedded-video.md      # 录屏页真内嵌 mp4
scripts/
  dense_deck.py          # 内容 JSON → PPTX
  info_style.py          # 科研风的页面骨架、组件、版式
  lint.py / render.py    # 检查 / 渲染
  prep_figs.py           # 缺图 → 占位框 + 提示词
  insert_image.py        # 往占位框里填图
  notes_docx.py          # 讲稿 → Word
  delivery_doc.py        # 交付说明 md → docx
  embed_video.py         # 内嵌 mp4
assets/
  page-types.jpg/.json   # 16 个页型的总览图与骨架
  info-demo.json         # 28 页示例，答辩 / 开题
  scenes-demo.json       # 19 页示例，组会 / 文献 / 竞选 / 活动
  images/ components/    # 示例数据图、可复用图标
  branding/              # 默认机构标识，用之前换掉
docs/                    # README 里的预览图
```

---

## 已知问题

- 生成完还得逐页看。脚本保证规范和字号，版面饱不饱满、图够不够大，只能看。
- 配图要另外准备。要么你给论文原图和数据，要么让网页端 GPT 画。缺图的时候会先生成占位框和提示词，但占位框不能交出去。
- 定量图里的数字是示意。图注会写"示意数据"，正式汇报前换成你的实测值，不然结论站不住。
- `assets/branding/xmu-*.png` 是我的学院标识，用之前换成你自己的。
- Windows 上渲染预览字体和 PowerPoint 有差别，终稿在 PowerPoint 里过一遍最稳。

---

## 如果它省了你两天排版

右上角点个 star，别人搜"答辩 PPT""组会 PPT 模板"时才可能翻到这儿。同门还在手搓对齐的，也可以直接转给他。

有需求开 issue，说清场合、听众、页数，我按场景加页型或风格。做出来的稿子想分享，也欢迎贴在 issue 里。

---

## English

A skill package for AI agents that produces native, editable Chinese academic decks: defense, proposal, mid-term review, group meeting, paper reading, technical talk.

Give it a topic and your material. It returns a `.pptx` where every text box, table and shape is still editable, a Word speaker-note script with per-page timing, a PDF plus contact sheet, and a lint report (overflow, orphan lines, red-emphasis count, leftover placeholders).

A 21-page generated deck lives in `docs/`: a contact sheet plus four representative pages.

What's inside: `SKILL.md` (the agent manual), 16 reusable page types with an outline prompt for web GPT (`assets/page-types.jpg` / `.json`, `references/outline.md`), a component and layout library (`references/info-style.md`, 22 layouts), and scripts for generation, linting, rendering and speaker notes. Two example decks ship with it: a 28-page defense and a 19-page group meeting.

Requirements: Python with `python-pptx`, `python-docx`, `lxml`, `Pillow`; LibreOffice for preview rendering. The full pipeline was tested on Hermes Agent together with the `chatgpt-web-planning` companion skill, which hands the outline and the illustrations to web ChatGPT. Without that companion you plan the outline yourself and bring your own figures.

The blue-and-white academic style is the only one released so far. A business-minimal and a dark technical variant are planned. Open an issue describing your scenario if you want one sooner.

MIT licensed. Files under `assets/branding/` belong to their owner and should be replaced. If this saves you a couple of days of layout work, a star helps other people find it.
