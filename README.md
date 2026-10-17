# 科研场景 PPT Skill（蓝白信息图式）

[English](#english) | 中文

把“做一份答辩 / 开题 / 组会汇报 PPT”从每次重新摸索，变成一条可复用的流水线：同一套蓝白信息图式视觉，原生可编辑的 PPTX，内容来自你的材料。

这是给 AI agent 用的技能包（Hermes Agent / Claude Code / Codex 都能加载）：`SKILL.md` 是给 agent 的操作手册，`scripts/` 是干活的 Python 脚本，`references/info-style.md` 是视觉规范与版式库。

## 长什么样

每个内容页从上到下：页标题（“2.1 ｜ 两种方案对比”）→ ✓ 引导句 → 深蓝标题带（这页的一句话判断）→ 图文区（图 + 浅蓝小标题条 + ■ 要点、技术路线矩阵、闭环、时间轴、对比表……）→ ■ 底部结论条，红字点睛。配色与字号取自《蓝白科研答辩模板》实测：深蓝 #073DA5、正文 #111722、红 #DE1515、细边框 #9CB3D3、浅蓝 #EAF0FC。

`assets/info-demo.json` 是 28 页完整示例（园区综合能源系统负荷预测与优化调度）：封面、4 次目录、22 种内容页、结束页。

## 16 个推荐页型

大纲阶段先定页型，再从骨架起步，不要每页现编版式。总览图 `assets/page-types.jpg` 可以直接发给网页端 GPT；可套用的 JSON 骨架在 `assets/page-types.json`（按编号取 `skeleton`，只换文字和图）。

| 页型 | 用在哪 | 图位 |
|---|---|---|
| P01 背景需求页 | 研究背景、现实需求、研究对象 | 中图 1 + 中图 1 + 小图 1 |
| P02 双栏对比页 | 两个瓶颈、两种现状、前后对比 | 中图 2 |
| P03 方案优缺点页 | 两种方法 / 方案的选择 | 小图 2 |
| P04 大图解读页 | 论文框架、技术总图、核心结果 | 大图 1（role: big） |
| P05 流程步骤页 | 研究流程、方法步骤、实施过程 | 小图 4 |
| P06 技术路线矩阵页 | 数据层 / 方法层 / 验证层 | 小图 9 |
| P07 分层架构页 | 系统 / 平台架构、模块关系 | 无图 |
| P08 四环闭环页 | 四个环节的循环、迭代机制 | 小图 4 |
| P09 三栏并列页 | 三个方向、三项任务、三条结论 | 中图 3 |
| P10 对比表页 | 多维对比、方案选择、效益汇总 | 小图 4（表下图带） |
| P11 时间轴计划页 | 研究计划、阶段安排 | 无图 |
| P12 结论行页 | 主要结论、创新点、预期成果 | 小图 3 |
| P13 问题拆解页 | 当前卡点、讨论重点 | 无图 |
| P14 进展总表页 | 组会进展、任务完成情况 | 无图 |
| P15 论文信息卡页 | 文献汇报的出发点 | 小图 1 |
| P16 分点结论页 | 文献 / 结果的三条结论 | 中图 3 |

`references/outline.md` 里有把这些页型交给网页端 GPT 的大纲提示词：每页标页型编号、大 / 中 / 小图位、分点内容，拿回来就能直接填骨架。

## 配图分工

| 图的用途 | 谁画 |
|---|---|
| 概念、流程、角色、状态（检测 / 隔离 / 切换 / 重联、边界、三层、数据流、充放电……） | 网页端 ChatGPT，蓝白扁平风，一次要 8–9 张、每张独立 |
| 定量对比（成本高低、SOC 轨迹、误差分布、日负荷曲线） | 本机 matplotlib，图注写“示意数据” |
| 系统架构 / 单线接线 / 分层控制 | 网页端 ChatGPT，一页一张大图 |

插画贴进页里之前：要读的标签 ≥ 3 个就给到 3.5 in 宽（放不下换三栏 / 四栏，别缩图）；图里自带的标题条要裁掉；交付前降采样到 1200 px 宽 + 量化 200 色。

## 三条原则

- **一页只讲清楚一件事**：标题带就是这页的一句话判断，第二个判断就拆页。
- **信息密、留白少，但清晰**：空白用具体的图、小标题要点、流程条、表格填，不放大字号。
- **图要具体**：数据图、论文原图、平台截图优先，GPT 只画具体对象。

## 能产出什么

| 产物 | 说明 |
|---|---|
| 原生可编辑 PPTX | 文本框、表格、图形、流程、矩阵都能在 PowerPoint 里直接改 |
| Word 逐字稿 | 每页一段口语化讲稿 |
| 渲染预览 + 联系表 + PDF | 一次看全稿 |
| 检查报告 | 溢出、孤行、拆词、红字数量、留白、免责字样、编号跳号 |

## 安装

**Hermes Agent**：

```bash
git clone https://github.com/d2dmh/research-ppt-skill.git \
  ~/.hermes/skills/productivity/research-ppt-skill
```

**其他 agent**：把整个目录放到它能读到技能的地方，或在会话里说“读 `research-ppt-skill/SKILL.md`，按它执行”。

**依赖**：`pip install python-pptx python-docx lxml Pillow`；可选 `jieba`（断词检查）、`PyMuPDF` + LibreOffice（渲染预览）。详见 `INSTALL.md`。

## 怎么用

1. 让 agent 加载技能，告诉它主题；
2. 它会先要材料（场合与听众、时长、论文原图与数据、署名与 logo），再出逐页大纲（每页页型编号 + 大 / 中 / 小图位）；
3. 按大纲写内容 JSON → 生成 PPTX → 自动检查 → 渲染看图修改 → 交付 PPTX、讲稿、预览。改内容只动 JSON，不手改 PPTX。

## 目录结构

```
SKILL.md                 # 给 agent 的操作手册：流程、规则、踩坑清单、自检
references/
  outline.md             # 给网页端 GPT 的大纲提示词（页型编号 + 大中小图位）
  info-style.md          # 视觉规范、组件、22 种版式、“要讲的事 → 版式”对照
  visuals.md             # 配图来源、提示词、核对
  speaker-notes.md       # 讲稿规则
  embedded-video.md      # 视频内嵌
scripts/
  dense_deck.py          # 内容 JSON → PPTX
  info_style.py          # 蓝白信息图式：页面骨架、组件、版式
  lint.py / render.py    # 检查 / 渲染
  prep_figs.py           # 缺图 → 占位框
  insert_image.py        # 往占位框填图
  notes_docx.py          # 讲稿 → Word
  embed_video.py         # 内嵌 mp4
assets/
  page-types.jpg/.json   # 16 个推荐页型：总览图 + 可套用骨架
  info-demo.json         # 28 页示例
  images/                # 示例数据图（示意数据）与重画脚本
  components/            # 可复用图标
  branding/              # 默认机构标识（请替换成你自己的）
```

## 已知边界

- 生成后仍要逐页看渲染图修改：脚本保证规范，版面是否饱满、图够不够大要靠看。
- 配图需要你提供或用网页端 ChatGPT 生成；没有的图位先出占位框和提示词。
- `assets/branding/xmu-*.png` 是作者的默认机构素材，使用前请替换。

## 许可

代码与文档 MIT。`assets/branding/` 下的机构标识版权归原机构，不在 MIT 授权范围内。

---

## English

**research-ppt-skill** — a skill package for AI agents that produces **native, editable** research presentation decks (thesis defense, proposal, group meeting) in one consistent blue-and-white infographic style: page title, ✓ lead sentence, navy title band stating the page's single message, a dense but clean figure-and-bullet area (tech-route matrices, layered architectures, cycles, timelines, comparison tables), and a ■ takeaway bar.

Includes an agent-facing `SKILL.md`, a library of 16 reusable page types (`assets/page-types.jpg` / `.json`, plus an outline prompt in `references/outline.md`), a component and layout library (`references/info-style.md`, 22 layouts), Python scripts for generation (`python-pptx`), linting (overflow, orphan lines, red-emphasis count, disclaimers), rendering previews and Word speaker notes, plus a 28-page example deck.

MIT licensed. Branding assets belong to their owner and should be replaced.
