# 信息图式科研页（info 风格，内容页默认）

来源：用户给的《蓝白科研答辩模板｜12 页》（Final_editable.pptx 实测参数）、三张生图版式稿（流程 / 技术路线 / 对比，图文 / 三并列 / 四并列，时间轴 / 架构 / 结论），以及“像素答辩”参考页（2.1 研究问题、2.2 目标解决、3.1 认知边界、3.4 成果应用、4.1 应用推广、4.2 经济社会效益）。
示例：`assets/info-demo.json`（科研答辩 / 开题：园区综合能源系统负荷预测与优化调度，28 页）与 `assets/scenes-demo.json`（组会进展、文献汇报、竞选、活动策划与复盘，19 页）。

## 零、总原则

1. **一页只讲清楚一件事**。动笔前先写出这页的一句话判断，它就是标题带；引导句交代背景，总结条交代落脚点，三句连起来就是这页的逻辑。出现第二个判断就拆页；能两页讲清的不用三页。
2. **信息密、留白少，但一眼能看清**。空白用具体内容填（图、小标题 + 要点、标签、表格、流程条），不靠放大字号；用户明确说过“放大字体还不如不改”。分区之间只留 0.1–0.15 in。
3. **一句话说不清就分点**：有并列、先后、条件关系的内容一律拆成 ■ 要点，每点一行（全局适用，生成日志会提示长句）。
4. **图分大小，大纲阶段就定**：大图信息密度最高（框架、技术路线、核心结果），用 `info_big` 给足空间；中图一张证明一个结论；小图信息密度最低（图标、缩略、简单示意），只看形态。定义与数量见 SKILL.md「图的分级」，16 个推荐页型见 `assets/page-types.jpg`。
5. **图要具体**。数据图、论文原图、平台截图、设备 / 场景图优先；GPT 生图要画具体对象和关系，不画抽象概念图。一页至少一张图或一个图形组件（流程、矩阵、闭环、时间轴）。

## 一、视觉规范（模板实测）

| 项 | 取值 |
|---|---|
| 主色 | 深蓝 `#073DA5`（标题、标题带、标签、编号、方块符号） |
| 正文 | `#111722`；强调红 `#DE1515` |
| 分区框 | 白底，细边框 `#9CB3D3` 0.6 pt；分区可选虚线框 |
| 浅蓝 | `#EAF0FC`：小标题条、说明块、表格首列、单元格、输出条 |
| 箭头 | 浅蓝大箭头 `#BCD3F6`；矩阵行间用深蓝细箭头 |
| 页眉 | “2.1 ｜ 标题”：编号 33 pt + 深蓝竖线 + 标题 32 pt，右上机构标识，下方 `#D5DEEB` 通栏细线（y = 0.8 in） |
| 边距 | 左右 0.32 in，内容宽 12.7 in；无页脚（`meta.info_footer: true` 恢复深蓝页脚与页码） |

### 字号（pt）

| 元素 | 字号 | 元素 | 字号 |
|---|---|---|---|
| 页标题 / 编号 | 32 / 33 | 引导句（✓） | 16.8 加粗 |
| 标题带 | 20.4 白粗 | 副引导句 | 15.6 深蓝粗 |
| 分区标题带 | 18 白粗 | 小标题条 | 15.6 深蓝粗 |
| 说明块 | 15 | 方块要点 / 正文 | 14.4 |
| 矩阵单元格 | 15.6 深蓝粗 | 图注 / 表格 | 13.2 |
| 总结条（■） | 15.6 加粗 | 下限 | 12 |

文字组件只在放不下时自动往下收（不低于 12 pt），**从不自动放大**。

## 二、页面骨架（从上到下）

| 层 | 写法 | 说明 |
|---|---|---|
| ① 页标题 | `"title": "2.1 两种方案对比"` | 结论式或主题式，≤ 14 字 |
| ② 引导句 | `"lead": {"text", "highlights", "tag"}` | 深蓝细框 + ✓，一行最好；写 `tag: "核心目标"` 时左侧换成深蓝标签（参考图 3.1） |
| ③ 副引导句（可省） | `"sublead": {"text", "highlights"}` | 无框、居中、深蓝粗体，点出研究路径（参考图 2.1 / 2.2 / 3.4） |
| ④ 标题带（可省） | `"band": "…"` 或 `{"text", "images": [...]}` | 本页的一句话判断；写 `images` 是图片衬底的标题带（参考图 4.2） |
| ⑤ 流程条（可省） | `"process": ["…", "…"]` 或 `{items, style: "pill"}` | 全宽箭头流程（参考图 3.1），可替代标题带 |
| ⑥ 图文区 | `"layout"` + `"regions"` | 见四、五 |
| ⑦ 总结条 | `"summary": {"text", "highlights"}` | 白底细框 + ■，一句话；可用“结论：”“使用建议：”开头 |

字数：引导句 ≤ 40 字（一行），总结条 ≤ 40 字，标题带 ≤ 26 字。红字一页最多 4 处，落在判断、关键数字、核心提法上；红色标签说明行的标签不计。

## 三、区域选项（region）

| 键 | 作用 |
|---|---|
| `grid: [列, 列数, 行, 行数]` | 24 × 12 网格自由拼版（不用预设时） |
| `fill` | `frame`（默认白底细框）/ `dashed` 虚线框 / `light` 浅蓝底 / `none` / 色号 |
| `band` + `num` | 分区顶部深蓝标题带，`num` 加白底编号圆（生图稿四项任务） |
| `head` + `num` + `head_style` | 分区内浅蓝小标题条；`num` 加深蓝编号圆（模板 3.1）；`head_style: "bar"` 左侧加深蓝竖条（生图稿结果一 / 二 / 三） |
| `side` + `side_w` | 左侧深蓝标签：窄（默认 0.44 in）竖排“优化措施”，宽（≥ 0.9 in）横排“适用场景\n与选择建议” |
| `foot` | 分区底部深蓝结论条（参考图 3.1） |
| `title` | 分区内居中深蓝小标题 |
| `valign` | `middle`（默认）/ `top` / `spread`（条目间均分） |
| `gap`、`pad`、`grow` | 间距、内边距、文字撑满分区（表格与标签说明行默认撑满） |

## 四、组件（region 的 items）

### 文字类
| type | 样子 | 用途 |
|---|---|---|
| `section` | 浅蓝小标题条 + ■ 要点：`{head, items, text, num, head_style}` | **最常用单元**：方法思路 / 优点 / 局限、主要研究任务 / 预期成果、阶段目标 |
| `markers` | ■ 要点：`items: ["…"]` 或 `[{label, text, highlights}]`；`fill: "light"` 放进浅蓝块 | 条件、要点、“研究基础：数据与方法支撑” |
| `note` | 浅蓝说明块：`{title, text}`；`tone: "red"` 浅红底红框红粗字，`dashed` 虚线 | 一两句说明、问题一 / 问题二、关键结论 |
| `para` | 无框文字：`color: ink / navy / red / white`，`align` | 栏首说明、红色小标题（参考图 4.1） |
| `chevrons` | » 小标题 + 说明；`tone: "navy"` 深蓝（参考图 2.2） | 思路、措施、解决的问题 |
| `label_rows` | 虚线框行：左红色粗体标签（可竖排两行）+ 右说明 | 现存问题 / 研究思路（参考图 2.1） |
| `links` | 虚线短语框 —— 深蓝关键词：`[["突破…", "自主可控"], …]` | 目标与定位对应（参考图 2.2） |
| `callout` | 左红竖条 + 加粗论断 | 一页最想让人记住的那句话 |
| `pill` | 深蓝块白粗字；`light` 浅蓝底，`full` 撑满 | 图组名、研究对象、关键问题 |
| `hexes` / `tags` | 白底深蓝框六边形 / 标签排 | 并列关键词 |

### 结构类
| type | 样子 | 用途 |
|---|---|---|
| `stages` | 阶段标题条（`head_style: light` 浅蓝）+ 图 + 要点 + 浅蓝“输出”条，阶段间浅蓝箭头 | 研究流程、关键步骤（模板 04 / 生图稿 1.1） |
| `matrix` | 行标签 + 单元格（文字、`{title,text}`、`{figure,caption}`、`{figure,items}`）+ 行间细箭头 + 顶部目标条 / 底部输出条 / 右侧支撑列 / 每行要点 | 总体技术路线、分层架构、多维对比（模板 05、11，生图稿 2.2 / 3.3 / 2.1） |
| `timeline` | 浅蓝阶段框 → 带节点的轴线（末端箭头）→ 时间 → `blocks` 小节；阶段间虚线 | 研究计划（模板 10 / 生图稿 1.1） |
| `cycle` | 四段环 + 中心文字（从左上起顺时针） | 四个环节的循环（模板 03） |
| `process` | 箭头流程条；`gradient` 由浅到深，`style: pill` 圆角块 + 小箭头 | 选择路径、技术链条 |
| `rows` | 编号结论行 + 右侧小图 | 主要结论、创新点 |
| `fig_list` | 左图右“● 标题 + 说明”竖排 | 一组图的逐条解读（参考图 3.1 右栏） |
| `card` | 左图 + 标题 + 红 / 蓝点要点 | 平台卡、负责人卡、样品卡（参考图 3.4） |
| `layers` / `steps_v` / `quadrant` | 层级条 / 竖向步骤 / 2×2 象限 | 体系、操作步骤、优先级 |
| `arrow` / `side_tag` | 浅蓝大箭头 / 竖排深蓝标签 | 推导方向、现状 / 目标 |
| `table` | 表头深蓝、首列浅蓝深蓝粗、全部居中、细框线；`highlights` 单元格关键词红，`highlight_rows` 整行红框 | 方案对比、效益、参数（模板 06 / 参考图 4.2） |
| `kv` | 浅蓝信息条：`items: [["数据版本", "v2"], ["控制条件", "…"]]`，一行排开 | 实验条件、活动要素、比较条件 |
| `numbered` | 深蓝方块编号 + 粗标题 + 说明，行间细线：`items: [{title, text}]` | 汇报安排（作为内容页的目录） |
| `branch` | 左侧根节点（卡点：现象 / 影响 / 已尝试）→ 主干箭头 → 右侧多个假设框 | 当前问题与讨论重点 |
| `row` / `stack` / `grid` | 横排 / 竖排 / 矩阵组合 | 在一个分区里拼复杂结构 |

`matrix` 的单元格写 `{"body": "一句说明"}` 时是浅蓝底、黑字、左对齐的说明块：配 `arrows: "down"` 就是“已有工作 ↓ 尚存不足 ↓ 具体问题”的动机链，不配箭头就是“第一步 / 第二步 / 第三步”的行动计划。

### 图
`figure` 默认无边框、按比例尽量放大。`caption_style: "box"` 把图注放进浅蓝块并贴在格子底部，可加 `caption_title` 做两级图注（“图 1 关键结果示意图 / 简要说明…”）；`overlay: {text, highlights}` 在图下部压一块半透明深蓝说明（高亮变浅黄）。`"file": "@/components/icon-data.png"` 引用技能自带素材（`curve-a/b`、`icon-data / model / deploy / feedback`）。

## 五、版式预设与场景（24 列 × 12 行）

| layout | 槽位 | 典型场景（示例页） |
|---|---|---|
| `info_mosaic` | left、top、bottom | 研究背景与核心需求（1.1） |
| `info_split` | left、right | 两个瓶颈（1.2）、研究问题（1.3）、研究目标（1.4）、成果应用（3.4）、优先级（4.4） |
| `info_split_strip` | left、right、bottom | 两种方案优缺点 + 选择建议（2.1） |
| `info_three` / `info_four` | a–c / a–d | 三个研究方向（4.1）、四项任务（4.2）、推广路径（3.5） |
| `info_three_strip` | a–c、bottom | 三项结果 + 贡献与后续（3.2） |
| `info_quad` | a–d | 四类场景验证（3.3） |
| `info_hub` | lt、lb、center、rt、rb | 四个环节循环（2.6） |
| `info_hero` / `info_hero_left` | main、side | 一张关键图 + 主要发现 / 证据解读 / 研究意义（3.1） |
| `info_side3` | main + s1–s3 | 大图 + 三条分点解读 |
| `info_big` / `info_big_left` | main（17 列）、side（7 列） | **大图页**：论文框架、技术路线、核心结果；**重点图详细解读**（3.1、文献“图 3 详细解读”）；不放标题带 |
| `info_table_fig` | left、right | 效益表 + 指标图（3.6） |
| `info_focus` / `info_stack` / `info_rows3` / `info_top_two` / `info_two_top` / `info_compare` / `info_grid` | 见「参数速查」 | 大图 + 底部说明、思路 + 图带、三行通栏、总图 + 两路径、对照 |
| `info_full` | main | 流程（2.2）、技术路线（2.3）、对比表（2.4）、分层架构（2.5）、计划（4.3）、结论（4.5） |

**按这页要讲的那件事选版式**：

| 要讲的事 | 首选 |
|---|---|
| 为什么做（背景、需求） | `info_mosaic`：图 + 对象 + 关键问题 + 六边形；右侧说明 → 图 |
| 卡在哪（瓶颈、问题） | `info_split` 各带标题带 + 现状 / 目标；或 `label_rows` 现存问题 / 研究思路 |
| 要达到什么（目标） | `sublead` + `links` + `chevrons(navy)` + `process` |
| 怎么做（流程、路线） | `stages`、`matrix`（数据层 / 方法层 / 验证层） |
| 由什么组成（架构） | `matrix` + `side` 支撑列、`cycle`、`layers` |
| 选哪个（方案对比） | `info_split_strip` 优缺点 + 选择建议；`table` + `highlight_rows` |
| 发现了什么（结果） | `info_hero` 图 + 三小节；`info_three_strip` 三结果 + 贡献 |
| 在哪些场景成立 | `info_quad` |
| 怎么落地（应用、推广、效益） | `card` + `overlay`；三栏 `para` + 图 + 图注框；图片标题带 + 表 |
| 什么时候做完（计划） | `timeline` + 阶段目标 / 预期成果 |
| 先做哪个 | `quadrant` + `callout` |
| 得到什么结论 | `rows` 带小图 |

一份 20 页级的稿子至少用 8 种版式，同一版式不连续超过 2 页。

## 五点五、其他汇报场景的页序（示例 `assets/scenes-demo.json`）

| 场景 | 推荐页序与写法 |
|---|---|
| **组会进展** | 封面（`style: "panel"`）→ 汇报安排（`numbered`）→ 进展总表（`table` 上次约定 / 本期完成·证据 / 状态 / 偏差与下一步，偏差行 `highlight_rows`；下方浅蓝块写“本期最重要的变化 / 待确认”）→ 实验记录（`kv` 实验条件 + 三栏 基线 / 本次 / 差异，每栏图 + “观察 / 解释”）→ 当前问题（`branch` 卡点 → 三个假设）→ 下一步计划（三栏 立即推进 / 随后安排 / 需要协助 + 底部“本次需要达成的共识”）→ 结束页（`points` 带走一个结论 / 明确一项行动 / 留出一个问题） |
| **文献汇报** | ① 出发点（一句话说清论文在做什么写进引导句；论文信息 `kv`；“为什么选这篇”“今天想讨论”分点）→ ② 数据来源与研究背景（数据表：来源、时间范围、样本量、变量；背景分点：现状 / 不足 / 为什么值得做）→ ③ 论文框架（`info_big` 大图 + 右侧 `steps_v` 逐步说明，不放标题带）→ ④ 主要结论（`info_three` 每条结论一栏：小标题 + 中图证据 + 分点）→ ⑤ 重点图详细解读（关键 Fig 每张一页：`info_big` 大图 + 两级图注 + 怎么读这张图 / 主要发现 / 对我们的启发）→ ⑥ 我的评价（值得借鉴 / 需要谨慎 + 与我相关 / 下一步尝试 / 讨论问题） |
| **竞选** | 个人介绍（照片 + 我想竞选 + 原因 / 服务对象 + 三个优势小节 + `callout` 承诺）→ 经历证明（`table` 情境 / 行动 / 可核实结果）→ 想解决的三件事（三栏 需求 / 做法 / 反馈 + 预期效果）→ 行动计划（`matrix` 三步 + `body`） |
| **活动策划 / 复盘** | 策划：`kv` 活动要素 → `table` 环节 / 任务 / 负责人 / 资源 → `kv` 备用方案 / 待确认；复盘：原定目标 / 实际结果两小节 → `table` 差异原因 / 改进动作 / 责任与检查点 |

封面两种：默认居中式（顶部 / 底部深蓝色带）；`"style": "panel"` 是左侧深蓝竖栏（`panel_title`、`tagline`、`figure`）+ 右侧左对齐题目（`kicker`、`title`、`subtitle`、`info`、`note`）。结束页两种：默认“感谢聆听”；写 `points`（≤ 3 条 `{title, text}`）是“回到本次汇报的核心”+ 浅蓝致谢条。

## 五点八、参数速查（不用读脚本）

### 文件与图
- **路径**：`file` 的相对路径按**内容 JSON 所在目录**解析（不是命令行当前目录）；`@/components/xxx.png` 指技能自带素材。`prep_figs.py` 的输出 JSON 要和输入放在同一目录。
- **缺图**：图位只写 `prompt`（+ `aspect`，如 `"16:9"`，只写进提示词文件，不影响版面）和 `label`；占位框上显示“预留图位 + label”，提示词导出到 **PPTX 输出目录下的 `images/<label>.prompt.md`**。占位框只用于草稿，交付前换成真图或改版式。
- **角色**：`role: "big"`（框架图、技术路线、核心结果、架构、平台主界面）放不到宽 7.5 in 或高 4 in 时日志报警；`role: "small"` 只作提示，表示不承载要读的细节。
- **图注**：不写 `caption` 就没有图注（`label` 只是名字）。`caption_style: "box"` 图注进浅蓝块；`caption_title` 两级图注；`caption_color: "red" / "navy"`；`crop: [左, 上, 右, 下]`（比例）裁子图；`fit: "cover"` 裁满格子。

### 文字类组件
| type | 参数 |
|---|---|
| `section` | `head`、`items`（字符串或 `{label, text, highlights}`）、`text`（代替要点的一段话）、`num`、`head_style: "fill" / "bar"`、`size` |
| `markers` | `items`、`fill: "light"`、`after`（段后 pt）、`size` |
| `note` | `title`、`text`（`\n` 分行）、`highlights`、`strong`、`tone: "red" / "plain"`、`dashed`、`align: "left"`、`fill`、`line`、`bold`、`title_color: "red"` |
| `para` | `text`、`color: "ink" / "navy" / "red" / "white" / "grey"`、`align: "left"`、`bold`、`size`、`highlights`、`strong` |
| `chevrons` | `items: [{title, text, highlights}]`、`tone: "navy"` |
| `label_rows` | `items: [{label（可 \n）, text, highlights}]`、`label_color: "navy"`、`label_w`、`gap` |
| `links` | `items: [["左短语", "右关键词"]]`、`split`（左框占比，默认 0.52） |
| `kv` | `items: [["键", "值"]]`、`fill`（默认浅蓝，可写 `"FFFFFF"`） |
| `numbered` | `items: [{title, text, highlights}]`、`title_w`（标题列占比 0.32）、`size`、`text_size` |
| `callout` | `text`、`highlights`、`max_size`（默认 18） |
| `pill` | `text`、`full`、`light`、`size` |
| `hexes` / `tags` | `items`（可 `\n`）、`max_w` / `tag_h` |
| `side_tag` | `text`（每字一行竖排，2–4 字）——放进 `row`：`{"type": "row", "widths": [0.5, 5], "items": [side_tag, note]}` |
| `arrow` | `direction: right / left / down / up`、`label`、`size`（默认 0.55 in）——放进 `row` 当两块之间的推导箭头 |

### 结构类组件
| type | 参数 |
|---|---|
| `stages` | `items: [{title, figure, items, text, highlights, output}]`、`head_style: "light"`、`frame: false`、`arrow_w` |
| `matrix` | `rows: [{label, sub, cells, notes}]`；单元格：文字 / `{title, text}` / `{body}` / `{figure, caption}` / `{figure, items}`；`columns`（表头，第一个对行标签列）、`arrows: "down" / "up"`、`top`、`bottom`、`side: {title, items: [{text, sub}]}`、`label_style: "navy" / "light"`、`label_w`（默认 1.75 in）、`notes_w`（2.6）、`gap_x`（0.25）、`row_frame` |
| `timeline` | `items: [{stage, title, time, blocks: [{head, items}], text, figure}]`、`fig_share`（有图时图占比 0.45） |
| `branch` | `root: {title, items: [{label, text}], text}`、`items: [{title, text, highlights}]`、`root_w`（0.24）、`head_w`（0.42） |
| `cycle` | `items`（从左上起顺时针，可 `\n`）、`center`、`colors`、`size`、`center_size` |
| `process` | `items`、`style: "pill"`、`gradient`、`colors`、`bar_h`；每块文字太长时自动缩字并在日志提示 |
| `rows` | `items: [{title, items（分点）, text, highlights, figure}]`、`gap` |
| `fig_list` | `items: [{figure, title, text, highlights}]`、`fig_w`（0.36） |
| `card` | `figure`、`title`、`items`、`text`、`fig_w`（0.34）、`items_red`（默认 true） |
| `layers` / `steps_v` / `quadrant` | `items: [{title, text}]`；`pyramid: false`；`x_label`、`y_label`、`fills` |
| `table` | `columns`、`rows`（每行列数要全）、`widths`、`align`（每列 `"center"/"left"`）、`highlights`、`strong`、`highlight_rows`（**从 1 数正文行，不含表头**）、`first_col: false`、`size`；默认撑满分区 |
| `chart` | `kind: line / bar / stacked / barh`、`categories`、`series: [{name, values}]`、`title`、`x_title`、`y_title`、`min`、`max`、`labels`、`size` |
| `formula` | `lines: ["T = T_0 + \\Delta T"]`、`latex: true`（LaTeX 子集：`\\frac`、上下标、`\\sqrt`、希腊字母、`\\sum`）、`size`（默认 18）、`align`、`blue` |

### 组合与通用
- `row`：`items`、`widths`（比例）、`gap`、`h`（固定高度，不写则撑满）；`stack`：`items`、`gap`、`valign: top / middle / spread`。
- 任何条目都可写 `h`（固定高度）或 `weight`（分剩余高度的权重）。
- 页标题 ≤ 14 字（不含编号；英文字母按半个字算）。

### 版式槽位（24 列 × 12 行，`[起始列, 列数, 起始行, 行数]`）
| layout | 槽位 |
|---|---|
| `info_full` | main [0,24,0,12] |
| `info_focus` | main [0,24,0,9]、bottom [0,24,9,3] |
| `info_big` / `info_big_left` | main [0,17,0,12]、side [17,7,0,12]（big_left 左右互换） |
| `info_hero` / `info_hero_left` | main 14 列 + side 10 列（hero_left 左右互换） |
| `info_side3` | main [0,14,0,12]、s1 / s2 / s3 [14,10,0/4/8,4] |
| `info_split` | left [0,12,0,12]、right [12,12,0,12] |
| `info_split_strip` | left [0,12,0,9]、right [12,12,0,9]、bottom [0,24,9,3] |
| `info_three` / `info_three_strip` | a / b / c 各 8 列（strip 版高 9 行 + bottom [0,24,9,3]） |
| `info_four` | a–d 各 6 列 |
| `info_quad` | a [0,12,0,6]、b [12,12,0,6]、c [0,12,6,6]、d [12,12,6,6] |
| `info_rows3` | r1 / r2 / r3 各 4 行通栏 |
| `info_mosaic` | left [0,9,0,12]、top [9,15,0,6]、bottom [9,15,6,6] |
| `info_top_two` | top [0,24,0,6]、bl [0,12,6,6]、br [12,12,6,6] |
| `info_two_top` | tl [0,12,0,6]、tr [12,12,0,6]、bottom [0,24,6,6] |
| `info_stack` | top [0,24,0,5]、bottom [0,24,5,7] |
| `info_compare` | left [0,11,0,12]、mid [11,2,0,12]（无框，放箭头）、right [13,11,0,12] |
| `info_hub` | lt [0,9,0,6]、lb [0,9,6,6]、center [9,6,0,12]（无框）、rt [15,9,0,6]、rb [15,9,6,6] |
| `info_grid` | main [0,16,0,12]、side [16,8,0,12] |
| `info_table_fig` | left [0,13,0,12]、right [13,11,0,12] |

预设槽位都可以用 `grid` 覆盖，也可以新增任意名字的槽位（写 `grid` 即可）。流程条、深蓝标签这类细长组件至少给 2 行网格。

## 六、写页清单

- 先写标题带那一句，再选版式，最后填组件；和这句无关的内容删掉。
- 每个分区“图 + 小标题 + 2–3 条要点”是基本单元；要点一行说完（窄栏 ≤ 12 字），不写成段落。
- 图的比例按格子算：半栏约 6 × 3.4 in（16:9），三栏约 4 × 2.2 in，矩阵单元格约 2.4 × 1.1 in；GPT 生图时把比例写进提示词。
- 演示用的合成数据在图题写“示意数据”；正式稿只放真实数据。
- 生成后跑 `lint.py`，再逐页看渲染图：留白、孤行、红字数量、图是否够大。
