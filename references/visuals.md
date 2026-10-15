# 视觉素材与空间填充

## 一、选图顺序

1. **用户提供的材料**：论文原图（用 `paper_figure`，保持原色，可裁出子图）、照片、软件截图、用户自己的数据图。
2. **GPT 生图（可选）**：背景、研究对象、场景、机制示意、技术路线配图、展望配图，都可以用 GPT 生成。前提是提示词按第三节去掉“AI味”，每张图都有明确的论证用途。
3. **真实参考图**：论文、官方报告、统计年鉴里的图，图注写“图源：…”。
4. **原生图表和公式**：趋势、对比必须用真实数据画原生 `chart`；模型用公式。

**任何时候都不强求配图**：用户说不用图、或这一稿只是把版式跑通时，一张 AI 图都不要放，直接把图位删掉或换成原生图表 / 公式 / 表格 / `flow` / `stat` 把分区填满。图好不好看是次要的，别让配图拖住交付。

两条底线：
- 结果页的证据只能是数据、真实图或原生图表。GPT 图可以作辅助（例如场景说明），不能当作结果。
- 数据图不用 GPT 生成。

## 关于图里的文字：不禁止，按需要写

图是让人看懂的，**需要标签就写标签**，不要一概禁止文字或数字。做法是：
- 要写就在提示词里**逐字给出**要出现的字（宁可少、不要错），并加上“汉字必须正确、工整，不要多余的字、不要乱码”，每张只放 1–4 个短标签；长句仍然容易写坏。
- 不需要文字时，也不必特意写“不准出现任何文字”——这条限制对理解没有帮助，反而会让画面变空。
- 图里的标签与 PPT 正文用词保持一致；关键标签也可以由图内文字承担，不一定要改成 PPT 原生文本。

## 二、少用“方框带字”

**一份稿里的配图要刻意拉开“类型”，不要都是同一类画法。** 实测教训：自己写三条提示词，出来三张“左侧设备 + 中部框 + 右侧箭头”的雷同图（用户一眼看出“你这三个图完全没区别”）。
把提示词交给网页端写就没这个问题：它自己把三张排成三类——**概念图**（只讲机制，不画设备内部）、**原理图**（讲能流转换与方向）、**平面总平面图**（讲实体空间与管网）。
所以：需要 2 张以上配图时，**先让网页端写提示词**（告诉它每张图要说什么、图位比例，其余让它定），再原样贴到另一个新会话里生图。

方框流程、关键数字卡、并列的标签框叠多了，页面就像模板，不如原稿高级。
- `flow` 默认是无框样式：文字之间用蓝色箭头连接。`style: "box"` 只用来画真正的多步流程，每页最多1个。
- `stat` 默认把数字直接写在灰底上，不加卡片；整页最多一排。
- 能用一张好图说明的，就不要用三个方框说。

## 三、空间填充：分区下方不留大片空白

生成时出现 `region ... has X in empty`、或 lint 报“灰底分区下方留白”时，按顺序处理：

1. **补一张配图**：GPT 生图或真实参考图，放进该分区，图片用 `fit: contain` 按比例贴合。分区宽而矮时，生成横幅比例的图；窄而高时，生成竖版或方形的图。
2. **补数据或公式**：一张小的原生图表（真实或标“示意”的数据，高度不低于1.5英寸），或公式加变量说明。
   **不要用“大数字 + 下面一行小字”的数字块**（用户说很丑）：关键数字写进正文句子里用红色强调，或进表格。
3. **调整结构**：缩短或合并分区，换更紧凑的版式，让主图更大。

不要放大字号硬撑，也不要用一排方框凑满。

## 四、GPT 生图：去掉“AI味”的提示词

### 先避开这些“AI味”特征
光滑塑料质感、等距（isometric）微缩城市、发光线条和粒子、全息界面、霓虹蓝紫渐变、镜头光晕、戏剧化天空、闪电、过度锐化和高饱和、对称居中的“海报式”构图、元素堆满画面、不合常理的设备结构、乱码文字。

### 写提示词的原则
- 写清**具体对象和关系**（是什么设备、几台、怎么连接、从什么视角看），不要写“未来感”“科技感”“智慧”这类空泛词。
- 指定**画风参照**：期刊插图、技术手册线稿、纪实摄影，而不是“插画”“3D 渲染”。
- 指定**色彩上限**：哑光、低饱和，只用2–3个颜色，与PPT的蓝 #1D50A2、灰、少量红协调。
- 指定**构图**：主体偏一侧或横向展开，边缘留白；画幅比例与图位一致。
- 写明**负面约束**：不要水印和 logo、不要发光、不要渐变、不要等距3D。（**不要**写“不准出现任何文字”——需要标签就逐字给出要写的字。）

### 模板 A：科研示意插图（机制、系统、流程配图）

> Technical illustration for a research presentation: [具体对象与关系]. Style of a figure in an engineering journal or a technical manual: clean 2D side/front view, precise thin outlines, flat matte fills, realistic proportions, restrained detail. Muted palette limited to slate blue (#1D50A2), warm gray and a small amount of brick red. Plain white background, generous margins, subject arranged horizontally. [需要标签时：“Keep only these labels, exactly: “风电”“光伏”“市场出清”——汉字要正确工整，不要多余的字符。”；不需要就整句删掉] No logos, no watermarks. No isometric view, no 3D render, no glossy or plastic look, no glow, no gradients, no particles, no holograms, no dramatic lighting. Aspect ratio [与图位一致].

### 模板 B：纪实照片（真实场景、设备、气候事件）

> Documentary photograph of [具体场景，如 "a row of rooftop air-conditioner condensers on an old residential building in a hot humid city at noon"]. Shot on a 35 mm lens at eye level, overcast natural daylight, true-to-life muted colours, ordinary real-world details and slight imperfections, natural depth of field. Looks like a photo from a news agency or a field survey, not a render. No people posing, no text, no logos, no HDR, no oversaturation, no cinematic sky, no lens flare. Aspect ratio [与图位一致].

### 模板 C：概念示意（只有一个核心概念时使用）

> Minimal editorial illustration explaining [一个概念]，drawn like a figure in The Economist or Nature news: simple shapes, thin lines, flat colours, lots of white space, one accent colour (#C00000) on the key element, the rest slate blue and gray. [需要标签就逐字给出；不需要就不用写“no text”] No 3D, no glow, no gradients. Aspect ratio [与图位一致].

### 生成后的自查（任一项不过关就改提示词重画）
1. 对象数量、连接方向、尺度是否符合物理常识和研究内容？
2. 有没有出现第一小节列出的“AI味”特征？
3. 和灰底分区、原生图表放在一起，配色和风格是否协调？
4. 缩小到图位大小后，主体是否还能看清？**图里的标签另有算法**：图位宽 ÷ 图片像素宽 × 标签字高 ≥ 9–10 pt 才算可读。
   图位很窄（宽度 < 2.5 英寸）时不要在图里写标签，用不带文字的版本，靠图注和正文说明（否则标签只有 6–7 pt，白写）。
5. 图里写了文字的话，汉字是否逐字正确、没有多余的字或乱码？写不好就把字拿掉、改由 PPT 原生文本补。

### 内标签实测（2026-10-14，gpt-image）

逐字给出 3–4 个短标签时，模型能全部写对：实测三张（「天气 / 新能源出力 / 市场供给」、「电价 / 负荷 / 气象 / 15 分钟网格」、
「新能源 / 常规机组 / 市场出清 / 负荷」）**汉字全部正确、无多余字**。数字、坐标刻度同样可以交给它写（用户也确认不用限制）。
写长句仍然会坏——每张只放短标签。

图注写“AI生成示意”；提示词存到 `images/<名称>.prompt.md`，便于复现和修改。
