# 版式模板：只填内容，不写坐标

`dense_deck.py` 支持在页面上写 `"layout": "<名称>"` 和 `"regions": {槽位: {...}}`。区域的位置取自26页密度样例，脚本负责画面板、把内容自上而下排好。完整示例：`assets/layout-demo.json`（7页，包含下面大部分写法）。

## 可用版式

| layout | 槽位（默认底色） | 对应原稿页 | 适合 |
|---|---|---|---|
| `background` | q1、q2（灰），aim、evidence（无），figure（无），side（灰） | 3 | 背景与问题：两个问题区，一句目标加一张小统计图，右侧配图加目标面板 |
| `review_table` | scope（灰），main（无） | 4 | 研究现状：左边定义和范围，右边对照表、流程、研究缺口 |
| `framework5` | m1–m5（灰） | 6 | 技术路线五模块；模块里可以放小图、公式、小表 |
| `three_columns` | left、center、right（灰） | 7、9 | 三个方向，或“方法列表 / 公式 / 示意曲线” |
| `two_columns` | left、right（灰） | 10、13、15 | 原理对流程，输入对模型 |
| `results_side` | main（无），side（灰） | 19、23、25 | 多图矩阵加右侧解读 |
| `results_bottom` | main（无），left、right（灰） | 20、21 | 双图对照加底部两个条件性结论 |
| `case_setup` | left（灰），right（无） | 18 | 算例对象或拓扑，加参数表和实验条件 |
| `full` | main（无） | — | 整页大图或大表 |
| `conclusions` | 用 `cards`（2–4张），不用 regions | 28 | 编号结论卡：标题加两行说明 |

封面（`type: cover`）、目录（`type: agenda`，`active` 为当前章，已讲过的章节自动变灰）的写法不变。

## 区域写法

```json
"regions": {
  "center": {
    "title": "干球温度：平移 + 伸缩",
    "fill": "gray",
    "items": [
      {"type": "text", "text": "……", "size": 16, "highlights": ["日较差"]},
      {"type": "formula", "latex": true, "size": 18, "blue": true, "h": 1.75,
       "lines": ["T = T_0 + \\Delta T_m + \\alpha_m (T_0 - \\langle T_0 \\rangle_m)"]},
      {"type": "chart", "kind": "line", "min": 20, "categories": ["0","3"], "series": [...]}
    ]
  }
}
```

- `title`、`fill` 可以省略，省略时用版式默认值：所有分区统一灰底 `#F2F2F2`，页底结论条（页面的 `takeaway`）用浅橙。`fill` 也接受 `white` / `none` / 十六进制色，但除非用户要求，保持统一灰底。`box` 可以覆盖槽位位置，`gap` 调整条目间距。
- **高度规则**：
  - 文本、公式、表格按内容自动量高度；
  - 图表、图片、流程图、`row`、`grid`、`stack` 平分剩下的空间，用 `weight` 调整比例；
  - 写了 `h`（英寸）就固定为这个高度。
  - 生成时出现 `region over-full` 提示，说明内容放不下：删减内容或拆页，不要缩小字号。
- **留白**：区域内没有可伸缩条目、内容又放不满时，生成会提示 `region '...' has X in empty`。处理办法见 [visuals.md](visuals.md) 第二节。`valign: "middle" / "spread"` 只用来处理少量余量。可伸缩条目分到的高度低于下限（图表1.5英寸、图片1.2英寸）时会报警。
- 版式没有定义的额外槽位（例如在 `three_columns` 下面加一个 `bottom`），必须写 `box`，默认也是灰底。
- **组合条目**：
  - `{"type":"row","items":[...],"widths":[2,1]}`：横向并排；
  - `{"type":"grid","cols":2,"items":[...]}`：n列矩阵，结果页常用；
  - `{"type":"stack","items":[...]}`：在格子里再竖向堆叠，比如“小标题 + 表格”。
- 同一页可以同时写 `regions` 和 `blocks`。`blocks` 用绝对坐标，适合补一个跨区箭头或标注。

## 条目类型速查

| type | 关键字段 |
|---|---|
| `text` | `text`、`size`、`color`、`bold`、`highlights`（红色加粗）、`strong`（黑色加粗）、`bullet: true`（参考稿的“•”要点列表，130%行距）、`spacing`、`align: center` |
| `chart` | `kind`（line / bar / stacked / barh）、`categories`、`series[{name, values, color}]`、`min`、`max`、`major`（刻度间隔）、`x_title`、`y_title`、`labels` + `label_format`、`number_format`、`legend`、`size`（默认12，不要低于12） |
| `table` | `columns`、`rows`、`widths`、`size`、`align`（默认全部居中，水平和垂直都居中；需要例外时按列给 left / center）、`emphasis: [[行,列]]`（标红；行从1数，表头不算；列从0数）、`first_col_bold` |
| `formula` | `lines`、`latex`、`size`、`blue`、`align` |
| `diagram` | `nodes[{id,label,x,y,w,h,decision,circle,fill}]`（坐标是0–1的相对比例，节点默认白底蓝框）、`edges`（空列表表示并列、不画箭头）、`arrows: false` 只画线不画箭头 |
| `figure` | `file`、`fit`（contain / cover）、`caption`、`annotations`、`label`、`valign`（middle / top）、`prompt`、`aspect`。contain 时图位按图片比例收缩，细边框、图片、图注紧贴；不给 `file` 时生成预留图位；写了 `prompt` 时，生成脚本会把提示词存到输出目录的 `images/<label>.prompt.md`，并列出待生成清单 |
| `flow` | `items`（文字，或 `{label, bold, color}`）、`direction`（right / down）、`style`（默认 `plain`：文字加蓝色箭头、无框；`box`：带框，只用于真正的多步流程，每页最多1个）、`arrows`、`size` |
| `stat` | **不建议再用**（用户说“一个数字下面跟几个字很丑”）：关键数字写进正文句子里或表格里。字段仍在：`items[{value, label, emphasis}]`、`size`（默认26）、`h`（默认0.85）、`card`（默认 false） |
| `stack` | `items`、`valign: "middle"`（整块垂直居中）、`gap` |
| `cards` | **并列卡片**：白底圆角 + 柔和阴影，标题与正文都在框内居中。`items:[{title, body, highlights}]`、`direction`（right/down）、`gap`、`title_size`、`size`。弹性条目，放进 region 会自己吃满剩余高度；只放短内容 |
| `paper_figure` | 论文原图：没有灰框，保持原比例。`crop: [左,上,右,下]` 是占原图的比例，PPT原生裁切，原图文件不受影响；`tag: "(a)"` 标子图号 |

## LaTeX 子集（`"latex": true`）

支持：`\frac{}{}`、`_`、`^`、`\sqrt{}`、`\hat` / `\bar` / `\tilde` / `\dot` / `\vec` / `\overline`、`\text{}` / `\mathrm{}`、希腊字母、`\sum \prod \int \partial \infty \cdot \times \pm \le \ge \ne \approx \in \to \langle \rangle`，以及 `\max \min \log \exp \ln \sin \cos` 等正体函数名。`\left` 和 `\right` 会被忽略。其他命令按字面输出，所以写完要看渲染。

生成的是原生Office公式，外面包了一层标准的 `mc:AlternateContent`：PowerPoint显示真正的分式；LibreOffice和WPS显示线性的Unicode备用文字。
