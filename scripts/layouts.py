"""Named page layouts with slots, distilled from the 26-page density demo.

A slide may use `"layout": "<name>"` plus `"regions": {slot: {...}}`. Each region draws its
panel (title + fill) and stacks its `items` top to bottom, so content JSON never needs
coordinates. Items are ordinary blocks without `box`; give `h` (inches) to fix a height,
otherwise text/formula/table heights are measured and charts/figures/diagrams share the
rest (use `weight` to bias). `{"type": "row"}` splits horizontally, `{"type": "grid"}`
makes an n-column matrix, `{"type": "stack"}` nests a vertical group (e.g. label + table in a grid cell). A region may override its slot with `box`.
"""
import math
from textfit import measure

FILLS = {'gray': 'F2F3F5', 'pale': 'E3EFF9', 'orange': 'FCEBDD', 'white': 'FFFFFF'}

# slot: (box [x, y, w, h] in inches, default fill or None)
LAYOUTS = {
    # ref p3 – two problem panels, statement + small evidence chart, figure + target panel
    'background': {
        'q1': ([0.5, 1.45, 8.0, 1.72], 'gray'), 'q2': ([0.5, 3.3, 8.0, 1.68], 'gray'),
        'aim': ([0.5, 5.12, 4.2, 1.45], None), 'evidence': ([4.85, 5.1, 3.65, 1.5], None),
        'figure': ([8.78, 1.45, 4.2, 2.3], None), 'side': ([8.78, 3.95, 4.2, 2.6], 'gray')},
    # ref p4 – scope/definition panel + comparison table + gap statement
    'review_table': {'scope': ([0.5, 1.46, 3.1, 5.05], 'gray'), 'main': ([3.84, 1.46, 9.16, 5.05], None)},
    # ref p6 – five-module framework, modules may hold small charts/figures/formulas
    'framework5': {
        'm1': ([0.5, 1.5, 2.62, 4.95], 'gray'), 'm2': ([3.25, 1.5, 3.17, 2.55], 'gray'),
        'm3': ([3.25, 4.21, 5.85, 2.24], 'gray'), 'm4': ([6.55, 1.5, 2.55, 2.55], 'gray'),
        'm5': ([9.23, 1.5, 3.77, 4.95], 'gray')},
    # ref p7/p9 – three directions or model comparison columns
    'three_columns': {'left': ([0.5, 1.5, 3.55, 5.06], 'gray'), 'center': ([4.23, 1.5, 4.8, 5.06], 'gray'),
                      'right': ([9.2, 1.5, 3.8, 5.06], 'gray')},
    # ref p10/p13/p15 – principle vs process, input vs model
    'two_columns': {'left': ([0.5, 1.5, 6.15, 5.06], 'gray'), 'right': ([6.85, 1.5, 6.15, 5.06], 'gray')},
    # ref p19/p23/p25 – chart matrix + explanation column
    'results_side': {'main': ([0.5, 1.5, 8.65, 5.1], None), 'side': ([9.3, 1.5, 3.7, 5.1], 'gray')},
    # ref p20/p21 – paired charts + two conditional conclusions underneath
    'results_bottom': {'main': ([0.5, 1.5, 12.5, 3.45], None), 'left': ([0.5, 5.08, 6.15, 1.48], 'gray'),
                       'right': ([6.85, 5.08, 6.15, 1.48], 'gray')},
    # ref p18 – object/topology + parameters and experiment conditions
    'case_setup': {'left': ([0.5, 1.5, 6.6, 5.06], 'gray'), 'right': ([7.3, 1.5, 5.7, 5.06], None)},
    'full': {'main': ([0.5, 1.5, 12.5, 5.06], None)},
    # 用户稿《智能体、Skill 与 MCP》的主力页型（25 页里 10 页）：
    # 左侧 16:9 大图 + 一句图注，右侧窄栏“看这张图看什么”3–4 条，侧栏灰底按内容收高
    'figure_side': {'main': ([0.8, 1.5, 8.05, 5.1], None), 'side': ([9.3, 1.5, 3.7, 5.1], 'gray')},
    # 录屏 / 截图页：素材占主体，侧栏写“看什么”
    'video_side': {'main': ([0.5, 1.5, 8.6, 5.2], None), 'side': ([9.3, 1.5, 3.7, 5.1], 'gray')},   # 不写图注时素材 8.6×4.85
    # 文字在左、图在右（介绍对象 / 工具时用）
    'figure_right': {'left': ([0.5, 1.5, 6.0, 5.06], 'gray'), 'right': ([6.7, 1.5, 6.3, 5.06], None)},
}
FLEX = {'chart', 'figure', 'paper_figure', 'diagram', 'row', 'grid', 'stack', 'cards', 'stages', 'timeline', 'rows', 'side_tag', 'layers', 'steps_v', 'quadrant'}   # stat / flow have natural heights
GAP = 0.12


CURRENT_PAGE = ''   # 生成日志里标出是哪一页、哪个分区
CURRENT_REGION = ''


def natural_height(item, width):
    """Height for content-sized items; None for flexible ones."""
    if 'h' in item:
        return item['h']
    if item.get('grow'):
        return None
    kind = item['type']
    if kind == 'text':
        return measure(item['text'], item.get('size', 16), width, item.get('bold', False),
                       item.get('bullet', False), item.get('spacing'), item.get('sub', False))[1]
    if kind == 'formula':
        return len(item['lines']) * item.get('size', 18) * 1.9 / 72 + 0.08
    if kind == 'table':
        size = item.get('size', 14)
        rows = [item['columns']] + item['rows']
        weights = item.get('widths') or [1] * len(rows[0])
        total = 0
        for row in rows:
            lines = max(len(measure(str(v), size, width * wt / sum(weights) - 0.2)[0])
                        for v, wt in zip(row, weights))
            total += lines * size * 1.25 / 72 + 0.13
        return total
    from info_style import NATURAL
    if kind in NATURAL:
        return NATURAL[kind](item, width)
    if kind == 'stat':
        raise ValueError('stat 数字块已停用：把关键数字写进正文句子或表格')
    if kind == 'transition':
        return item.get('h', 0.75)
    if kind == 'flow':
        n = len(item['items'])
        size = item.get('size', 15)
        if item.get('style', 'plain') == 'plain':
            if item.get('direction', 'right') == 'right':
                return size * 1.5 / 72 + 0.12
            lines = 2 * n - 1 if item.get('arrows', True) else n
            return lines * size * 1.3 / 72 + 0.1
        if item.get('direction', 'right') == 'right':
            return item.get('box_h', 0.55) + 0.07
        return n * item.get('box_h', 0.5) + (n - 1) * item.get('gap', 0.28) + 0.04
    return None


def stack(items, x, y, w, h, place, gap=GAP, valign='top', label=''):
    """Stack items top to bottom. Returns the unused height (bottom white space)."""
    fixed = [natural_height(item, w) for item in items]
    flexible = [i for i, v in enumerate(fixed) if v is None]
    used = sum(v for v in fixed if v is not None) + gap * (len(items) - 1)
    spare = h - used
    if spare < -0.05:
        print(f'  ! [{CURRENT_PAGE} 分区 {CURRENT_REGION}] {label or "内部组合"} 超出 {-spare:.2f} in（over-full）：删减内容、换更大的格子或拆页')
    weights = sum(items[i].get('weight', 1) for i in flexible)
    cursor = y
    free = 0.0 if flexible else max(spare, 0.0)
    # spread 但只有一个多行文本：条目间没有可分的空隙，改为加大行距撑满（上限 2.2 倍）
    if free and valign == 'spread' and len(items) == 1 and items[0]['type'] == 'text' and '\n' in items[0]['text']:
        item = items[0]
        base = item.get('spacing') or (1.3 if (item.get('bullet') or item.get('sub')) else 1.18)
        target = min(2.2, base * h / max(fixed[0], 0.1))
        items = [dict(item, spacing=target, h=h)]
        fixed = [h]
        free = max(0.0, h - natural_height(dict(item, spacing=target), w))
    if free and valign == 'middle':
        cursor += free / 2
    elif free and valign == 'spread' and len(items) > 1:
        gap += min(free / (len(items) - 1), 0.45)
        free = max(0.0, h - (used - GAP * (len(items) - 1)) - gap * (len(items) - 1))
    if label is not None and free > max(0.6, 0.25 * h):
        print(f'  · [{CURRENT_PAGE}] {label} 内容下方空了 {free:.2f} in（{free / h:.0%}，has empty）；'
              f'先放大图，再补具体的图 / 小标题要点 / 表格，或收窄分区')
    minimum = {'chart': 1.5, 'figure': 1.2, 'paper_figure': 1.2, 'diagram': 1.0, 'grid': 1.5}
    for item, height in zip(items, fixed):
        if height is None:
            height = max(spare, 0.4) * item.get('weight', 1) / weights if weights else 0.4
            if height < minimum.get(item['type'], 0) - 0.01:
                print(f"  ! [{CURRENT_PAGE} 分区 {CURRENT_REGION}] {item['type']} 只分到 {height:.2f} in 高（< {minimum[item['type']]} in），太小看不清——"
                      f"挪到更大的格子或删掉同格文字")
        place(item, x, cursor, w, height)
        cursor += height + gap
    return free


def place_factory(slide, block):
    def place(item, x, y, w, h):
        kind = item['type']
        if kind == 'row':
            widths = item.get('widths') or [1] * len(item['items'])
            g = item.get('gap', 0.15)
            avail = w - g * (len(widths) - 1)
            cx = x
            for child, weight in zip(item['items'], widths):
                cw = avail * weight / sum(widths)
                place(child, cx, y, cw, child.get('h', h))
                cx += cw + g
        elif kind == 'stack':
            stack(item['items'], x, y, w, h, place, item.get('gap', GAP), item.get('valign', 'top'))
        elif kind == 'grid':
            cols = item.get('cols', 2)
            rows = math.ceil(len(item['items']) / cols)
            g = item.get('gap', 0.15)
            cw = (w - g * (cols - 1)) / cols
            ch = (h - g * (rows - 1)) / rows
            for n, child in enumerate(item['items']):
                place(child, x + (n % cols) * (cw + g), y + (n // cols) * (ch + g), cw, ch)
        else:
            spec = dict(item)
            if kind in ('figure', 'paper_figure') and spec.get('caption', '') != '' \
                    and not spec.get('caption_style'):
                h = h - 0.58          # figure() draws its caption below the box
            spec['box'] = [x, y, w, h]
            block(slide, spec)
    return place


def render_regions(slide, spec, block, panel):
    layout = LAYOUTS.get(spec['layout'])
    if layout is None:
        raise ValueError(f"Unknown layout {spec['layout']}; choose from {sorted(LAYOUTS)} or use absolute blocks")
    place = place_factory(slide, block)
    for slot, region in spec.get('regions', {}).items():
        if slot not in layout and 'box' not in region:
            raise ValueError(f"Layout {spec['layout']} has no slot '{slot}'; slots: {sorted(layout)}")
        default_box, default_fill = layout.get(slot, (None, 'gray'))   # extra slots default to the unified gray
        x, y, w, h = region.get('box', default_box)
        fill = region.get('fill', default_fill)
        style = region.get('title_style', 'plain')
        title = region.get('title', '')
        # fit: 灰底按内容收高，不留半截空灰底（figure_side / video_side 的侧栏默认开启）
        if region.get('fit', spec['layout'] in ('figure_side', 'video_side') and slot == 'side'):
            pad_ = region.get('pad', 0.2)
            heights = [natural_height(it, w - 2 * pad_) for it in region.get('items', [])]
            if heights and None not in heights:
                top_ = (0.62 if style in ('band', 'tab') else 0.55) if title else 0.18
                need = top_ + sum(heights) + region.get('gap', GAP) * (len(heights) - 1) + 0.3
                h = min(h, max(need, 1.2))
        fill = FILLS.get(fill, fill) if fill not in (None, 'none') else None
        if fill or title:
            panel(slide, {'type': 'panel', 'title': title, 'box': [x, y, w, h], 'fill': fill or 'FFFFFF',
                          'size': region.get('title_size', 17), 'title_style': region.get('title_style', 'plain')})
        boxed = bool(fill or title)
        pad = region.get('pad', 0.2 if boxed else 0.0)
        top = (0.62 if style in ('band', 'tab') else 0.55) if title else (0.18 if boxed else 0.0)
        bottom = 0.15 if boxed else 0.0
        stack(region.get('items', []), x + pad, y + top, w - 2 * pad, h - top - bottom, place,
              region.get('gap', GAP), region.get('valign', 'top'), f"region '{slot}'")


def render_conclusions(slide, spec, text_fn, panel, blue):
    cards = spec['cards']
    if not 2 <= len(cards) <= 4:
        raise ValueError('conclusions layout takes 2–4 cards; split the page otherwise')
    gap = 0.2
    height = (5.05 - gap * (len(cards) - 1)) / len(cards)
    for i, card in enumerate(cards):
        y = 1.55 + i * (height + gap)
        panel(slide, {'type': 'panel', 'title': '', 'box': [1.3, y, 11.47, height], 'fill': 'F2F3F5'})
        text_fn(slide, str(i + 1), 0.63, y + height / 2 - 0.25, 0.48, 0.45, 24, blue, True)
        text_fn(slide, card['title'], 1.56, y + 0.14, 10.91, 0.41, 19, blue, True)
        text_fn(slide, card['body'], 1.56, y + 0.62, 10.91, height - 0.7, card.get('size', 16),
                highlights=card.get('highlights'))
