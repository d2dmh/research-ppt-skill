# 安装与依赖说明（research-ppt-skill）

这是一份**给 AI agent 用的技能包**（Hermes Agent / Claude Code / Codex 都能加载）。
`SKILL.md` 是给 agent 看的操作手册，`scripts/` 是真正干活的 Python 脚本，`references/` 是按需展开的细节。

做出来的是**蓝白信息图式、原生可编辑的 PPTX**（不是图片拼的）：视觉规范和版式库写在脚本与
`references/info-style.md` 里，内容换成你给的材料。示例见 `assets/info-demo.json`。

---

## 1. 装到哪里

**Hermes Agent**

```bash
mkdir -p "$HERMES_HOME/skills/productivity"        # 默认 $HERMES_HOME = ~/.hermes
cp -R research-ppt-skill "$HERMES_HOME/skills/productivity/"
hermes skills list | grep research-ppt              # 确认已加载
```

放好后新开一个会话即可生效；不用重启网关。也可以直接 `git clone` 一份到该目录。

**Claude Code / Codex 等其他 agent**

把整个 `research-ppt-skill/` 放进项目的 `.hermes/skills/` 或 `.agents/skills/`，或者干脆在会话里
说「读 `research-ppt-skill/SKILL.md`，按它执行」——SKILL.md 本身就是自足的手册。

## 2. 依赖（一次装好，长期复用）

| 用途 | 需要什么 | 说明 |
|---|---|---|
| 生成/改 PPTX | `python-pptx` | 本包已按 1.0.2 验证 |
| 讲稿 Word（逐字稿） | `python-docx` | |
| 改 XML / 修排版 | `lxml` | |
| 图片处理、配图裁剪 | `Pillow` | |
| 拆词检查（可选） | `jieba` | `lint.py` 用；不装时窄栏换行只给提示 |
| 渲染核对（把 pptx 出图看效果） | `PyMuPDF` + **LibreOffice** | `render.py` 走 soffice |

```bash
uv venv ~/.venvs/ppt                                             # 或 python -m venv
uv pip install --python ~/.venvs/ppt/bin/python \
    python-pptx python-docx lxml Pillow jieba PyMuPDF
```

LibreOffice（只为把成品渲染成图做核对，不装也能出 PPT）：

```bash
brew install --cask libreoffice        # macOS
```

> macOS 上 `brew` 若在 keg relocation 阶段报 EPERM（非标准前缀环境），改用官方 dmg 安装，
> 装完确认 `/Applications/LibreOffice.app/Contents/MacOS/soffice --version` 有输出即可。

中文字体缺失会导致渲染出方框：把微软雅黑（或思源黑体）装进系统，
`assets/fonts/fonts.conf` 是给 Linux 首装的 fontconfig 兜底。

**脚本一律用这个 venv 的 python 跑**（系统 python 没有上面这些库）：

```bash
~/.venvs/ppt/bin/python scripts/dense_deck.py --content content.json --output out/deck.pptx
```

## 3. 怎么用

直接说：「用 research-ppt-skill 做一份科研汇报 PPT，主题是 X」，并附上材料（论文、原图、数据、署名）。
技能会先问清主题、索要缺的材料，出逐页大纲给你确认，再生成原生可编辑的 PPTX、讲稿与预览。
细节见 `SKILL.md`；视觉规范与版式库在 `references/info-style.md`，完整示例是 `assets/info-demo.json`。

## 4. 已知的坑（照做能省几小时）

1. **别用 WPS / PowerPoint 反复打开成品做「验收」**：它们打开后可能自动重存，改掉交付物。
   留一份构建产物原件并比对 md5，验收另存副本。
2. **配图不要调本机生图模型**（内存不够时会被系统杀进程并卡死电脑）；
   配图走网页端 ChatGPT 生成，取回后逐张看图确认真表现了该页要讲的机制，再写进「图 → 页」映射。
3. **改 `content.json` 要按结构改，不要做字符串替换**：缩进差两格就 0 命中，而脚本照样 exit 0、稿子没变。
   用 `json.load` → 按标题定位 → 改 dict → `json.dumps(ensure_ascii=False, indent=1)`。
4. **内嵌视频只能拿 PowerPoint / WPS 验收**：LibreOffice 只把视频当静帧渲染，判定不了是否嵌好。
   媒体部件名用 ASCII，一个 `p:cNvPr` 里只能有一个 `hlinkClick`，整份压缩到 25 MB 内并另交独立 mp4 后备。
5. **生成后一定逐页看渲染图**：脚本保证配色字号，但分区空了一块、图在格子里太小这类问题只有看图才发现。
6. `scripts/render.py` 依赖 LibreOffice；`scripts/notes_docx.py` 依赖 python-docx；
   `scripts/embed_video.py`、`scripts/prep_figs.py` 是较新加入的脚本，跑之前先看它们顶部的命令行用法。

---

版本：2026-10-16（蓝白信息图式 v6）。
