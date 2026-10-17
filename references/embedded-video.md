# 把 mp4 真内嵌进 pptx（录屏/演示页）

录屏页的成稿标准是**内嵌视频**：点击封面就在放映中播放，不是截图、不是超链接。
做法与 PowerPoint 自己写出的结构一致（脚本：`scripts/embed_video.py`）。

## 步骤

1. 先正常把**封面图当普通 `figure` 填进图位**（`add_picture`）：它就是静帧，也是媒体对象的载体；
   封面自己叠播放三角 + 深蓝条（写案例名与时长）。
2. 视频作为**媒体部件**放进 `ppt/media/`，同一点上挂两条关系：
   `relationships/video` 与 `relationships/media`（`RT.VIDEO` / `RT.MEDIA`）。
3. 在该 `p:pic` 里注入：`p:nvPr` 最前放 `<a:videoFile r:link="rIdVideo"/>`，
   最后放 `<p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-04AC343B8D0A}"><p14:media r:embed="rIdMedia"/></p:ext></p:extLst>`；
   `p:cNvPr` 里加 `<a:hlinkClick r:id="" action="ppaction://media"/>`（点击播放）。
4. **一个 `p:cNvPr` 里只能有一个 `hlinkClick`**：给图挂过文件超链接时必须先删再插，
   否则 PowerPoint 报“需要修复”。只给宽度 > 3 in 的大图挂——页眉校徽也是图片形状。

## 保存后自检

- `[Content_Types].xml` 里有 `Extension="mp4"`；
- 包内 `ppt/media/*.mp4` 存在；
- 对应 slide 的 rels 含 `relationships/video`；
- slide XML 里 `videoFile` 与 `p14:media` 都在。

## 验收

用 WPS / PowerPoint 打开看**是否报修复**；没报修复、且选中封面时出现“视频工具/播放”控件
（WPS 会显示进度条）就是成了。LibreOffice 只当静帧渲染，**不能拿它判定视频是否嵌好**。

## 体积与兼容

内嵌会让 pptx 变大（实测 17.8 MB → 24.8 MB，两条片子共 6.8 MB）。
压缩命令、25 MB 上限、备用独立 mp4、以及「验证用的打开方式会改文件」这条坑，
见 SKILL.md 的「跨机器要稳：内嵌视频的两条规定」——那份是常驻规则，这里是配方。
