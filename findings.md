# Findings: thesis2ppt Skill

## 论文信息
- 标题：后向散射标签节点调制技术研究与设计
- 作者：成俊良 | 导师：张俊 | 学院：信息工程学院
- 格式：DOCX，31 张图片（18 EMF + 11 PNG + 2 JPEG）
- 结构：6 章 + 参考文献 + 致谢 + 3 附录

## 关键技术发现

### 1. python-pptx 不支持 EMF
- EMF 是 Windows 矢量图格式，python-pptx 无法嵌入
- 解决：PowerShell 调用 `System.Drawing.Imaging.Metafile` → PNG
- 转换时机：网页加载时 + 导出前，确保浏览器和 PPTX 都能用

### 2. 图片自动匹配不可靠
- 关键词重叠算法有大量误匹配
- **正确方案**：网页人工勾选 — 用户可视化匹配，所见即所得
- 支持每节多图：点击添加/取消，首图嵌入内容页，其余自动附图

### 3. python-pptx 单位陷阱
- `self.w`/`self.h` 存储 cm 值（25.4 / 19.05），`Inches()` 期望英寸
- 修复：`_setup_theme` 将 cm 转为 inches，slide_width 保持 cm

### 4. DOCX 章节结构特点
- 部分 level-2 节无直接内容（如 2.1），内容在 level-3 子节中
- 修复：`map_sections_to_slides` 检测空内容节，自动合并子节前几句
- 三级标题（如 2.1.1）有独立内容，纳入章节列表

### 5. 图注提取策略
- 论文图片的图注（"图X.X ..."）通常在图片的**下一个段落**
- 修复：`find_image_references` 优先匹配 `next_text` 中"图+数字"模式
- 查 3 个位置：next_text → prev_text → para_text，优先"图X.X"

### 6. Flask dev server 行为
- 单线程，大文件上传可用（FormData multipart）
- 需注意端口占用和旧进程清理
- 中文文件路径在 curl 中可能乱码，拖拽上传避开了此问题
