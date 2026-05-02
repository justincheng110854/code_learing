# Task Plan: thesis2ppt Skill 完善

## Goal
从毕业论文生成毕业答辩 PPT。已迭代为**网页拖拽匹配工具**，用户自主勾选图文配对。

## Phases

### Phase 1: 修复目录 — 一级标题 [complete]
- TOC 改用 level-1 章节标题（如"第一章 绪论"）

### Phase 2: CLI 图片审核机制 [complete]
- 添加 --image-mapping 参数，支持 Claude 审核后的 JSON 映射

### Phase 3: 网页拖拽匹配工具 [complete]
- Flask 单页应用 (scripts/thesis2ppt_web.py)
- 拖拽/浏览上传论文 → 自动解析
- 左侧章节列表 + 右侧图片网格 → 点击勾选匹配，支持每节多图
- EMF 自动转 PNG，浏览器直接显示
- 一键导出 PPTX（首图嵌入内容页，多余图自动生成附图页）

### Phase 4: 章节结构增强 [complete]
- 修复 2.1 节丢失：空内容节自动合并子节（2.1.1~2.1.4）内容
- 加入三级标题识别：2.1.1、2.1.2、3.2.1、3.2.2 等子节纳入章节列表
- 章节数从 15 → 22

### Phase 5: 图注提取改进 [complete]
- `find_image_references` 优先匹配"图X.X"格式的下段文字作为图注
- 图片卡片下方显示正确的图注而非混乱正文

### Phase 6: PPT 内容质量审阅 [pending]
- 审核 bullet points 准确性
- 可考虑网页中加入 bullet 编辑功能

## Architecture
```
thesis2ppt.py        — 核心引擎（解析 + 生成）
thesis2ppt_web.py    — Flask 网页（拖拽上传 + 人工匹配 + 导出）
thesis2ppt.md        — Skill 定义
```

## Decisions
- 图片匹配：由自动匹配 → 网页人工勾选（用户最了解自己的论文）
- 章节结构：二级 + 三级标题，空内容节合并子节内容
- 多图匹配：每节可配多图，导出时首图嵌入 + 其余附图
- EMF→PNG：PowerShell GDI 即时转换
- 图注提取：优先匹配 next_text 中"图+数字"模式

## Errors Resolved
| Error | Resolution |
|-------|------------|
| EMF 无法嵌入/显示 | 加载时 + 导出前自动转 PNG |
| Inches/Cm 单位混淆 | self.w/h 改为 inches |
| 中文路径编码 | tempfile 目录 |
| 旧服务器占用端口 | taskkill 清理后重启 |
| 浏览器缓存旧页面 | 重启服务 + 强制刷新 |
| 2.1 节丢失 | 空内容节合并子节内容 |
| 图注为正文非标题 | 优先匹配"图X.X"模式 |
