# Thesis to PPT Skill

根据毕业论文生成毕业答辩 PPT。

## 使用方式

### 方式一：网页匹配工具（推荐）

适合需要人工审核图文匹配的场景。

```bash
# 1. 启动网页工具
python scripts/thesis2ppt_web.py

# 2. 浏览器打开 http://127.0.0.1:5000

# 3. 输入论文路径 → 加载 → 勾选匹配章节和图片 → 导出PPTX
```

网页工具提供：
- 左侧章节列表（点击选中）
- 右侧图片网格（点击匹配到当前选中章节）
- 实时匹配状态显示
- 一键导出 PPTX

### 方式二：命令行（含图文审核）

适合 Claude 辅助审核图片匹配的场景。

```
/thesis2ppt <thesis_file>
```

工作流：
1. 解析论文 → 提取章节结构和图片引用
2. Claude 审核每张图片的 caption 和上下文，逐张判断归属章节
3. 生成 image_mapping.json（审核后的匹配结果）
4. 生成 PPTX（带 --image-mapping 参数）

### 方式三：纯命令行

```bash
python scripts/thesis2ppt.py <thesis_file> \
  --output defense.pptx \
  --title "论文标题" \
  --author "作者" \
  --advisor "导师" \
  --university "学校" \
  --date "日期" \
  --image-mapping image_mapping.json  # 可选
```

## PPT 结构

- 封面：标题、作者、导师、学校、日期
- 目录：一级章节标题
- 内容页：二级标题 + 左图右文环绕布局
- 致谢 + Q&A 结束页

## 依赖

```bash
pip install -r requirements.txt
pip install flask  # 网页工具需要
```
