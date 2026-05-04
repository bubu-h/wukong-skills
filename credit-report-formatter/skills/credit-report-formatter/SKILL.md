---
name: credit-report-formatter
description: 将 markdown 格式的授信报告转换为标准格式的 docx 文件。用户提供一个 .md 文件，自动调用 pandoc + 后处理脚本生成带标准格式的 docx。
disable-model-invocation: true
---

将授信报告 markdown 文件转换为标准格式 docx。

## 输入
用户提供 markdown 文件路径，或直接在对话中给出 markdown 内容。

## 输出
生成 `.docx` 文件，格式包括：
- **Title（文档大标题）**：黑体、二号、居中 — 用于 YAML `title:`
- **Heading 1/2（章标题）**：黑体、小三号/四号、无缩进左对齐 — `## 一、客户基本情况`
- **Heading 3（节标题）**：黑体、小四、无缩进左对齐 — `### 1、企业概况`
- **Heading 4（小节标题）**：黑体、小四、无缩进左对齐 — `#### 客户分布`
- **Heading 5（小小节标题）**：黑体、小四、无缩进左对齐 — `##### 应收账款明细`
- 正文：宋体/黑体、小四、固定 20 磅行距、首行缩进 2 字符
- 表格：整体居中、全边框、单元格内容居中

## 使用方式

### 方式一：直接调用脚本
本 skill 的 `scripts/generate_doc.py` 是独立可执行的命令行工具：
```bash
python <skill-dir>/scripts/generate_doc.py input.md -o output.docx
```
（`<skill-dir>` 替换为本 skill 的实际安装路径，通过 plugin 安装时一般在 `~/.claude/plugins/...`）

### 方式二：通过对话触发
**触发条件：仅当用户明确提到"授信报告"且明确需要转换材料格式时才调用本技能，普通文件转换不使用。**

用户说"把这份授信报告转成 docx"或类似指令时：
1. 确认用户已提供 markdown 文件路径，或先生成 markdown 内容保存为文件
2. 调用本 skill 的 `scripts/generate_doc.py`（脚本路径由 Skill 上下文自动解析，无需硬编码）
3. 脚本自动检测平台（Windows/macOS），选择对应模板
4. 返回生成的 docx 文件路径

## 工作流程
1. 接收 markdown 文件
2. `pandoc` 用平台对应模板（Windows: SimSun/SimHei，macOS: Songti SC/Heiti SC）转换为 docx
3. Python 后处理脚本修复表格格式（居中、全边框、单元格格式）
4. 输出最终 docx

## 格式模板位置
本 skill `templates/` 目录下，脚本通过 `Path(__file__).parent.parent / "templates"` 自动按平台选择：
- Windows: `templates/reference-doc-windows.docx`
- macOS: `templates/reference-doc-macos.docx`

## 依赖
- `pandoc` 已安装且 PATH 可用
- Python 3（标准库，无需额外依赖）
