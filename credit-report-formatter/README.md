# credit-report-formatter

将 markdown 格式的授信报告转换为符合标准格式的 docx 文件。

## 功能

接收一个授信报告 markdown 文件，输出格式标准的 docx：

- **Title（文档大标题）**：黑体、二号、居中
- **Heading 1/2（章标题）**：黑体、小三号/四号、无缩进左对齐
- **Heading 3（节标题）**：黑体、小四、无缩进左对齐
- **Heading 4/5（小节/小小节）**：黑体、小四、无缩进左对齐
- **正文**：宋体/黑体、小四、固定 20 磅行距、首行缩进 2 字符
- **表格**：整体居中、全边框、单元格内容居中

## 安装

```
/plugin marketplace add bubu-h/wukong-skills
/plugin install credit-report-formatter@wukong-skills
```

## 依赖

- **pandoc**：安装后需在 PATH 中可用
  - Windows: 从 [pandoc.org](https://pandoc.org/installing.html) 下载安装包
  - macOS: `brew install pandoc`
- **Python 3**（仅标准库，无需 `pip install`）

## 使用方式

### 方式一：通过对话触发（推荐）

在 Claude Code 中说：

> 把这份授信报告转成 docx，markdown 文件在 xxx.md

Claude 会自动调用本 skill 处理。

### 方式二：直接调用脚本

```bash
python <plugin-skill-dir>/scripts/generate_doc.py input.md -o output.docx
```

`<plugin-skill-dir>` 是 plugin 安装后 skill 的实际路径，一般在 `~/.claude/plugins/<...>/credit-report-formatter/skills/credit-report-formatter/` 下。

## 工作流程

1. 接收 markdown 文件
2. `pandoc` 用平台对应模板（Windows: SimSun/SimHei，macOS: Songti SC/Heiti SC）转换为 docx
3. Python 后处理脚本修复表格格式（居中、全边框、单元格格式）
4. 输出最终 docx

## 平台支持

| 平台 | 字体 | 模板文件 |
|---|---|---|
| Windows | SimSun（宋体）/ SimHei（黑体） | `templates/reference-doc-windows.docx` |
| macOS | Songti SC / Heiti SC | `templates/reference-doc-macos.docx` |

脚本通过 `sys.platform` 自动检测当前系统并选择对应模板，无需用户配置。

Linux 暂未提供模板，欢迎贡献。

## 文件结构

```
credit-report-formatter/
├── .claude-plugin/
│   └── plugin.json
├── skills/
│   └── credit-report-formatter/
│       ├── SKILL.md            # Skill 定义
│       ├── scripts/
│       │   ├── generate_doc.py # 主入口脚本
│       │   └── postprocess.py  # 后处理工具(独立可调用)
│       └── templates/
│           ├── reference-doc-windows.docx
│           └── reference-doc-macos.docx
└── README.md
```
