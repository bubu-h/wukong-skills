# Changelog

## 2026-05-06

### Removed
- 删除 `credit-report-formatter` skill

---

## [credit-report-formatter] 0.1.3 - 2026-05-04

### Changed
- Title（文档大标题）由左对齐改回居中（恢复 v0.1.1 之前的样式，Windows / macOS 双模板均已修正）
- 同步更新根目录 README 中的版本号表格

## [credit-report-formatter] 0.1.2 - 2026-05-04

### Changed
- Title（文档大标题）由居中对齐改为左对齐（Windows / macOS 双模板均已修正）

## [credit-report-formatter] 0.1.1 - 2026-05-04

### Fixed
- 修复 Title（文档大标题）继承 Normal 样式的固定 20 磅行距问题，改为单倍行距（Windows / macOS 双模板均已修正）

## [credit-report-formatter] 0.1.0 - 2026-05-04

### Added
- 初始发布
- 支持 Windows / macOS 双平台模板（自动检测平台）
- 支持 Title / Heading 1-5 / 正文 / 表格 的标准格式输出
- pandoc + Python 后处理两阶段流水线
