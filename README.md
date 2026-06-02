# ynu-thesis-format

云南大学本科生学年论文/毕业论文（设计）DOCX 格式审查与格式辅助调整 skill。该 skill 面向 Codex 使用，内置云南大学 2024 年本科毕业论文（设计）写作规范相关检查流程，并提供 PowerShell/Word COM 辅助脚本。

## 功能概览

- 严格审查 DOCX 格式，并输出 Markdown、PDF、JSON 报告。
- 在不覆盖原始文件的前提下生成格式调整版 DOCX。
- 检查 A4 页面、页边距、标题层级、标题格式、正文格式、目录、页眉页码、章节分页、摘要/关键词、参考文献、引文标注、图题表题、公式、附录等项目。
- 支持生成 before/after PDF 与 PNG 页面预览，用于人工复核目录、页眉页码、分页和字段显示。
- 支持修复不改变正文语义的 Word 自动域问题，例如目标书签存在但因无效字段开关导致的 REF 交叉引用错误。

## 重要免责声明

本 skill 仅限于论文格式审查和格式辅助调整，不修改论文的具体学术内容、观点、论证、参考文献事实、数据、公式语义或作者表达。

本 skill 不保证审查结果或格式调整结果一定正确，也不保证一定符合学院、导师、教务部门或提交系统的最终要求。不同学院、课程、教师或模板可能存在额外要求或局部差异。

本 skill 只是辅助功能，不能替代人工复核。使用者必须在提交前自行打开原始 DOCX、修正版 DOCX、Markdown/PDF 检查报告和页面预览，逐项核对官方最新版要求、学院通知、导师要求和 Word 实际渲染效果。

使用本 skill 前请保留原始文件备份。脚本默认不覆盖原始 DOCX，但仍建议自行备份重要论文文件。

## 适用范围

适合用于：

- 云南大学本科生学年论文格式审查。
- 云南大学本科毕业论文（设计）格式审查。
- 基于云南大学通用写作规范的 DOCX 格式辅助调整。
- 需要生成 Markdown/PDF 格式检查报告和格式修正版 DOCX 的场景。

不适合用于：

- 论文降重、润色、改写、扩写或学术内容编辑。
- 参考文献事实真伪校验。
- 替代导师、学院或教务部门的最终审核。
- 在没有人工复核的情况下直接提交最终论文。

## 目录结构

```text
ynu-thesis-format/
├── SKILL.md
├── README.md
├── .gitignore
├── agents/
│   └── openai.yaml
├── references/
│   ├── ynu-2024-pdf-article-map.md
│   ├── ynu-2024-undergraduate-thesis-writing-spec.md
│   └── ynu-format-requirements.md
└── scripts/
    ├── Invoke-YnuThesisFormat.ps1
    └── Invoke-YnuThesisOneShot.ps1
```

## 安装

将整个 `ynu-thesis-format` 文件夹复制到 Codex skills 目录：

```powershell
Copy-Item -Recurse -Force .\ynu-thesis-format "$env:USERPROFILE\.codex\skills\ynu-thesis-format"
```

重新打开 Codex 会话后，使用 `$ynu-thesis-format` 或在任务中明确要求审查/修正云南大学论文格式即可触发。

## 使用方式

推荐使用一把梭脚本：

```powershell
$script = "$env:USERPROFILE\.codex\skills\ynu-thesis-format\scripts\Invoke-YnuThesisOneShot.ps1"
pwsh -NoProfile -ExecutionPolicy Bypass -File $script `
  -InputDocx "D:\path\paper.docx" `
  -Mode AuditAndFix `
  -ExportPreviewImages
```

只审查、不生成修正版：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File $script `
  -InputDocx "D:\path\paper.docx" `
  -Mode Audit
```

指定输出文件：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File $script `
  -InputDocx "D:\path\paper.docx" `
  -Mode AuditAndFix `
  -OutputDocx "D:\path\paper_ynu_format_adjusted.docx" `
  -ReportMd "D:\path\paper_ynu_format_audit.md" `
  -ReportPdf "D:\path\paper_ynu_format_audit.pdf"
```

默认会导出 PDF 报告。如 PDF 工具不可用或不需要 PDF，可关闭：

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File $script `
  -InputDocx "D:\path\paper.docx" `
  -Mode AuditAndFix `
  -NoReportPdf
```

## 输出文件

典型输出包括：

- `*_ynu_format_backup_YYYYMMDD_HHMMSS.docx`：原始文件备份。
- `*_ynu_format_adjusted.docx`：格式调整版 DOCX。
- `*_ynu_format_audit_before.md`：原始 DOCX Markdown 检查报告。
- `*_ynu_format_audit_before.pdf`：原始 DOCX PDF 检查报告。
- `*_ynu_format_audit_before.json`：原始 DOCX 机器可读检查报告。
- `*_ynu_format_audit.md`：修正版 Markdown 检查报告。
- `*_ynu_format_audit.pdf`：修正版 PDF 检查报告。
- `*_ynu_format_audit.json`：修正版机器可读检查报告。
- `*.intermediate.*`：第一轮格式化后的中间报告，用于调试。
- `*_ynu_format_previews/`：可选的 before/after PDF 与 PNG 页面预览。

## 自动修复边界

允许自动处理：

- 样式、字体、字号、行距、缩进、对齐、页边距、A4 页面设置。
- 页眉、页脚、页码分节。
- 自动目录创建、更新和格式化。
- 章节分页和标题样式应用。
- 不改变正文语义的 Word 自动域刷新。
- 在目标书签存在的前提下，修复由无效开关导致的 REF 交叉引用错误。

不自动处理：

- 改写论文正文、标题文字、摘要、关键词、致谢或参考文献内容。
- 手工改引文编号、手工改参考文献编号、编造缺失引用目标。
- 修改公式语义、图表数据、研究结论或学术观点。
- 补写缺失章节、补充致谢、扩充关键词或替换参考文献。

## 环境要求

- Windows。
- Microsoft Word，可通过 COM 自动化打开和导出 DOCX/PDF。
- PowerShell 5.1 或 PowerShell 7，推荐 `pwsh`。
- 可选：`pdftoppm`，用于将 PDF 页面导出为 PNG 预览。

## 已知限制

- Word 字体、模板、系统语言和插件差异可能导致渲染差异。
- 自动审查无法完全判断 GB/T 7714-2015 著录项事实准确性。
- 自动审查无法保证所有学院/教师/课程的个性化格式要求。
- 自动修复后的文件必须人工打开复核，尤其是目录页码、页眉页脚、章节分页、交叉引用和参考文献。

## 维护说明

当云南大学发布新的写作规范、学院模板或格式通知时，应先更新 `references/`，再调整 `SKILL.md` 和脚本检查逻辑。更新后至少用一个真实 DOCX 做完整 `AuditAndFix` 验证，并确认 Markdown/PDF 报告、修正版 DOCX 和页面预览均能正常生成。
