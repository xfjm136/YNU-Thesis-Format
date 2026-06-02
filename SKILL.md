---
name: ynu-thesis-format
description: Strictly audit and normalize Microsoft Word DOCX files against Yunnan University undergraduate year-paper and graduation-thesis formatting requirements, especially the official 2024 云大本科〔2024〕54号《云南大学本科毕业论文（设计）写作规范》. Use when Codex is asked to check, nitpick, back up, or format a 云南大学本科生学年论文/毕业论文 without changing thesis content, including margins, fonts, heading systems, page breaks, page numbers, headers, automatic table of contents, figures, tables, formulas, references, citations, abstract, acknowledgements, and appendices.
---

# YNU Thesis Format

## Scope

Use this skill for Yunnan University undergraduate 学年论文 or 毕业论文 DOCX formatting tasks. For undergraduate graduation theses/design papers, default to the 2024 official standard in `references/ynu-2024-undergraduate-thesis-writing-spec.md`; use `references/ynu-format-requirements.md` only as a condensed fallback or for generic year-paper work when no newer official file is supplied.

The non-negotiable boundary is: preserve content. Do not rewrite, delete, add, paraphrase, translate, renumber, or reorganize thesis text unless the user separately asks for content editing. Formatting-only changes may include styles, fonts, margins, line spacing, indentation, page/section breaks, captions' visual style, table layout, headers, footers, page numbers, field/TOC refresh, and non-destructive Word field repair.

Non-destructive Word field repair is allowed when it restores existing automatic structure without authoring new content. Examples: remove invalid formatting switches such as `\* MERGEFORMAT` from a broken `REF` field when the referenced bookmark already exists; update TOC, page-number, `STYLEREF`, `REF`, `PAGEREF`, or `SEQ` fields; repair generated field results that currently render as Word errors. This may change visible generated field results from an error or stale value to the current automatic result, but it must not invent a target, manually type citation numbers, or change surrounding prose.

If a requested action would change visible words, heading text, manually authored citation numbers, reference entries, formula semantics, or figure/table content, report it as a content issue instead of fixing it automatically.

## Required Companion Skills

For DOCX work, also use `minimax-docx` when available. For academic-thesis auditing, also use `docx-thesis-format` when available. When the governing standard is a PDF supplied by the user, also use the `pdf` skill to extract or inspect the rule text before applying it.

Do not treat `Invoke-YnuThesisOneShot.ps1` as the whole skill. It is only a helper. The skill-level job is to orchestrate: load the official rule reference, audit the original, produce a Markdown report, use DOCX editing tools to create a format-only copy, optionally use page images for visual inspection, then run and summarize the final audit.

Default report policy: every audit should produce Markdown first, then a PDF copy of that Markdown report. JSON remains a machine-readable companion artifact. If PDF export fails because Word/PDF tooling is unavailable, keep the Markdown report and state the PDF failure clearly.

## Main Workflow

1. Identify the target DOCX and the governing standard.
   - If the user supplies an official template, PDF, ZIP, or department-specific rule file, read it first and treat it as higher priority.
   - If no rule file is supplied and the task is a 云南大学本科毕业论文（设计）, use `references/ynu-2024-undergraduate-thesis-writing-spec.md`.
   - For strict/full audits, also read `references/ynu-2024-pdf-article-map.md` to ensure every PDF article is covered.
   - If no rule file is supplied and the task is a generic 学年论文, use `references/ynu-format-requirements.md` but keep the 2024 heading/page/TOC rules in mind.
2. Strictly audit the original before editing.
   - Produce `*_ynu_format_audit_before.md` and `*_ynu_format_audit_before.json`.
   - Also produce `*_ynu_format_audit_before.pdf` by default as a PDF rendering of the Markdown report.
   - The Markdown report must be useful by itself: include PASS/WARN/FAIL, exact category, observed value, expected rule, affected paragraph/page/section when available, and whether the item is automatically fixable.
   - Include report-only content/compliance issues separately from format-fixable issues.
   - Audit categories must cover all PDF articles 1-13, not just common Word layout checks.
3. Use visual inspection when layout risk is high.
   - For header/footer, page-number, TOC, cover, table/figure split, or dense table issues, export the DOCX to PDF/images and inspect representative pages.
   - Use page images as evidence for visual layout problems; do not OCR or rewrite content from images.
4. Back up before formatting.
   - Always create a timestamped backup copy before writing an adjusted DOCX.
   - Never overwrite the user's original DOCX unless the user explicitly requests it.
5. Create a formatted copy.
   - Use a filename suffix such as `_ynu_format_adjusted.docx`.
   - Apply only formatting changes. Keep content text identical outside generated TOC/page-number field display updates.
   - Replace static/manual TOC text with a real Word automatic TOC when safe. Long or punctuated headings in the TOC are valid; do not flag them merely because they contain punctuation or exceed a short length.
   - Detect heading systems using top-level anchors first. In the official 2024 PDF, the second system is `第一章` / `第一节` / `一、` / `（一）`; this is not mixed numbering.
   - Use `minimax-docx`/OpenXML/Word COM for precise DOCX edits when the helper script is insufficient. Prefer semantic styles and section properties over direct formatting when possible, but direct formatting is allowed when the existing document is inconsistent and the task is formatting-only.
6. Verify and report the adjusted copy.
   - Reopen the adjusted DOCX in Word.
   - Confirm no Word field errors such as `错误!未找到引用源`, `错误!不能识别的开关参数`, or `Error! Reference source not found`.
   - Confirm thesis text is unchanged, excluding automatically generated TOC/page-number field results.
   - Produce final `*_ynu_format_audit.md` and `*_ynu_format_audit.json`.
   - Also produce final `*_ynu_format_audit.pdf` by default.
   - Report remaining non-format issues as warnings rather than silently changing them.

## Script Quick Start

Run the one-shot path from PowerShell. Prefer PowerShell 7 (`pwsh`) when available; the scripts are UTF-8 with BOM and also work in Windows PowerShell on normal Office installations.

```powershell
$script = Join-Path $env:USERPROFILE ".codex\skills\ynu-thesis-format\scripts\Invoke-YnuThesisOneShot.ps1"
pwsh -ExecutionPolicy Bypass -File $script `
  -InputDocx "D:\path\paper.docx" `
  -Mode AuditAndFix
```

Useful options:

```powershell
# Report only, no output copy
-Mode Audit

# Create output at a chosen path
-OutputDocx "D:\path\paper_ynu_format_adjusted.docx"

# Choose title numbering system for audits: Auto, First, Second, Third
-HeadingSystem Auto

# Disable PDF report export only when PDF tooling is unavailable or undesired
-NoReportPdf
```

The helper one-shot script emits:

- `*_ynu_format_backup_YYYYMMDD_HHMMSS.docx`
- `*_ynu_format_adjusted.docx`
- `*_ynu_format_audit_before.json` / `*_ynu_format_audit_before.md` for the original
- `*_ynu_format_audit_before.pdf` for the original Markdown audit report
- `*.intermediate.json` / `*.intermediate.md` from the first formatting pass, useful for debugging
- `*_ynu_format_audit.json`
- `*_ynu_format_audit.md`
- `*_ynu_format_audit.pdf`
- optional page previews when `-ExportPreviewImages` is used

## Skill-Level Acceptance Gate

Do not present the skill run as fully complete until these checks are true or clearly reported as blocked:

- A Markdown audit report exists for the original DOCX before modification.
- A PDF rendering of the original Markdown audit report exists, unless PDF export was explicitly disabled or clearly failed.
- A separate adjusted DOCX exists, and the original DOCX was not overwritten.
- A Markdown audit report exists for the adjusted DOCX.
- A PDF rendering of the adjusted Markdown audit report exists, unless PDF export was explicitly disabled or clearly failed.
- The adjusted DOCX opens in Word without repair prompts.
- Word reports exactly one automatic TOC when a TOC title exists.
- No field result in the adjusted DOCX contains `错误`, `不能识别`, `引用源`, `Error`, or `Reference source`. Repair existing broken REF/PAGEREF/SEQ fields only by fixing the Word field structure and updating fields when the target already exists; otherwise report the broken field instead of inventing visible text.
- Thesis prose, headings, captions, formulas, references, and manually authored citation numbers are unchanged unless the user explicitly requested content edits. Automatically generated Word field results may update when fields are repaired or refreshed.
- The report distinguishes format fixes from content issues. Missing sections, unnumbered reference entries, citation-order problems, and keyword term text are reported, not silently authored.
- Any visual review performed is summarized in the Markdown report or final response.
- The final response and report include the disclaimer that this skill is an auxiliary formatter/auditor, does not guarantee correctness, and requires human review.

## Strict Audit Categories

Check at least:

- Page setup: A4, margins top 2.5 cm, bottom 2.0 cm, left 3.0 cm, right 2.0 cm.
- Structure: Chinese abstract, Abstract, TOC, body, conclusion, references, acknowledgement, appendix when present. Cover, evaluation forms, originality statement, and authorization pages are template-specific; check them only when the user supplies a template or explicitly asks.
- Heading system: one consistent system only. Do not automatically rewrite heading text.
  - Official 2024 second system: `第一章` / `第一节` / `一、` / `（一）`.
  - Do not treat `第一章` plus `一、` as mixed unless another top-level system such as `1 标题` or top-level `一、` is also present.
- Heading format: first level 三号黑体 centered; second level 四号黑体; third and lower 小四黑体.
- Body: 小四宋体; English/numbers Times New Roman; fixed 22 pt line spacing; justified; first-line indent 2 Chinese characters.
- TOC: automatic, includes abstract, Abstract, headings through level 3, references, acknowledgement, appendix; dot leaders and right-aligned page numbers; no body paragraphs mixed into TOC.
- Abstract/Abstract: each starts on a new page; title format correct; keywords label bold and delimiter format correct.
- Page numbers: if cover/declaration/evaluation pages already exist, audit them only against the supplied template and do not require them in generic work; preliminary pages use uppercase Roman numerals; body/references/acknowledgement/appendix use `第×页（共×页）`.
- Header: starts from body page 1, centered 五号宋体, current first-level heading.
- Figures/tables: chapter-based numbering, correct caption position, caption font, and not split from object.
- Formulas: Word equation objects, own centered line, chapter-based numbering at right, no dotted leaders, no slash fractions where stacked structure is expected.
- Citations/references: sequential numeric citations, no citations in headings, no broken cross-reference fields, references follow GB/T 7714-2015 formatting.
- Content-preservation guard: adjusted copy must not change thesis prose, reference entries, headings, formula semantics, or data.

## When Fixing Is Unsafe

Do not automatically fix these; report them clearly:

- Mixed or wrong heading numbering text.
- Missing required sections whose text would need to be authored.
- Wrong or missing citation/reference content.
- Broken cross-reference fields whose target bookmark/caption/sequence is missing or ambiguous.
- Manual citation-number edits, manual reference renumbering, or typed replacement of a broken field result.
- Keyword term changes, keyword translation, punctuation corrections that change visible text, title shortening, abstract rewriting, acknowledgement writing, or word-count padding.
- Formula semantic errors.
- Figure/table numbering that requires changing visible caption text.
- Department-specific requirements that conflict with the general YNU standard.
- Missing cover, evaluation form, originality statement, authorization page, or other template-specific front matter when no official template or explicit user request requires it.

## Disclaimer

Always communicate this limitation in reports or final responses when relevant: this skill is limited to auxiliary format auditing and format adjustment. It does not modify thesis substance, does not perform academic/content correctness review, does not guarantee that every change is correct or accepted by a department, and cannot replace manual review against the latest official requirements, advisor requirements, and submitted Word rendering.

## Visual And DOCX Editing Guidance

- Use Word COM for opening/saving, field updates, page layout, and page-rendered checks when Microsoft Word is available.
- Use `minimax-docx` / OpenXML for structural edits that are easier or safer in XML: style definitions, section properties, table formatting, header/footer XML, and content-preservation diffs.
- Use exported PDF/page images to inspect things automation often misses: visible page number style, header start page, TOC blank line, table/figure split across pages, caption placement, cover/front-matter page numbering, and visually wrong font/spacing.
- For page images, inspect only the layout and formatting. Do not transcribe or alter thesis content based on visual inspection.
- If tools disagree, trust the official PDF rule first, then Word-rendered visual evidence, then XML/COM measurements. Explain unresolved conflicts in the Markdown report.

## Reference

Use `references/ynu-2024-undergraduate-thesis-writing-spec.md` for the official 2024 undergraduate thesis/design checklist. Use `references/ynu-2024-pdf-article-map.md` for strict article-by-article coverage. Use `references/ynu-format-requirements.md` only as the short fallback checklist.
