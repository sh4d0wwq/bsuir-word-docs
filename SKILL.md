---
name: bsuir-word-docs
description: Creates Word (.docx) documents formatted per BSUIR standard STP 01-2017 (ГОСТ 2.105) - отчеты по лабораторным работам, пояснительные записки к курсовым и дипломным проектам, рефераты. Uses the word-document-server MCP plus a builder script. Use when the user asks to write/format a lab report, пояснительная записка, курсовой/дипломный проект, отчет, or a .docx "по стандарту/по СТП/по ГОСТ".
---

# BSUIR Word documents (СТП 01–2017)

Script `S=$HOME/.cursor/skills/bsuir-word-docs/scripts/bsuir_docx.py` (python-docx; Word needed for `build`/`fix`/`preview`, PyMuPDF for `preview`).
Always use absolute .docx paths. The file must be closed in Word while editing.

Scope: this skill only builds the Word document from material the user provides (text, code files, images). Producing that material (running programs, taking screenshots) is outside the skill.

## Rules
- Unknown data (ФИО, группа, преподаватель, дисциплина, кафедра, тема) → do NOT search the web or other documents. Put a placeholder `[[ФИО студента]]`, `[[Преподаватель]]`, `[[Дисциплина]]`… `fix` lists all `[[…]]`; name them in the final answer.
- Code listings from files: `@code path` in the markup instead of pasting code into markdown by hand.
- Write chunk/config files with the Write tool (not via shell strings), keep them in a temp folder outside the result folder, delete at the end.
- If the target .docx is open (script prints `is locked`), build in the temp folder and copy over it; if the copy fails, save as `<name>_v2.docx` and tell the user.

## Workflow
1. Write `cfg.json` (below) and chunks `ch1.md`, `ch2.md`… in markup (below), one section per chunk.
2. `python $S build cfg.json OUT.docx ch1.md ch2.md …` — builds from scratch: styles, margins 30/15/20/27 mm, TNR 14, title page, contents (course/diploma), page numbers, all chunks, then `fix`. Output: page count, lint (missing references to figures/tables/appendices/sources, citation order, trailing dots, `[[placeholders]]`) and `layout:` lines (near-empty pages, figure split from caption, big gap before a moved figure; oversized images are scaled automatically).
3. Resolve every lint/layout line by editing chunks with StrReplace, then rerun `build`. Stop when only the placeholder line (if any) remains.
4. Optional, once at the end: `python $S preview OUT.docx DIR` → one `sheet_NN.png` per 12 pages (low-res overview). Only if something looks wrong: `preview OUT.docx DIR 5,7` for those pages at full size.
5. Delete temp files.

Later edits to an existing document without sources: MCP point edits (below), then `python $S fix OUT.docx`. `build` overwrites MCP edits.

## Token economy
- Never read `bsuir_docx.py`; this file documents everything needed.
- Do not add bulk text via MCP (one call per paragraph). Do not rewrite whole chunk files; StrReplace the fragment.
- Do not re-read files you wrote or that did not change; do not preview after text-only edits — trust `build` output.
- Large sources (notebooks, logs): extract only the needed text/figures with a short script, never Read whole files with outputs.

## cfg.json
```json
{"type":"lab|course|diploma","faculty":"Факультет компьютерных систем и сетей",
 "department":"Кафедра информатики","discipline":"...","topic":"...","number":"3",
 "student":"И. И. Иванов","group":"351001","teacher":"П. П. Петров","year":2026,
 "code":"БГУИР КП 1-40 01 01 012 ПЗ"}
```
Optional: `signers` [[label,name],...] (overrides student/teacher; diploma: Студент, Руководитель, Консультанты:, Нормоконтролер, Рецензент), `head`+`dept_short` (diploma "К защите допустить"), `abstract` (lines by \n; RU: РЕФЕРАТ page), `abstract_header`, `assignment_pages` (placeholder pages for бланк задания), `toc` (default true for course/diploma, false for lab), `section_new_page` (default true: every level-1 heading starts a new page), `title_block` (override, `**bold**`, \n; add `Вариант N` here for labs), `header`, `city`, `sign_tab_cm`.
Default `title_block` — lab: `**ОТЧЕТ** / по лабораторной работе № {number} / по дисциплине «{discipline}» / на тему / «{topic}»`; course: `**ПОЯСНИТЕЛЬНАЯ ЗАПИСКА** / к курсовому проекту / по дисциплине «…» / на тему / **ТЕМА**`; diploma: `**ПОЯСНИТЕЛЬНАЯ ЗАПИСКА** / к дипломному проекту / на тему / **ТЕМА**` (` / ` = line break `\n`). Signers default: lab `Выполнил: студент гр. {group}` / `Проверил:`; course `Студент гр. {group}` / `Руководитель`.
Names on the title page are full, not abbreviated. Code format: `БГУИР ДП|ДР|КП|КР 1-XX XX XX [XX] NNN ПЗ`.

## Markup (one line = one paragraph)
| Line | Result |
|---|---|
| `# 1 Название` / `## 1.1 ...` / `### 1.1.1 ...` | section (auto CAPS, new page) / subsection / пункт; Word auto-numbering (typed number optional, checked by lint); in appendices `## А.1 ...` stays as typed |
| `#! Введение` | centered unnumbered heading in contents: Введение, Заключение, Список использованных источников |
| `#= Текст` | centered heading not in contents |
| `#@ А \| обязательное \| Заголовок` | ПРИЛОЖЕНИЕ А (обязательное / рекомендуемое / справочное), new page |
| `- пункт;` or `1) ` / `а) ` / `  - ` or `  а) ` / `  1) ` | list item / nested level (Word lists, numbering restarts per list) |
| `$$ P=U^2/R, $$ (2.1)` | centered formula (UnicodeMath: `x_1`, `a/b`, `√(x)`, `∑_(i=1)^n`; no spaces inside products like `k(k-1)`), number at right |
| `где P – мощность, Вт;` + `  U – напряжение, В.` | explanation after formula |
| `![](img.png){w=15}` then `Рисунок 2.1 – Название` | centered figure (width cm ≤ 16.5), caption |
| `Таблица 2.1 – Название` then `\| a \| b \|` rows | caption + table (1st row = head; empty cell → «–») |
| ```` ``` ```` block | code listing (Courier New 10, no hyphenation) |
| `@code path` | listing with the contents of a text file |
| `\newpage` | page break |
Code listings get an empty paragraph before and after when they border ordinary text.
Inline: `**bold**`, `*italic*`, `U_{вх}`, `10^{3}`, placeholder `[[...]]`. Text: `[1]` citations, «ёлочки», « – » en dash.

## MCP (user-word-document-server)
- Inspect: `get_document_outline`, `find_text_in_document`, `get_paragraph_text_from_document` (avoid `get_document_text`/`get_document_xml` on big docs).
- Edit: `search_and_replace`, `insert_line_or_paragraph_near_text` (`line_style`), `insert_header_near_text` (`header_style`), `delete_paragraph`.
- Append single items: `add_heading` (level 1-3, text without number – numbered automatically), `add_paragraph` with `style` ∈ {StructHeading, PlainHeading, AppendixHeading, FigureCaption, TableCaption, Formula, Where, Sublist, Code}; `add_picture` (width inches), `add_table`, `merge_table_cells*`.
- List items via MCP: `add_paragraph`/`insert_line_or_paragraph_near_text` with text `– …`, `1) …`, `а) …` (style Normal, or Sublist for nested); `fix` turns them into Word lists.
- Never pass font_name/font_size/bold/color — styles carry formatting. Skip list/`highlight_table_header`/shading tools (non-standard look).

## Structure
- **lab**: title (with `Вариант N` if given) → `# 1 Цель работы` → sections per методичка/user → `# N Выводы` → sources (if cited). Each level-1 section on a new page, no contents.
- **course**: title → СОДЕРЖАНИЕ → (РЕФЕРАТ) → (задание) → Введение → numbered sections → Заключение → Список использованных источников → Приложения → (ведомость).
- **diploma**: title (`head`, all signers) → СОДЕРЖАНИЕ → РЕФЕРАТ (850–1200 chars, 1 page; header line `ТЕМА : дипломный проект / И. О. Фамилия. – Минск : БГУИР, 2026. – п.з. – 79 с., чертежей (плакатов) – 6 л. формата А1.`) → задание (`assignment_pages: 2`) → перечень обозначений (if any) → Введение (≤2 pages: goal, tasks, section summary) → main sections → ТЭО (≤18 %) → охрана труда/энергосбережение (5–7 %) → Заключение (≤2 pages, «разработана», «исследованы»...) → sources → приложения → перечень элементов → ведомость. Target 60–80 pages.

Writing rules (numbers, formulas, tables, figures, bibliography, appendices): read [reference.md](reference.md) before writing content.
