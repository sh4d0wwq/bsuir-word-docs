# СТП 01–2017 writing rules (layout is handled by the script)

## Headings, numbering
- Sections `1`, subsections `1.1`, пункты `1.1.1` (no trailing dot after number or heading). Пункты usually untitled: write `1.1.1 Текст пункта...` as a normal paragraph.
- Section headings CAPS (auto), subsections sentence case. No hyphenation, no underline; two sentences are separated by a dot.
- Contents, Введение, Заключение, Список, Реферат are unnumbered.

## Text
- Impersonal style: «получаем», «определяем», «находим»; not «мы получили», «будем иметь». Requirements: «должен», «следует», «не допускается».
- No general textbook/help/internet material copied into the work.
- Lists: simple list → `– item;` with the last item ending in `.`; complex items (sentences) → `1 Текст.` with capital letter; items referenced later → `а)`, `б)`; nested → `1)`. The intro phrase must not end with a preposition («состоит из:» ✗ → «В состав входят:»).
- Reference to a list item: «в пункте 1.7, б».
- In text write in words: minus, >, <, =, №, %, sin... when no number follows. Numbers 1–9 without units in words, >9 in digits; decimals with comma (`0,25`).
- Unit after the number with space: `600 Вт` (not «в 600 Вт»); range: `от 2 до 20 км`; limits: «должно быть не более 18 %». SI units only (ГОСТ 8.417).
- Ordinals: `2-м`, `1-го`, `3-й`.
- Subscripts from Russian abbreviations upright with dots: `K_{о.с}`, `U_{вх}`.

## Formulas
- Separate line, centered, blank line around (script). Numbered within the section `(2.7)`; ≤10 formulas → through numbering `(7)` allowed; in appendices `(Б.2)`. Number on the last line when a formula is broken.
- Before a formula introduced by a generalizing phrase use `:`; punctuation goes right after the formula (put `,`/`.` inside `$$ ... $$`).
- `где` from a new line without indent and without colon; each symbol `– explanation, unit;`, last ends with `.`. Units are not written after a letter-only formula (put them in the lead-in: «момент инерции J, кг·м², определяем по формуле»); with numeric substitution the unit follows the result as quoted text: `$$ J = 0,5·0,334 = 0,167 " кг·м²" $$`.
- Break a long formula on `=`, `+`, `−`, `×` (repeat sign on new line); never on division.
- References: «в формуле (2.1)», «подставляя выражение (3.6) в уравнение (3.2)».

## Figures
- All illustrations are «Рисунок». Place after the paragraph with the first reference; reference every figure: «на рисунке 2.1», «(см. рисунок 2)», «в соответствии с рисунком 5.1».
- Caption: `Рисунок 2.1 – Название` (section numbering) / `Рисунок 7` (through) / `Рисунок А.2` (appendix); centered, no final dot. Legend between picture and caption in one line: `1 – усилитель; 2 – датчик`.
- Recommended sizes ≈ 92×150 or 150×240 mm; readable without rotation or rotated 90° clockwise. Schemes: no frames/stamps, only relevant elements. Monochrome by default.
- Split figure: repeat `Рисунок 2.1, лист 2`.

## Tables
- `Таблица 1.2 – Заголовок` above, left-aligned, no dot; numbering like figures; appendix `Таблица Б.2`. Reference each table in text; explain conclusions from it.
- Column/row headings capitalized, nominative singular, no abbreviations; subheadings lowercase. Units via comma: `Время, с`. No «№ п/п» column (number inside the first cell: `1 Динамическая ошибка`). No diagonal split cells. Same decimals per column; missing data → `–`, no empty cells.
- Notes to a table: last row above the bottom line `Примечание – ...`.
- Continuation on a new page: `Продолжение таблицы 1.2` + repeated head (script repeats the head row).
- Little data → text with dot leaders instead of a table.

## Notes, footnotes
- `Примечание – Текст.` from the paragraph indent; several → `Примечания` then `1 ...`, `2 ...`.
- Footnote marks `1)` superscript (MCP `add_footnote_after_text`).

## Appendices
- Letters А, Б, В, Г, Д, Е, Ж, И, К, Л, М, Н, П, Р, С, Т, У, Ф, Х, Ц, Ш, Щ, Э, Ю, Я (no Ё З Й О Ч Ъ Ы Ь). Single appendix is still «А».
- Ordered by first reference in the text; each referenced («в приложении А»). Included in page numbering and contents.

## Список использованных источников (ГОСТ 7.1–2003)
Numbered `[n]` in order of first citation in the text; every entry cited. Lecture materials go at the end without citations. No Wikipedia/help. Spaces around `: ; – /`; initials `М. О.`; «Минск» in full; publication type lowercase.
```
[1] Гук, М. Процессоры Pentium II / М. Гук. – СПб. : Питер Ком, 1999. – 288 с.
[2] Кузелин, М. О. Современные семейства ПЛИС : справ. пособие / М. О. Кузелин, Д. А. Кнышев, В. Ю. Зотов. – М. : Горячая линия – Телеком, 2004. – 440 с.
[3] Технические средства диагностирования : справочник / В. В. Клюев [и др.]. – М. : Машиностроение, 1989. – 672 с.
[4] Берски, Д. Набор ЭСЛ-микросхем / Д. Берски // Электроника. – 1989. – № 12. – С. 21–25.
[5] Аксенов, О. Ю. Методика формирования выборок / О. Ю. Аксенов // Нейроинформатика-2004 : сб. науч. тр. – М. : МИФИ, 2004. – С. 215–222.
[6] Xilinx [Электронный ресурс]. – Режим доступа : http://www.plis.ru/.
[7] Embedded Microcontrollers : Databook / Intel Corporation. – Santa Clara, 1994.
```

## Schemes/algorithms in figures
Flowcharts per ГОСТ 19.701–90; electrical schemes per ГОСТ 2.701–2008 (Э1 structural, Э2 functional, Э3 principal); graphic document codes `ГУИР.421233.001Э1`.
