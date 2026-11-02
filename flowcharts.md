# Flowcharts (схемы алгоритмов, ГОСТ 19.701–90)

Script `F=$HOME/.cursor/skills/bsuir-word-docs/scripts/flowchart.py` (Pillow). One description file → `OUT.png` (300 dpi, true size, for Word) + `OUT.drawio` (editable in draw.io; from there File → Export as → VSDX for Visio).

## Workflow
1. Write the description `chart.txt` (syntax below) in the temp folder; one file per figure.
2. `python $F chart.txt RESULT_DIR/схемы/рисунок_2_1` — put outputs next to the .docx (`схемы` folder) so the user keeps the editable .drawio. Output: size in cm and `warn:` lines.
3. Resolve warnings: height > 23 cm → split into sheets with `ref` connectors (one file per sheet, caption `Рисунок 2.1, лист 2`); width > 16,5 cm → shorten texts or flatten nesting.
4. In the chunk: `![](RESULT_DIR/схемы/рисунок_2_1.png)` without `{w=}` (true size), then `Рисунок 2.1 – Схема алгоритма ...`, reference it in the text.
5. Tell the user where the .drawio files are. After editing in draw.io: export PNG and replace in Word via right-click → «Изменить рисунок» (keeps size and caption).

## Syntax (one line = one block, 2-space indent = nesting)
| Line | Symbol |
|---|---|
| `start` / `start Текст` | terminator «Начало» (no input) |
| `end` / `end Текст` | terminator «Конец» (no output) |
| `Текст` or `do Текст` | process (rectangle) |
| `io Ввод a, b` | data (parallelogram): input/output |
| `call Сортировка(m, n)` | predefined process (subroutine) |
| `prep s = 0, k = 0` | preparation (hexagon) as a plain block |
| `ref А` | connector (circle): first/last line of a sheet to continue on another sheet |
| `if условие` + body, optional `else` + body | decision: body – down, else (or bypass) – right; labels «да»/«нет» (swapped by `yes: right`) |
| `while условие` + body | precondition loop: «да» – body, back line enters the left vertex, «нет» – exit right |
| `repeat` + body + `until условие` | postcondition loop: «нет» – back, «да» – exit; `while условие` instead of `until`: «да» – back |
| `for i = 1, n` + body (`for заголовок \| конец`) | loop limit symbols (начало/конец цикла); end text defaults to the header |
| `loop i = 1, n` + body | loop with modification symbol (hexagon), back line into its left vertex |
| `... // текст` | comment (dashed line + bracket) on any non-decision block |
Options (first lines only): `a: 20` symbol height mm (width 1.5a, default 20; 15 for compact charts), `font: 12` pt, `gap: 8`, `hgap: 8` mm, `uniform: 1` (all symbols get one width and height; decisions keep that width and grow only in height, so keep their words short), `yes: right` (`if`: «да» – right, «нет» – down; phrase conditions so the main path is «нет», e.g. «Токен неверен?»). `;` starts a comment line. `\n` forces a line break in a block.
User preference: always write `uniform: 1` and `yes: right`.

Example:
```
start
io Ввод n, массива a
prep s = 0
for i = 1, n
  if a[i] > 0
    s = s + a[i] // сумма положительных
io Вывод s
end
```

## GOST rules applied by the script (do not fight them)
- Symbol size a × 1.5a (grows in 5 mm steps if the text does not fit), Times New Roman, monochrome.
- Flow top→bottom and left→right without arrows; lines going up or left end with an arrow; a line entering another line ends with an arrow at the junction. Arrows are open (two strokes), not filled.
- Decision outputs labeled «да»/«нет»; lines join below the construct.
- Text in symbols: short, formulas as in code (`s = s + a[i]`), lines break only after `,` `;` or between words.
- Caption under the figure «Рисунок N – Схема алгоритма <чего>»; a chart split over sheets keeps one number: «Рисунок 2.1, лист 2».
