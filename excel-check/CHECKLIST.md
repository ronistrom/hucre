# Excel checks

Each pair is the same input, written by `main` (**before**) and by the fix
branch (**after**). Open both in Excel. For every file, first note whether
Excel opens it cleanly or asks _"We found a problem with some content… Do
you want us to try to recover as much as we can?"_. A repair prompt on an
**after** file is a failure, whatever else it shows.

"Before" is what the issue predicts; most of it was never checked in Excel,
which is the point of doing this. Write down what you actually see.

| #   | File                  | Look at                                                                               | Expected after                                                                                                                                                                                                                                                                             |
| --- | --------------------- | ------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 594 | `594-font-inherits-*` | Click A1 and read the font box on the Home tab                                        | A1 is **Arial 10**, bold, dark blue. Before: not Arial 10 (Calibri, or the font box blank)                                                                                                                                                                                                 |
| 595 | `595-cf-color-argb-*` | A1:A10 colour scale, B1:B10 data bars                                                 | Opens cleanly; A runs red → green, B shows blue bars. Before: may prompt for repair, or show no/odd colours                                                                                                                                                                                |
| 596 | `596-fit-to-page-*`   | Page Layout tab → Scale to Fit (or File → Page Setup → Page)                          | **Fit to 1 page wide**, height automatic; print preview shows all 20 columns on one page width. Before: 100%, columns split over pages                                                                                                                                                     |
| 597 | `597-xlfn-prefix-*`   | D1:D5                                                                                 | No `#NAME?`. D1 `2`, D2 `small`, D3 `a,b,c`, D4 `b`, D5 `a`. FILTER/SORT may show with `@` in the formula bar and only their first value, because no dynamic-array metadata is written — that is separate from this fix                                                                    |
| 598 | `598-image-size-*`    | Click the blue bar → Picture Format → size; right-click → Format Picture → Properties | Bar at B2, **5.29 cm × 1.06 cm** (2.08″ × 0.42″), and Properties most likely "Move but don't size with cells". Before: stretched into a 3-column × 5-row box                                                                                                                               |
| 599 | `599-formula-error-*` | A1, B1, A2, B2                                                                        | A1 `#DIV/0!` (error), B1 `#DIV/0!` (text), A2 `TRUE`, B2 `FALSE`. hucre asks Excel to recalculate on open, so **both files will likely look the same in Excel**; this pair only checks the after file opens cleanly. The fix matters for apps that use cached values without recalculating |
| 601 | `601-row-style-*`     | Row 2 across empty columns; type text into E2                                         | Whole row 2 yellow; text typed into E2 is bold on yellow; A2 bold on yellow. Before: row 2 only taller (24 pt), no yellow                                                                                                                                                                  |
| 602 | `602-rich-text-*`     | A1                                                                                    | "**Bold red** then normal" with mixed formatting, opens cleanly. Before looks the same — the difference is where the text is stored, which Excel doesn't show                                                                                                                              |

## The other direction: a file Excel wrote

#599, #600 and #601 also change what hucre **reads**. To check
them against real Excel output, make a new workbook in Excel:

- **A1** `=1/0` and **B1** `="#DIV/0!"` (#599)
- **A3** `15.1.2024 12:00` formatted as a date and time, **B3** `6:30` formatted `hh:mm`, **C3** the number `60` formatted as a date (#600)
- Select the **row 5 header**, make it bold with a fill colour (#601)

Save it as `.xlsx` and run `excel-check/run.sh read that-file.xlsx`. Expected after:

- #599: A1 has `formulaResultType: "error"`, and B1 doesn't.
- #600: with `dates: "serial"`, A3, B3 and C3 are plain numbers (C3 exactly `60`).
- #601: `row 5 def` has a `style`.
