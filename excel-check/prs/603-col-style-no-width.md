<!-- title: fix(xlsx): give a <col> without a width Excel's default width -->

Closes #603

### What was wrong

A `ColumnDef` that sets `style`, `numFmt`, `outlineLevel` or `hidden` but no `width` was written as a `<col>` with no `width` attribute. `width` is optional on `CT_Col` (ECMA-376 Part 1 §18.3.1.13), but Excel reads a `<col>` without one as zero wide, so the column was hidden. Excel opens the file without any warning; the column is simply missing.

```ts
await writeXlsx({
  sheets: [{
    name: "Sheet1",
    columns: [{}, { style: { fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FF9999" } } } }, {}, { outlineLevel: 1 }],
    rows: [["A1", "B1", "C1", "D1"]],
  }],
})
```

```xml
<!-- before -->
<cols><col min="2" max="2" style="1"/><col min="4" max="4" outlineLevel="1"/></cols>
<!-- after -->
<cols><col min="2" max="2" width="9.140625" style="1"/><col min="4" max="4" width="9.140625" outlineLevel="1"/></cols>
```

In Excel for Mac (`column width` / `hidden` from Excel itself):

```
before: B width=0 hidden=true    C width=8 hidden=false    D width=0 hidden=true
after:  B width=8.33 hidden=false  C width=8 hidden=false  D width=8.33 hidden=false
```

`writeXlsxStream` had the same problem for `outlineLevel` (it does not write column styles).

Excel itself always writes a width, even for a column that only has a format. Filling a column in Excel for Mac, without touching its width, wrote `<col min="2" max="2" width="8.83203125" style="1"/>`.

### Why the tests did not catch it

They check the XML hucre writes, and round trips through hucre's own reader. That reader treats a missing width as "no width", so the round trip looked right. Nothing checked how Excel reads the file. One test in `coverage-xlsx-writer.test.ts` even pinned the old output, `<col min="2" max="2" outlineLevel="2"/>`.

### What landed

- **Both writers** (`worksheet-writer.ts` and `stream-writer.ts`) now write a width on any `<col>` that has none of its own. It is the sheet's `defaultColWidth` when there is one (same units, §18.3.1.81), otherwise `DEFAULT_COL_WIDTH = 9.140625`. That is Windows Excel's default for Calibri 11, shown as 8.43 characters, and the value XlsxWriter writes for a formatted column with no width. The reason is in the doc comment on the constant.
- **`customWidth` is not set**, so the file still says the width is the default and not one the author chose. Excel does the same.
- A width the author sets, or `autoWidth` computes, is written as before, with `customWidth`.
- A column that says nothing still gets no `<col>` at all.
- **`docs/PARITY.md`**: a new section, "A column with no width comes back with one". After a write and a read, such a column has `width: 9.140625` (or the sheet default), where before it had none.

Tests: six new ones in `test/column-and-sheet-defaults.test.ts`. Four fail against `main`:

- a styled column gets the default width, without `customWidth`
- an outline-level-only column gets it too
- the sheet's `defaultColWidth` is used when it has one
- the stream writer does the same

The two that pass on `main` check what must not change: an explicit width keeps `customWidth`, and the column style still round-trips. I updated the assertion in `coverage-xlsx-writer.test.ts` that pinned the old output. `bun run test` passes, with 240 files and 10,704 tests.

### Not verified

- **Windows Excel.** I checked in Excel for Mac only. There, `9.140625` shows as 8.33 characters next to the Mac default of 8.0, so these columns are a little wider than the columns around them. On Windows it should show as exactly 8.43, the default there, but I have not opened it there.
- **LibreOffice and Google Sheets.** Not opened.
- **The font.** The fallback does not depend on the workbook's default font. Excel for Mac wrote the same `8.83203125` for a Calibri 11 and an Arial 10 workbook, so one constant looked enough, but other fonts and Windows were not tried.
