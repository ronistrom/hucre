<!-- title: feat(xlsx): write and read a row format on RowDef -->

Closes #601

### What was missing

`RowDef` had `height`, `hidden`, `outlineLevel` and `collapsed`, but nothing for a row-level format. Excel writes a row formatted as a whole as `<row s="N" customFormat="1">` (ECMA-376 Part 1 §18.3.1.73), and the reader dropped both attributes. So a row format in an Excel file was lost on a read-write round trip, and there was no way to write one. `ColumnDef.style` has done the column equivalent (`<col style="N">`) since #439.

The issue's input, a height and a bold style on row 2:

```ts
await writeXlsx({
  sheets: [{
    name: "S",
    rows: [["Name", "Qty"], ["Total", 3]],
    rowDefs: new Map([[1, { height: 20, style: { font: { bold: true } } }]]),
  }],
})
```

```xml
<!-- before -->
<row r="2" ht="20" customHeight="1"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>3</v></c></row>
<!-- after -->
<row r="2" ht="20" customHeight="1" s="1" customFormat="1"><c r="A2" t="s" s="1"><v>2</v></c><c r="B2" s="1"><v>3</v></c></row>
```

Reading that file back with `readStyles: true`:

```
before: rowDefs.get(1) = {"height":20}
after:  rowDefs.get(1) = {"height":20,"style":{"font":{"bold":true}}}
```

And `test/fixtures/excel-styleonly.xlsx`, which Excel wrote with four formatted rows (rows 3 and 4 have no cells, so the `<row>` is the only place their format lives):

```
before: rowDefs = []
after:  rowDefs = [[0, {style: {font: {bold: true, …}}}],
                   [1, {style: {numFmt: "0.00", fill: {pattern: "solid", fgColor: {rgb: "DDEBF7"}, …}}}],
                   [2, …same…], [3, …same…]]
```

### What landed

- `RowDef.style?: CellStyle`, with a doc comment.
- **Writer.** `rowAttributes()` in `worksheet-writer.ts` is shared by `writeXlsx` and the streaming `RowSerializer`, so both writers emit `s` and `customFormat="1"`. They are written only when the style resolves to a non-default xf, the same rule `<col style>` follows, so a row definition with no style writes exactly the bytes it did before. A test pins that output. A row that has a format and no cells still gets a self-closing `<row … s customFormat/>`. `writeWorksheetXml` now uses `hasRowAttributes()` instead of its own inline copy of the check.
- **The row style is also put on the row's cells.** This follows the column precedent, and it is what Excel itself writes: the fixture has the row's `s` on every cell in the row. Any cell in that row with no style of its own gets the row style. That covers cells from `rows`, from `data[]` (including the generated header row), and from a `cells` override that adds a cell to the row. A cell that has its own style keeps it.
- **A row style wins over a column style** on a cell that neither styles directly, and the column's `numFmt` doesn't apply there either. The order is cell, then row, then column, which is how Excel is understood to resolve a cell's format. The streaming writer used to merge a column's `numFmt` into any style a cell ended up with. It no longer does that to a row style, so `writeXlsx` and `writeXlsxStream` produce the same cell styles. A test checks both writers.
- **Reader.** When `customFormat` is `1`/`true` and styles are being read, `s` becomes `RowDef.style`. It is resolved and cloned once per row, as column styles are, so editing one row's style doesn't change another's. A bare `s` without `customFormat` is ignored, because the spec says `s` only applies to the row when `customFormat` is set. Without `readStyles`, the row comes back without the style, like everything else that depends on styles.
- `cloneSheet` deep-copies the row style.
- `customFormat` is removed from the `REVIEWED` map in `scripts/spec-coverage.mjs`. `docs/SPEC-COVERAGE.md` is regenerated: `customFormat` moves out of the reviewed list, and the counts go 31 → 30 and 644 → 645. I checked that the script reproduces `main`'s committed file before the change and this one after.
- `docs/PARITY.md`: row and column definitions are at XLSX parity "a row or column format included", and ODS lists row formats among what neither direction models. In the README, a short paragraph and example for `rowDefs` `style`, and a "row formats" entry in the ODS table. `SheetBuilder.rowDef` and the `StreamWriterOptions.rowDefs` doc comments name the new field.
- The parity register's `rowDefs` probe now includes a numFmt-only style, which keeps it clear of #594's font change. The `cloneSheet` register now has a row style and checks that the copy doesn't share it.

Tests: 14 of the new or extended tests fail against `main`:

- 11 of the 14 in the new `test/row-style.test.ts`
- the Excel fixture test in `real-files.test.ts`
- the parity register round trip
- the `cloneSheet` detach check

The three in `row-style.test.ts` that pass on `main` all check something that should *not* happen: no `customFormat` on a row without a style, a bare `s` being ignored, and no style without `readStyles`. `bun run test` passes on Node 24, with 241 files and 10,713 tests.

### Not verified

I did not open the output in Excel or LibreOffice. In particular:

- **Precedence.** "Row wins over column" for an unstyled cell follows the cell, row, column order Excel is understood to use, but I have not checked it against Excel. It also only matters for cells hucre writes. For cells typed into the row later, Excel decides from `<row s>` and `<col style>` itself.
- **Rows without a style index.** `RowDef.style: {}` still gets its own xf (index 1), as `{}` already does on a cell or column. It is not treated as "no format".
