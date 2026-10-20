<!-- title: fix(xlsx): write the cells of a table's total row -->

Closes #604

### What was wrong

A table with `showTotalRow: true` got its total row only in the table part (`totalsRowLabel`, `totalsRowFunction`, `totalsRowFormula` in `xl/tables/table1.xml`). The sheet's total row had no cells at all.

```ts
await writeXlsx({
  sheets: [{
    name: "Sheet1",
    rows: [["Name", "Qty"], ["a", 1], ["b", 2], ["c", 3]],
    tables: [{
      name: "Table1", range: "A1:B5", showTotalRow: true,
      columns: [
        { name: "Name", totalLabel: "Total" },
        { name: "Qty", totalFunction: "custom", totalFormula: "SUM(Table1[Qty])*2" },
      ],
    }],
  }],
})
```

```xml
<!-- before: sheet1.xml has no row 5 -->
<!-- after -->
<row r="5"><c r="A5" t="s"><v>5</v></c><c r="B5"><f>SUM(Table1[Qty])*2</f></c></row>
```

In Excel for Mac 16:

- `totalFunction: "custom"`: Excel asks to repair the file, whether or not the formula uses newer functions. The repair deletes the total row, and the table becomes `A1:B4`.
- `totalFunction: "sum"`: the file opens, but row 5 is empty. Neither the label nor the sum shows.

Excel writes a total row twice. The table part describes it, and the sheet holds it as ordinary cells, which is what Excel reads. Here is what Excel for Mac wrote for the same table:

```xml
<c r="A5" t="s"><v>5</v></c>                                    <!-- "Total" -->
<c r="B5"><f>SUBTOTAL(109,Table1[Qty])</f><v>6</v></c>           <!-- totalsRowFunction="sum" -->
<c r="B5"><f>SUM(Table1[Qty])*2</f><v>12</v></c>                 <!-- custom -->
<c r="B5"><f>_xlfn.XLOOKUP("c",Table1[Name],Table1[Qty])</f>…    <!-- custom, newer function -->
```

With a newer function, the `_xlfn.` prefix is in both `<totalsRowFormula>` and the cell. #597 couldn't confirm in Excel that its total-row prefix worked, because any custom total row already needed repair.

### Why the tests did not catch it

They checked the table part's XML and round-tripped it through hucre's own reader. That reader reads `totalsRowFunction` and `totalsRowLabel` back from the table part, so the round trip looked right. Nothing looked at the sheet's total row, and nothing opened the file in Excel. The one existing total-row test in `tables.test.ts` hid the problem: its `rows` already include a hand-written `["Total", 30]` row.

### What landed

- **`table-writer.ts`** has a new `totalRowCells(table, range)` that returns the cells Excel writes:
  - the label as a string;
  - a built-in function as `SUBTOTAL(n,Name[Column])`, with Excel's function numbers (average 101, countNums 102, count 103, max 104, min 105, stdDev 107, sum 109, var 110);
  - a custom total as its formula.

  The reference uses the table's `displayName`, as Excel does: a table named "Odd Tbl" becomes `Odd_Tbl[...]`. `[ ] # '` in a column name are escaped with `'`, so `a[b]` becomes `Odd_Tbl[a'[b']]` and `it's` becomes `[it''s]`. Every one of these forms was checked against what Excel wrote.
- **`worksheet-writer.ts`**: `resolveRows` merges these cells into the grid. It places them on the last row of the table's range, or of the computed range when there is none. Both `writeXlsx` and `saveXlsx` go through `resolveRows`.
- **A total-row cell the caller already filled is kept.** A workbook read back from Excel already has these cells, with cached results. `openXlsx` → `saveXlsx` on Excel's own files still writes row 5 once, with its `<v>`.
- **No cached value is written.** hucre always writes `fullCalcOnLoad`, so Excel computes the total when it opens the file.
- **`docs/PARITY.md`** has a new section, "A table's total row comes back as cells". After a write and a read, the total row shows up in `rows` and `cells` even when the caller never wrote it.

Tests: six new ones in `test/tables.test.ts`. Four fail against `main`:

- a built-in function writes the label and `SUBTOTAL(109,Table1[Qty])`
- a custom total writes its formula into the cell
- every function's SUBTOTAL number, the `displayName` and the escaped column names, for a table placed at C3
- the total row lands below the data when the range is computed

The two that pass on `main` check what must not change: a total-row cell the caller filled is kept, and nothing is written for a column with no total or for a table without `showTotalRow`. `bun run test` passes, with 240 files and 10,704 tests.

### Checked in Excel

All in Excel for Mac 16. "Opens cleanly" means no repair prompt.

| file | result |
| --- | --- |
| `main`, custom total | repair prompt; the repair deletes the total row |
| `main`, `sum` total | opens; row 5 empty |
| this branch, `sum` total | opens cleanly; Total / 6. Excel reports the column's `totals calculation` as sum |
| this branch, custom total | opens cleanly; Total / 12. `totals calculation` is custom |
| this branch, table at C3 with all eight functions, a label and the escaped names above | opens cleanly; every total computed |
| this branch merged with #597's branch, custom `XLOOKUP` total | opens cleanly; shows 3 |

The last row works because #597 adds the prefix when it writes a cell, so the total-row cell gets `_xlfn.` with no code here. The two branches merge without conflicts.

### Not verified

- **Windows Excel, LibreOffice and Google Sheets.** None was opened. I checked in Excel for Mac only.
- **`writeXlsxStream`.** It does not write tables at all, so it is unchanged.
- **A table without a header row.** `TableDefinition` has no option for one, so I didn't try it.
- **A total-row cell the caller filled with something other than the table's total.** It is kept as the caller wrote it, and I did not check what Excel makes of a total row whose cell disagrees with `totalsRowFunction`.
