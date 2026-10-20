<!-- title: fix(xlsx): give a partial cell font the default font's name and size -->

Closes #594

### What was wrong

`defaultFont` is written as `fonts[0]`, and `Workbook.defaultFont`'s doc says it is "the entry every cell format inherits from unless it names another". A cell's `fontId` does not work that way. It picks one complete `<font>`, so a cell font that gave only part of itself was written without a face or size:

```ts
await writeXlsx({
  defaultFont: { name: "Arial", size: 10 },
  sheets: [{ name: "Sheet1", rows: [[{ value: "Bold header", style: { font: { bold: true, color: { rgb: "FF1F3A5F" } } } }]] }],
})
```

```xml
<!-- before -->
<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FF1F3A5F"/></font></fonts>
<!-- after -->
<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FF1F3A5F"/><sz val="10"/><name val="Arial"/></font></fonts>
```

Reading it back with `readStyles: true`, the cell's font was `{ bold: true, color: { rgb: "1F3A5F" } }` before and is `{ bold: true, color: { rgb: "1F3A5F" }, size: 10, name: "Arial" }` after. According to the issue, LibreOffice drew the old file's bold cell in Cambria 11pt.

### Why the tests didn't catch it

No test wrote a partial font next to a `defaultFont` and looked at the `<font>` that came out. The parity register did write one: its `columns` probe has `{ bold: true }` and the fixture has `defaultFont: { name: "Georgia", size: 13 }`. But it expected `font: { bold: true }` back, so it pinned the bug.

### What landed

- `inheritBaseFont` in `styles-writer.ts`, called from `registerXf` before `addFont`. A missing `name` or `size` comes from the base font. When the name is inherited, the base's `family`, `charset` and `scheme` come with it. A font that names its own face does not get them, because `scheme="minor"` would make Excel draw the theme font instead. Colour and the other parts are not inherited. A font that already has both a name and a size is passed through without a copy, because the streaming writers call this once per cell.
- Differential formats (`addDxf`, used by conditional formatting) are unchanged. There, a part that is left out means "keep the cell's own".
- The fix reaches `writeXlsx`, the streaming writers (Calibri 11, since they take no `defaultFont`) and the round-trip writer, because they all use the same collector.
- `WriteOptions.defaultFont` gets a doc comment. `docs/PARITY.md` has a short section on what comes back. The parity register's `columns` expectation now includes Georgia 13, with a comment that explains why.

Tests: `test/font-inherits-default.test.ts` adds 9 tests, and 7 of them fail against `main`. The 2 that pass on `main` are guards: the conditional formatting font stays unmerged, and the caller's style object is not mutated. The changed parity expectation also fails without the fix. `bun run test` passes on Node 24.22.0, with 241 files and 10,707 tests.

### Checked in Excel

In Excel for Mac, with the issue's case written by `main` and by this branch. Both files open without a repair prompt. In both, Excel reports A1 as Arial 10, bold, dark blue. Excel fills a font's missing `name` and `size` from font 0 itself, so the bug does not show in Excel. The difference shows in apps that don't fill the gap that way.

### Not verified

- **LibreOffice.** Not opened. The claim that it draws the old output's bold cell in Cambria 11pt comes from the issue.
- **Windows Excel.** Only Excel for Mac was checked.

Open question: rich text runs have the same problem. A run font without `name`/`size` is written as given. I left that out of this PR.
