<!-- title: fix(xlsx): write rich text into the shared string table in shared mode -->

Closes #602

### What was wrong

With `stringMode: "shared"`, which is the default, plain text goes into `xl/sharedStrings.xml`. A rich text cell was always written inline, whatever the mode:

```ts
await writeXlsx({
  stringMode: "shared",
  sheets: [{
    name: "S",
    rows: [["Plain"]],
    cells: new Map([["1,0", { richText: [{ text: "Bold", font: { bold: true } }, { text: " normal" }] }]]),
  }],
})
```

```xml
<!-- before: sheet1.xml -->
<row r="2"><c r="A2" t="inlineStr"><is><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> normal</t></r></is></c></row>
<!-- before: sharedStrings.xml -->
<sst … count="1" uniqueCount="1"><si><t>Plain</t></si></sst>

<!-- after: sheet1.xml -->
<row r="2"><c r="A2" t="s"><v>1</v></c></row>
<!-- after: sharedStrings.xml -->
<sst … count="2" uniqueCount="2"><si><t>Plain</t></si><si><r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> normal</t></r></si></sst>
```

The old output is valid. But tools that only read the shared string table, such as template engines, find/replace and string indexing, never saw the rich cells. It also isn't what Excel writes. An `<si>` can hold the same `<r>` runs as an inline `<is>` (ECMA-376 Part 1 §18.4.8, §18.4.4, §18.3.1.53).

### Why the tests didn't catch it

Several tests asserted the inline form for the default mode, so they pinned the behaviour: `coverage-xlsx-writer.test.ts` (`t="inlineStr"` on a styled rich cell), `rich-text-write.test.ts` (the mixed plain/rich sheet), `one-cell-serializer.test.ts` (`<is>` in the sheet XML), `xml-space-preserve.test.ts` (looked for the rich run's `<t>` in the sheet XML) and `batch-fixes.test.ts`. The round-trip tests passed either way, because the reader already turns both `<is>` and `<si>` runs into `richText`.

### What landed

- `SharedStringsCollector` gets `addRich(runs)`. The dedup key is the serialized `<r>` XML (`serializeRichTextRuns`), kept in a map of its own. Two cells with equal runs share one entry. A plain string with the same text, or the same text with different formatting, gets its own entry.
- `writeSharedStringsXml` writes a rich entry as `<si>` holding those runs.
- `serializeCell` writes a rich cell as `t="s"` with the shared index when `stringMode` is `"shared"`. The style index is kept. The comment there cites #602.
- Inline mode is byte-identical: the SHA-256 of an inline-mode workbook with a styled rich cell is the same before and after.
- `saveXlsx` always writes shared strings, so it now puts rich text in the table too. A rich cell read from an inline-only file moves into the table on save, which matches Excel.
- The streaming writers take plain cell values and can't carry rich text, so nothing changes there. The reader is unchanged. `docs/PARITY.md` doesn't say where rich text is stored, so it is unchanged too.
- The tests that pinned the inline form now either expect the shared form or pass `stringMode: "inline"` (or `inlineStrings: true`) when the test is about the inline form's runs.

Tests: a new `test/rich-text-shared-strings.test.ts` with seven tests. Five of them fail against `main`: shared `<si>` plus `t="s"`, the style index, dedup apart from plain text, the shared string part being written when rich text is the only text, and `saveXlsx`. The other two passed already: the read-back round trip, and the exact inline output. The four updated assertions in the existing tests also fail against `main`, nine failures in total. `bun run test` passes on Node 24.22.0, with 241 files and 10,705 tests.

### Checked in Excel

In Excel for Mac, with A1 as "**Bold red** then normal" and A2 as plain text, written by `main` and by this branch. Both files open without a repair prompt and look the same: the first run is bold and red, the rest is plain. In the after file the rich text is in `sharedStrings.xml` as an `<si>` with two `<r>` runs, and A1 is `t="s"`. In the before file it is an inline `<is>`. So Excel reads the new form as it reads the old one.

### Not verified

- **LibreOffice and Windows Excel.** Not opened.
- **The tools this is for.** Template engines and find/replace tools that only read the shared string table were not tried.
