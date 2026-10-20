<!-- title: fix(xlsx): let fitToWidth / fitToHeight imply fitToPage -->

Closes #596

### What was wrong

The writer emitted the fit-to-page counts only when `fitToPage: true` was also set. Without it, they were dropped without a warning:

```ts
for (const pageSetup of [
  { fitToWidth: 1, fitToHeight: 0 },
  { fitToHeight: 2 },
  { fitToPage: false, fitToWidth: 1, fitToHeight: 0 },
]) {
  await writeXlsx({ sheets: [{ name: "Sheet1", rows: [["x"]], pageSetup }] })
}
```

```xml
<!-- { fitToWidth: 1, fitToHeight: 0 } — before: no <sheetPr>, no <pageSetup> -->
<!-- after -->
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<pageSetup fitToWidth="1" fitToHeight="0"/>

<!-- { fitToHeight: 2 } — before: no <sheetPr>, no <pageSetup> -->
<!-- after -->
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<pageSetup fitToHeight="2"/>

<!-- { fitToPage: false, fitToWidth: 1, fitToHeight: 0 } — before and after: neither element -->
```

So the sheet printed at 100%. Reading the file back gave a `pageSetup` with only the default margins. With `fitToPage: true` added, the output was correct and it still is.

The reader already treats either count as turning fit mode on. The writer also already lets `firstPageNumber` imply `useFirstPageNumber`. Fit to page was the only setting where the writer required the flag.

### Why the tests didn't catch it

Every fit-to-page test in `print-settings.test.ts`, `real-world.test.ts`, `xlsx-write-read-parity.test.ts` and the `page-setup-attributes.test.ts` register set `fitToPage: true` together with the counts. No test wrote the counts on their own. The reader test for a lone `fitToHeight` (`coverage-xlsx-worksheet.test.ts`) parses hand-written XML, so it never touches the writer.

### What landed

- A small `isFitToPage(ps)` helper in `worksheet-writer.ts`: `fitToPage ?? (fitToWidth !== undefined || fitToHeight !== undefined)`. The `<pageSetUpPr>` check in `writeWorksheetXml` and the count attributes in `serializePageSetup` both use it. An explicit `fitToPage: false` still wins and writes neither the flag nor the counts, because the counts only take effect in fit mode.
- Doc comments on `PageSetup.fitToPage`, `fitToWidth` and `fitToHeight` describe the implication and say that `0` means "as many pages as needed".
- `writeXlsxStream` / `writeXlsxStreamSheets` have no `pageSetup` support, so nothing changes there. `docs/PARITY.md` already says every `CT_PageSetup` attribute round-trips, so this brings the code in line with it and the page needs no edit.

Tests: three of the four new tests in `print-settings.test.ts` fail against `main`: counts only, `fitToHeight` only, and the round trip through `readXlsx`. The fourth checks that `fitToPage: false` with counts writes nothing, which already worked and is there to stop the implication from overriding an explicit `false`. `bun run test` passes on Node 24.22.0, with 240 files and 10,702 tests.

### Checked in Excel

In Excel for Mac, with a 20-column sheet written by `main` and by this branch. Both files open without a repair prompt.

- **After:** Excel's page setup reports fit to 1 page wide with height automatic. The Page Layout tab shows Width "1 page" and Height "Automatic".
- **Before:** scaling is 100%.

The XML matches what the existing `fitToPage: true` path already writes, so this change only makes that path reachable without the flag.

### Not verified

- **Print preview.** I checked the page setup, not that all 20 columns actually land on one page width.
- **LibreOffice and Windows Excel.** Not opened.
