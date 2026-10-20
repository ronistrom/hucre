<!-- title: fix(xlsx): keep a formula's error result apart from a text result -->

Closes #599

### What was wrong

A formula cell's cached result goes into `formulaResult` as a plain string, whether the file stored it as an error (`t="e"`) or as text (`t="str"`). `#DIV/0!` is a valid text result too, so the two cells below read back the same:

```xml
<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>
<c r="B1" t="str"><f>"#DIV/0!"</f><v>#DIV/0!</v></c>
<c r="C1" t="e"><v>#N/A</v></c>
```

```
// before
A1 {"value":"#DIV/0!","type":"formula","formula":"1/0","formulaResult":"#DIV/0!"}
B1 {"value":"#DIV/0!","type":"formula","formula":"\"#DIV/0!\"","formulaResult":"#DIV/0!"}
C1 {"value":"#N/A","type":"error"}
// after
A1 {"value":"#DIV/0!","type":"formula","formula":"1/0","formulaResult":"#DIV/0!","formulaResultType":"error"}
B1 {"value":"#DIV/0!","type":"formula","formula":"\"#DIV/0!\"","formulaResult":"#DIV/0!"}
C1 {"value":"#N/A","type":"error"}
```

The writer emitted every string result as `t="str"`, so `readXlsx` → `writeXlsx` turned the error into text:

```xml
<!-- before -->
<c r="A1" t="str"><f>1/0</f><v>#DIV/0!</v></c>
<!-- after -->
<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>
```

B1 and C1 are written the same before and after. ECMA-376 Part 1 §18.18.11 ST_CellType keeps `e` (error) and `str` (formula string) separate.

### Why the tests didn't catch it

The round-trip test in `formula-result-types.test.ts` (from #497) only checked that the text `#DIV/0!` appeared somewhere in the rewritten sheet. It never checked the cell's `t`. Every read-side assertion compared `formulaResult` to the string, and that string is the same for both cell types.

### What landed

- A new field, `Cell.formulaResultType?: "error"`. The reader sets it only on a formula cell whose cached result is `t="e"`. Every other cell, including a hard-coded error with no formula, has no such field. Nothing that already exists changes value.
- The writer emits `t="e"` for a string `formulaResult` when `formulaResultType === "error"`, and `t="str"` otherwise. It never guesses from the token, because a formula can return the text `"#DIV/0!"`. If `formulaResult` is not a string, the writer ignores the field, since an error token is always text. The doc comment says so, and a test pins it.
- `cloneCell`, the worker serializer and deserializer, and the `FULL_CELL` fixture in `clone-sheet-coverage.test.ts` all carry the field.
- `docs/PARITY.md` (under "Cached formula results") now says the error and text results stay distinct.
- `writeXlsxStream` is unchanged. `StreamStyledCell` has no way to mark a streamed formula's result as an error, so string results there are still written as `t="str"`. PARITY.md says so. Supporting it would mean adding a field to `StreamStyledCell`, which I left out of this PR.

Tests: three of the five new tests fail against `main`: the read case, the read → write round trip, and the authored `{ formulaResultType: "error" }` write. The tightened #497 round-trip assertion fails too. The other two new tests pass on `main`: one checks that no other result type gets the field, the other that a non-string result ignores it. Adding the field to `FULL_CELL` also makes six `clone-sheet-coverage` tests fail at runtime on `main`, because `cloneCell` and the worker drop the field. With the field in the type, `tsc` would reject the fixture there as well. On Node 24.22.0, lint, format check and both typechecks pass. The full vitest run had 10,701 of 10,703 tests passing across 240 files. The two failures were timeouts in `html-import.test.ts` and `real-world.test.ts` while the machine was heavily loaded (load average around 29). The first run hit a similar timeout in `property-roundtrip.test.ts`. All three files pass when run alone (71 tests).

After review: a cell a shared formula covers (`<f t="shared" si="0"/>`, read as `formula: ""`) lost its cached error, because the error arm tested `formula` for truthiness. It now tests `!== undefined`. Two more tests cover the read and the rewrite; both fail without that commit. `bun run test` passes, with 240 files and 10,705 tests.

### Checked in Excel

In Excel for Mac.

- **Writing.** A1 `=1/0`, B1 `="#DIV/0!"`, and `ISERROR` of each in A2 and B2, written by `main` and by this branch. Both files open without a repair prompt and show `#DIV/0!`, `#DIV/0!`, `TRUE` and `FALSE`. That is expected: every hucre workbook sets `fullCalcOnLoad="1"`, so Excel recalculates on open and never uses the cached result. The fix matters to readers that use the cache without recalculating.
- **Reading a file Excel wrote.** I entered A1 `=1/0` and B1 `="#DIV/0!"` in Excel and saved. Excel stored `<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>` and `<c r="B1" t="str"><f>"#DIV/0!"</f><v>#DIV/0!</v></c>`, the same shapes this PR writes. Reading that file, `main` gives the two cells the same result. This branch gives A1 `formulaResultType: "error"` and leaves B1 without it.

### Not verified

- **A reader that uses cached results without recalculating**, the case this fix is for. I did not try one.
- **LibreOffice and Windows Excel.** Not opened.
