<!-- title: fix(xlsx): write newer functions with their _xlfn. prefix -->

Closes #597

### What was wrong

Functions added after Excel 2007 are stored in the file under a prefixed name, for example `_xlfn.XLOOKUP`, `_xlfn.IFS`, and `_xlfn._xlws.FILTER` / `_xlfn._xlws.SORT` for the two worksheet-only ones ([MS-XLSX] §2.2.2). Excel adds the prefix when it saves and hides it in the formula bar, so callers write the bare name. hucre wrote that bare name into `<f>`, and a file that says `XLOOKUP(` names no function a reader knows, so the cell computes `#NAME?`.

The reproduction from the issue:

```ts
const formulas = ['XLOOKUP("b",A1:A3,B1:B3)', 'IFS(A1>1,"big",TRUE,"small")', 'TEXTJOIN(",",TRUE,A1:A3)', "FILTER(A1:A3,B1:B3>1)", "SORT(A1:A3)"]
await writeXlsx({ sheets: [{ name: "Sheet1", rows: [["a", 1], ["b", 2], ["c", 3]], cells: new Map(formulas.map((formula, i) => [`${i},3`, { formula }])) }] })
```

```xml
<!-- before -->
<c r="D1"><f>XLOOKUP("b",A1:A3,B1:B3)</f></c>
<c r="D2"><f>IFS(A1&gt;1,"big",TRUE,"small")</f></c>
<c r="D3"><f>TEXTJOIN(",",TRUE,A1:A3)</f></c>
<c r="D4"><f>FILTER(A1:A3,B1:B3&gt;1)</f></c>
<c r="D5"><f>SORT(A1:A3)</f></c>
<!-- after -->
<c r="D1"><f>_xlfn.XLOOKUP("b",A1:A3,B1:B3)</f></c>
<c r="D2"><f>_xlfn.IFS(A1&gt;1,"big",TRUE,"small")</f></c>
<c r="D3"><f>_xlfn.TEXTJOIN(",",TRUE,A1:A3)</f></c>
<c r="D4"><f>_xlfn._xlws.FILTER(A1:A3,B1:B3&gt;1)</f></c>
<c r="D5"><f>_xlfn._xlws.SORT(A1:A3)</f></c>
```

### Why the tests didn't catch it

The tests compared the writer's output with what they had passed in, so they checked that the formula text survived. None of them checked that the text was in the form Excel stores. `formulas-advanced.test.ts` went further and pinned the bug: it wrote `SORT(A2:A10)` and `UNIQUE(A1:A5)` and asserted the bare names in the XML and on read-back. Those three assertions now expect the prefixed form.

### What landed

- A new `src/xlsx/formula-prefix.ts` with `addFuturePrefixes(formula)` and a map from function name to prefix. The names come from the [MS-XLSX] Functions table (linked in the file). FILTER and SORT map to `_xlfn._xlws.`, and everything else maps to `_xlfn.`. Four names are added that the spec does not list yet but Excel prefixes: XMATCH, ARRAYTOTEXT, VALUETOTEXT and IMAGE. XlsxWriter's future-function table has all four. Every name is in one of those two lists.
  - ECMA.CEILING, ISO.CEILING, NETWORKDAYS.INTL and WORKDAY.INTL are left bare, because the spec lists them without a prefix.
  - `_xlfn._xlws.PY`, `_xlfn.SINGLE` and `_xlfn.ANCHORARRAY` are left out, because they aren't names a caller types.
- What the scanner touches:
  - It prefixes a name only when the name is a call: directly followed by `(`, and not after `!`.
  - It skips `"…"` string literals (with `""` escapes), `'…'` quoted sheet names (with `''` escapes) and `[…]` structured references or external-book indexes.
  - Dots count as name characters, so `_xlfn.XLOOKUP` and `_xlfn._xlws.SORT` scan as one unknown name and are left alone. That makes the pass idempotent: formulas read from Excel's files, which already carry the prefix, write back unchanged.
  - Matching ignores case and keeps the caller's spelling (`xlookup(` becomes `_xlfn.xlookup(`).
  - `SORTBY` gets `_xlfn.SORTBY`, not SORT's `_xlws.`.
  - A formula with no `(` returns at once, because this runs on the per-cell write path.
- Where it runs: once in `serializeCell`, before the `<f>` branches. That covers normal, shared and array formulas, and both streaming writers, because they serialize through the same function. It also runs on the other places hucre stores a formula in the same grammar:
  - conditional-formatting `<formula>`
  - data-validation `formula1` and `formula2`
  - table `totalsRowFormula`
  - defined names
- Defined names and table totals. The plan was to leave defined names alone. I prefix them for these reasons:
  - `<definedName>` holds a formula, not only a range, and a named `LAMBDA` or `LET` is exactly where these functions show up.
  - Excel writes the prefix there too.
  - The usual contents (`Sheet1!$A$1:$B$5`, print areas, quoted sheet names) have no call in them, so they pass through unchanged.
  - The table writer has no calculated-column formula, only `totalsRowFormula`, which gets the same treatment for the same reason.
- Read-back changes. The reader keeps formula text as the file has it, so a formula written as `XLOOKUP(…)` now reads back as `_xlfn.XLOOKUP(…)`. It's the same formula, but a strict equality check on a round trip will see the difference. Writing the read-back text again does not prefix it twice. `docs/PARITY.md`, the `Cell.formula` doc and the `NamedRange.range` doc now say so.
- Out of scope: `_xlpm.` on LET and LAMBDA parameter names. Adding it means knowing which names a `LET` binds, and a scan for `NAME(` can't tell that. PARITY.md says hucre writes those names as given. In Excel this matters more than it sounds: a LAMBDA or LET without `_xlpm.` is still removed on open (see "Checked in Excel").

Tests: `test/formula-future-prefix.test.ts` adds 16 tests. 13 of them fail against `main`. The other three pass on `main` because they check what must not change: future functions the spec stores bare, already-prefixed formulas, and names that aren't calls. The three corrected assertions in `formulas-advanced.test.ts` fail against `main` too, which makes 16 failures in total. `bun run test` passes on Node 24.22.0, with 241 files and 10,714 tests.

### Checked in Excel

In Excel for Mac.

- **The issue's five formulas.** The file `main` writes is worse than `#NAME?`. Excel asks to repair it, and its repair log says it removed formulas from `sheet1.xml`. After the repair, D1–D3 show `#NAME?`, and D4 and D5 (FILTER, SORT) are empty. This branch's file opens without a prompt and shows D1 `2`, D2 `small`, D3 `a,b,c`, D4 `b` and D5 `a`. FILTER and SORT show only their first value because no dynamic-array metadata is written; that is separate from this PR.
- **Defined names.** Names holding `XLOOKUP(…)` and `TEXTJOIN(…)`, written by this branch, work: cells `=Lookup` and `=Joined` show `2` and `a,b,c`.
- **LAMBDA and LET still fail, because of the missing `_xlpm.`.** I wrote a defined name `Double` as `LAMBDA(x,x*2)` and a cell `LET(x,2,x*3)`. This branch writes `_xlfn.LAMBDA(x,x*2)` and `_xlfn.LET(x,2,x*3)`. Excel asks to repair the file, removes the named range and the formula, and `=Double(B2)` shows `#NAME?`. The same file with `_xlpm.x` hand-edited into both opens cleanly and computes `2` and `6`. So the prefix this PR adds is necessary but not enough for a named LAMBDA or a LET: those need `_xlpm.` as well. That is a follow-up, not something this PR makes worse, because `main`'s output fails the same way.
- **Table total rows.** These could not be checked through hucre. Any custom `totalsRowFormula` hucre writes makes Excel ask for repair, on `main` as well, even `SUM(Table1[Qty])*2` with no newer function in it. A `sum` total opens cleanly but shows an empty total row, because hucre does not write the total row's cells. That is an existing problem with table totals, separate from this PR.

### Not verified

- **Table total rows**, as above.
- **LibreOffice and Windows Excel.** Not opened.
