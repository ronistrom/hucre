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
- Out of scope: `_xlpm.` on LET and LAMBDA parameter names. Adding it means knowing which names a `LET` binds, and a scan for `NAME(` can't tell that. PARITY.md says hucre writes those names as given.

Tests: `test/formula-future-prefix.test.ts` adds 16 tests. 13 of them fail against `main`. The other three pass on `main` because they check what must not change: future functions the spec stores bare, already-prefixed formulas, and names that aren't calls. The three corrected assertions in `formulas-advanced.test.ts` fail against `main` too, which makes 16 failures in total. `bun run test` passes on Node 24.22.0, with 241 files and 10,714 tests.

### Not verified

I didn't open the output in Excel or LibreOffice; neither is available here. The prefixed forms match the spec table and XlsxWriter's output. I haven't confirmed in a real application that the five formulas from the issue now compute instead of showing `#NAME?`, and I haven't checked that a prefixed defined name or table total row behaves the same as one Excel saved.
