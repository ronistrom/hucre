<!-- title: fix(xlsx): write newer functions with their _xlfn. prefix, and LET/LAMBDA names with _xlpm. -->

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
- **`_xlpm.` on the names LET and LAMBDA bind.** The issue suggested leaving this out. Checking in Excel showed it can't be left out: with only `_xlfn.`, Excel asks to repair the file and removes the LET formula and the named LAMBDA (see "Checked in Excel"). A bound name is stored with the prefix both where it is bound and everywhere it is used: `_xlfn.LET(_xlpm.x,2,_xlpm.x*3)`. Which names to prefix depends on scope, which a scan for `NAME(` can't see. So a formula that mentions LET or LAMBDA goes through a parse of its argument lists, and every other formula keeps the plain scan. Each rule matches what Excel 16.114 for Mac stored when the bare formula was typed in:
  - A LET name is in scope from the next value onwards, not in its own value.
  - LAMBDA parameters are in scope in the body only, so the arguments of a LAMBDA invoked on the spot, `LAMBDA(x,x*2)(x)`, stay bare.
  - An optional `[b]` parameter is bound as `_xlop.b` and used as `_xlpm.b`.
  - A LET-bound LAMBDA called by name is written `_xlpm.f(2)`.
  - A use takes the spelling of its binding: `LET(Ab,2,aB)` is stored with `_xlpm.Ab` twice.
  - Inside a scope that binds `a`, `A:A` becomes `_xlpm.a:_xlpm.a`, but `$A$1`, `A$1` and `$A1` stay cell references.
  - Strings, quoted sheet names, sheet-qualified names, table references and already-prefixed text are left alone, so Excel's own stored form writes back unchanged.

  PARITY.md and the `Cell.formula` / `NamedRange.range` docs now describe this instead of saying the names are written as given.

Tests: `test/formula-future-prefix.test.ts` adds 23 tests.

- **`_xlfn.` (16 tests).** 13 fail against `main`. The other three pass on `main` because they check what must not change: future functions the spec stores bare, already-prefixed formulas, and names that aren't calls. The three corrected assertions in `formulas-advanced.test.ts` fail against `main` too.
- **`_xlpm.` (7 tests).** Six fail without the `_xlpm.` commit.

`bun run test` passes, with 241 files and 10,721 tests.

After review, a third commit adds ten functions that Excel 16.114 for Mac saved with `_xlfn.` when typed in bare, though neither [MS-XLSX] nor XlsxWriter lists them yet: DETECTLANGUAGE, GROUPBY, PERCENTOF, PIVOTBY, REGEXEXTRACT, REGEXREPLACE, REGEXTEST, STOCKHISTORY, TRANSLATE and TRIMRANGE. A file written with five of them opens cleanly and computes.

- **ENCODEURL is left out.** Neither source lists it, and Excel for Mac, which doesn't have the function, saves it bare.
- **One prefix is still missing.** A function passed as an argument, as in `GROUPBY(A2:A3,B2:B3,SUM)`, is stored as `_xleta.SUM`. Written bare, Excel computes `#NAME?`; with `_xleta.SUM` hand-edited in, it computes. Adding it needs Excel's whole function list, to tell such a name apart from a defined name. PARITY.md says so and points to a LAMBDA instead (`LAMBDA(x,SUM(x))`), which is written in its stored form.

The new test fails without that commit. `bun run test` passes, with 241 files and 10,723 tests.

### Checked in Excel

In Excel for Mac.

- **The issue's five formulas.** The file `main` writes is worse than `#NAME?`. Excel asks to repair it, and its repair log says it removed formulas from `sheet1.xml`. After the repair, D1–D3 show `#NAME?`, and D4 and D5 (FILTER, SORT) are empty. This branch's file opens without a prompt and shows D1 `2`, D2 `small`, D3 `a,b,c`, D4 `b` and D5 `a`. FILTER and SORT show only their first value because no dynamic-array metadata is written; that is separate from this PR.
- **Defined names.** Names holding `XLOOKUP(…)` and `TEXTJOIN(…)`, written by this branch, work: cells `=Lookup` and `=Joined` show `2` and `a,b,c`.
- **LAMBDA and LET needed `_xlpm.`, which is why it is in this PR.** I wrote a defined name `Double` as `LAMBDA(x,x*2)` and a cell `LET(x,2,x*3)`. With only `_xlfn.` (`_xlfn.LAMBDA(x,x*2)`, `_xlfn.LET(x,2,x*3)`), Excel asks to repair the file, removes the named range and the formula, and `=Double(B2)` shows `#NAME?`. `main`'s output fails the same way. The same file with `_xlpm.x` hand-edited in opens cleanly and computes `2` and `6`. With the `_xlpm.` commit, a file the branch writes also opens cleanly in Excel 16.114 for Mac, with every formula and defined name kept and the expected values. That file has a named LAMBDA, an optional parameter and a LET-bound call.
- **Table total rows.** These could not be checked through this branch alone. Any custom `totalsRowFormula` hucre writes makes Excel ask for repair, on `main` as well, even `SUM(Table1[Qty])*2` with no newer function in it, because hucre does not write the total row's cells. That is #604, a separate bug. With #604's branch merged in, a custom `XLOOKUP` total row opens cleanly and shows `3`: the total-row cell gets `_xlfn.` from this PR's cell-formula pass, and `<totalsRowFormula>` from this PR's table-writer change.

### Not verified

- **Table total rows on this branch alone**, as above: they need #604.
- **LibreOffice and Windows Excel.** Not opened.
