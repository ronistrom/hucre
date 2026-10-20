<!-- title: fix(xlsx): write colour scale and data bar colours as ARGB -->

Closes #595

### What was wrong

Every other colour field turns a 6-digit RGB into ARGB with an `FF` alpha prefix: fonts, fills, borders, rich text, tab colour and sparklines. Colour scale and data bar colours in `conditionalRules` were written exactly as given:

```ts
conditionalRules: [
  { type: "colorScale", priority: 1, range: "A1:A3", colorScale: { cfvo: [{ type: "min" }, { type: "max" }], colors: ["638EC6", "F8696B"] } },
  { type: "dataBar", priority: 2, range: "A1:A3", dataBar: { cfvo: [{ type: "min" }, { type: "max" }], color: "638EC6" } },
]
```

```xml
<!-- before -->
<colorScale>…<color rgb="638EC6"/><color rgb="F8696B"/></colorScale>
<dataBar>…<color rgb="638EC6"/></dataBar>
<!-- after -->
<colorScale>…<color rgb="FF638EC6"/><color rgb="FFF8696B"/></colorScale>
<dataBar>…<color rgb="FF638EC6"/></dataBar>
```

`CT_Color/@rgb` is `ST_UnsignedIntHex`, an `xsd:hexBinary` of length 4 (eight hex digits), so the 6-digit form is not schema-valid.

### Why the tests didn't catch it

`coverage-xlsx-writer.test.ts` had a data bar test that wrote `"638EC6"` and asserted `<color rgb="638EC6"/>`, so the test pinned the bug. The round-trip tests in `conditional-formatting.test.ts` all use 8-digit colours.

### What landed

- A small `toArgb(rgb)` helper in `worksheet-writer.ts`. `serializeColorAttrs` and the colour scale, data bar and sparkline sites all use it. The sparkline site had the same check written inline, and its output is unchanged.
- `serializeConditionalFormatting` is shared with `writeXlsxStream` / `writeXlsxStreamSheets`, so the streaming writers are fixed too. There is a test for that.
- Doc comments on `colorScale.colors` and `dataBar.color` say that a 6-digit value is accepted and reads back in 8-digit form. The reader keeps conditional-formatting colours exactly as written, so `"638EC6"` now reads back as `"FF638EC6"`. It is the same colour, but a strict equality check on a round trip will see the difference. Before this change, the value read back unchanged because the file held the invalid 6-digit form.

Tests: two of the three new tests fail against `main`, plus the corrected data bar assertion. The third new test checks that 8-digit colours pass through untouched, which already worked. `bun run test` passes on Node 24.22.0, with 240 files and 10,701 tests.

### Not verified

I did not open the output in Excel or LibreOffice. The change only makes the value schema-valid, and LibreOffice already accepted both forms according to the issue.
