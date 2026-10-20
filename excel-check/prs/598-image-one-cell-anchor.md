<!-- title: fix(xlsx): write an image without anchor.to as a oneCellAnchor -->

Closes #598

### What was wrong

`writeDrawing` always wrote images as `xdr:twoCellAnchor`. When `anchor.to` was absent, it made up the bottom-right corner as `from` plus 3 columns and 5 rows. `width`/`height` only reached the picture's `<a:ext>`, which Excel and LibreOffice ignore under a two-cell anchor. So the image was stretched over a 3 × 5 cell box whatever size it was given.

This is the issue's reproduction, a 200 × 40 px logo at B2 with no `to`:

```ts
images: [{ data: png, type: "png", anchor: { from: { row: 1, col: 1 } }, width: 200, height: 40 }]
```

```xml
<!-- before -->
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:pic>…</xdr:pic><xdr:clientData/>
</xdr:twoCellAnchor>
<!-- after -->
<xdr:oneCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:ext cx="1905000" cy="381000"/>
  <xdr:pic>…</xdr:pic><xdr:clientData/>
</xdr:oneCellAnchor>
```

Reading the file back with `readXlsx`:

```
before: {"anchor":{"from":{"row":1,"col":1},"to":{"row":6,"col":4}},"width":200,"height":40}
after:  {"anchor":{"from":{"row":1,"col":1}},"width":200,"height":40}
```

### Why the tests didn't catch it

`test/images.test.ts` had a test called "uses default to anchor when not specified" that asserted the invented `to` (col 4, row 7 for an image at col 1, row 2). The test pinned the bug. The `oneCellAnchor` tests in `test/one-cell-anchor.test.ts` only read hand-built files, so nothing checked what the writer produces for an image without `to`. The write → read parity register's image probe always passed a `to`.

### What landed

- `src/xlsx/drawing-writer.ts`: when `anchor.to` is absent, the image is written as `xdr:oneCellAnchor`. It holds `xdr:from`, then `<xdr:ext cx cy>` from `width`/`height` × 9525 EMU (or the existing default of 3000000 × 2000000 EMU), then the same `xdr:pic` and `<xdr:clientData/>`. Both anchor kinds share one `pic`. When `to` is given, the `twoCellAnchor` output is byte-identical to before, and a test pins the full string.
- The reader is unchanged. `parseOneCellAnchor` already returns `{ from }` plus the size from `xdr:ext`, so the image reads back as written. The comment in `reader.ts` that said hucre always writes `twoCellAnchor` (#407) is updated.
- Doc comments on `SheetImage.anchor` and `width`/`height` now say which anchor each case is written as and where the size goes.
- **Rendering change.** An image written without `anchor.to` now shows at its `width`/`height`, or at the default size of about 315 × 210 px. It no longer fills 3 columns × 5 rows, and it no longer resizes with those cells. Callers who want the old stretched layout can pass `anchor.to`. This is in the commit body, which is where release notes come from.
- Tests:
  - The default-`to` test in `images.test.ts` now expects a `oneCellAnchor` with the default `ext`.
  - Seven new tests in `one-cell-anchor.test.ts` cover:
    - the issue's reproduction, including child order `from, ext, pic, clientData`
    - read-back with no `to`
    - the default size when `width`/`height` are absent
    - `altText`/`title` surviving in the one-cell form
    - the byte-identical `twoCellAnchor` when `to` is given
    - both anchor kinds on one sheet, in declaration order
    - `openXlsx` → `saveXlsx`
  - The parity register's image probe gains an image without `to`.

Eight tests fail against `main`: the corrected `images.test.ts` test, the parity register, and six of the seven new tests. The seventh is the byte-identical `twoCellAnchor` guard, which already passed. `bun run test` passes on Node 24 (lint, typecheck, and 240 files with 10,705 tests).

Not in this PR: text boxes and charts still invent a `to` when it is absent (+3/+3 and +8/+15). Text boxes need reader work first, because the reader takes text boxes only from `twoCellAnchor`. Charts have no size field to put in an `ext`. Both are left for a follow-up.

README and `docs/PARITY.md` never described the 3 × 5 default, so they don't need changes.

### Not verified

I did not open the output in Excel or LibreOffice. The claim that both apps size a `oneCellAnchor` picture from `xdr:ext` and ignore `a:ext` under a `twoCellAnchor` comes from the issue and the schema, not from checking it myself.
