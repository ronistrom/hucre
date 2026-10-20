// Writes the workbook for one issue using the hucre source tree given by
// --src, so the same input can be written by `main` (before) and by the fix
// branch (after) and the two files compared side by side in Excel.
//
//   bun excel-check/make.ts --src <dir>/src --issue 597 --out out/597-after.xlsx

import { parseArgs } from "node:util"
import { writeFile } from "node:fs/promises"

const { values } = parseArgs({
  options: { src: { type: "string" }, issue: { type: "string" }, out: { type: "string" } },
})
if (!values.src || !values.issue || !values.out) {
  console.error("usage: make.ts --src <hucre>/src --issue <n> --out <file.xlsx>")
  process.exit(2)
}

const { writeXlsx } = await import(`${values.src}/xlsx/writer.ts`)

/** A solid-colour PNG of the given size, so #598 needs no image file. */
async function solidPng(width: number, height: number, rgb: [number, number, number]) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    return c >>> 0
  })
  const crc32 = (bytes: Uint8Array) => {
    let c = 0xffffffff
    for (const b of bytes) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8)
    return (c ^ 0xffffffff) >>> 0
  }
  const chunk = (type: string, data: Uint8Array) => {
    const body = new Uint8Array(4 + data.length)
    body.set(new TextEncoder().encode(type), 0)
    body.set(data, 4)
    const out = new Uint8Array(12 + data.length)
    const view = new DataView(out.buffer)
    view.setUint32(0, data.length)
    out.set(body, 4)
    view.setUint32(8 + data.length, crc32(body))
    return out
  }
  const ihdr = new Uint8Array(13)
  const hv = new DataView(ihdr.buffer)
  hv.setUint32(0, width)
  hv.setUint32(4, height)
  ihdr.set([8, 2, 0, 0, 0], 8) // 8-bit RGB
  const raw = new Uint8Array(height * (1 + width * 3))
  for (let y = 0; y < height; y++) {
    const row = y * (1 + width * 3)
    for (let x = 0; x < width; x++) raw.set(rgb, row + 1 + x * 3)
  }
  const idat = new Uint8Array(
    await new Response(
      new Blob([raw]).stream().pipeThrough(new CompressionStream("deflate")),
    ).arrayBuffer(),
  )
  const parts = [
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", ihdr),
    chunk("IDAT", idat),
    chunk("IEND", new Uint8Array()),
  ]
  const png = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    png.set(p, at)
    at += p.length
  }
  return png
}

const workbooks: Record<string, () => Promise<unknown>> = {
  // A partial cell font should take name and size from defaultFont.
  "594": async () => ({
    defaultFont: { name: "Arial", size: 10 },
    sheets: [
      {
        name: "Sheet1",
        rows: [
          [{ value: "Bold header", style: { font: { bold: true, color: { rgb: "FF1F3A5F" } } } }],
          ["Plain text in the default font"],
        ],
      },
    ],
  }),

  // 6-digit colour scale and data bar colours.
  "595": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: Array.from({ length: 10 }, (_, i) => [i + 1, (i + 1) * 10]),
        conditionalRules: [
          {
            type: "colorScale",
            priority: 1,
            range: "A1:A10",
            colorScale: { cfvo: [{ type: "min" }, { type: "max" }], colors: ["F8696B", "63BE7B"] },
          },
          {
            type: "dataBar",
            priority: 2,
            range: "B1:B10",
            dataBar: { cfvo: [{ type: "min" }, { type: "max" }], color: "638EC6" },
          },
        ],
      },
    ],
  }),

  // fitToWidth/fitToHeight without fitToPage: true.
  "596": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: Array.from({ length: 60 }, (_, r) =>
          Array.from({ length: 20 }, (_, c) => `row ${r + 1} col ${c + 1} — wide text`),
        ),
        pageSetup: { fitToWidth: 1, fitToHeight: 0 },
      },
    ],
  }),

  // Newer functions without the _xlfn. prefix (the issue's reproduction).
  "597": async () => {
    const formulas = [
      'XLOOKUP("b",A1:A3,B1:B3)',
      'IFS(B1>1,"big",TRUE,"small")',
      'TEXTJOIN(",",TRUE,A1:A3)',
      "FILTER(A1:A3,B1:B3>1)",
      "SORT(A1:A3)",
    ]
    return {
      sheets: [
        {
          name: "Sheet1",
          rows: [
            ["a", 1],
            ["b", 2],
            ["c", 3],
          ],
          cells: new Map(formulas.map((formula, i) => [`${i},3`, { formula }])),
        },
      ],
    }
  },

  // An image with `from` and a size but no `to`.
  "598": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: [["The blue bar should be 200 × 40 px, starting at B2"]],
        images: [
          {
            data: await solidPng(200, 40, [0x1f, 0x3a, 0x5f]),
            type: "png",
            anchor: { from: { row: 1, col: 1 } },
            width: 200,
            height: 40,
          },
        ],
      },
    ],
  }),

  // A formula whose cached result is an error, next to one returning that text.
  "599": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: [],
        cells: new Map<string, object>([
          ["0,0", { formula: "1/0", formulaResult: "#DIV/0!", formulaResultType: "error" }],
          ["0,1", { formula: '"#DIV/0!"', formulaResult: "#DIV/0!" }],
          ["1,0", { formula: "ISERROR(A1)", formulaResult: true }],
          ["1,1", { formula: "ISERROR(B1)", formulaResult: false }],
        ]),
      },
    ],
  }),

  // A row style on RowDef.
  "601": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: [["row 1"], ["row 2 has a row style"], ["row 3"]],
        rowDefs: new Map([
          [
            1,
            {
              height: 24,
              style: {
                font: { bold: true },
                fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
              },
            },
          ],
        ]),
      },
    ],
  }),

  // Rich text in the default shared-strings mode.
  "602": async () => ({
    sheets: [
      {
        name: "Sheet1",
        rows: [],
        cells: new Map<string, object>([
          [
            "0,0",
            {
              value: "Bold red then normal",
              richText: [
                { text: "Bold red", font: { bold: true, color: { rgb: "FFC00000" } } },
                { text: " then normal" },
              ],
            },
          ],
          ["1,0", { value: "plain text" }],
        ]),
      },
    ],
  }),
}

const make = workbooks[values.issue]
if (!make) {
  console.error(`no workbook for issue ${values.issue}`)
  process.exit(2)
}
await writeFile(values.out, await writeXlsx(await make()))
