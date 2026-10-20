// Reads a workbook saved by Excel with the hucre source tree given by --src
// and prints what the read-side fixes are about, so `main` and a fix branch
// can be compared on a real Excel file.
//
//   bun excel-check/read.ts --src <dir>/src --file my-excel-file.xlsx

import { parseArgs } from "node:util"
import { readFile } from "node:fs/promises"

const { values } = parseArgs({ options: { src: { type: "string" }, file: { type: "string" } } })
if (!values.src || !values.file) {
  console.error("usage: read.ts --src <hucre>/src --file <file.xlsx>")
  process.exit(2)
}

const { readXlsx } = await import(`${values.src}/xlsx/reader.ts`)
const bytes = new Uint8Array(await readFile(values.file))

const show = (v: unknown) =>
  JSON.stringify(v, (_k, x) => (x instanceof Date ? `Date(${x.toISOString()})` : x))

for (const [label, options] of [
  ["default read", {}],
  ['dates: "serial" (#600)', { dates: "serial" }],
] as const) {
  const wb = await readXlsx(bytes, options)
  console.log(`  ── ${label}`)
  for (const sheet of wb.sheets) {
    console.log(`  sheet "${sheet.name}"`)
    sheet.rows.forEach((row: unknown[], r: number) => {
      row.forEach((value, c) => {
        if (value === null || value === undefined || value === "") return
        const cell = sheet.cells?.get(`${r},${c}`)
        const ref = `${String.fromCharCode(65 + (c % 26))}${r + 1}`
        const detail = cell
          ? {
              type: cell.type,
              formula: cell.formula,
              formulaResult: cell.formulaResult,
              formulaResultType: cell.formulaResultType,
              richText: cell.richText ? `${cell.richText.length} runs` : undefined,
            }
          : {}
        console.log(`    ${ref} = ${show(value)} ${show(detail)}`)
      })
    })
    if (label !== "default read") continue
    for (const [r, def] of sheet.rowDefs ?? []) {
      console.log(`    row ${r + 1} def: ${show(def)}`)
    }
    for (const img of sheet.images ?? []) {
      console.log(
        `    image: ${show({ anchor: img.anchor, width: img.width, height: img.height })}`,
      )
    }
  }
}
