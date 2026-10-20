<!-- title: feat: a dates: "serial" read option that keeps date serials as numbers -->

Closes #600

### What was missing

A numeric cell under a date format always reads as a `Date`. That is the right default, but a `Date` can't always be turned back into what the cell held. Six serials, each written with `numFmt: "yyyy-mm-dd hh:mm:ss"` and read back:

```
readXlsx(bytes)                       # main, and still the default
  45306            ->  Date 2024-01-15T00:00:00.000Z -> dateToSerial 45306
  45306.5          ->  Date 2024-01-15T12:00:00.000Z -> dateToSerial 45306.5
  45306.123456789  ->  Date 2024-01-15T02:57:46.667Z -> dateToSerial 45306.123456794
  59               ->  Date 1900-02-28T00:00:00.000Z -> dateToSerial 59
  60               ->  Date 1900-02-28T00:00:00.000Z -> dateToSerial 59
  61               ->  Date 1900-03-01T00:00:00.000Z -> dateToSerial 61
```

On `main` there is no way to get the number. `{ dates: "serial" }` is ignored by both `readXlsx` and `streamXlsxRows`, which return exactly the same output as above. With this PR:

```
readXlsx(bytes, { dates: "serial" })        streamXlsxRows(bytes, { dates: "serial" })
  45306            ->  number 45306           number 45306
  45306.5          ->  number 45306.5         number 45306.5
  45306.123456789  ->  number 45306.123456789 number 45306.123456789
  59               ->  number 59              number 59
  60               ->  number 60              number 60
  61               ->  number 61              number 61
```

### What landed

- `ReadOptions.dates?: "date" | "serial"`, default `"date"`. The doc comment covers the three reasons from the issue: millisecond rounding, serial 60 and time-only formats.
- In `"serial"` mode:
  - A date-formatted number is returned as the number the file holds, with `type: "number"`. With `readStyles`, its `numFmt` is still on the style, so `isDateFormat(cell.style.numFmt)` tells you which numbers were dates. I picked `type: "number"` over the issue's other option (`type: "date"` holding a number) so that `type` always matches `typeof value`.
  - A formula's cached date result stays a number.
  - A `t="d"` ISO cell is converted with `dateToSerial(date, is1904)`, using the workbook's date system or the `dateSystem` override. That way every date in a workbook reads the same way, whichever form the producer used. A `t="d"` bare time (`13:45:30`) still reads as text, the same as in the default mode.
  - A stored serial is returned unconverted, so in a 1904 workbook it counts from 1904.
- Readers that honour it:
  - `readXlsx` (`WorksheetContext.dates`, used in the `n` and `d` branches of `processCell`), so `openXlsx`, `readXlsxObjects` and `read()` also honour it
  - `streamXlsxRows`, on both its `Uint8Array` and `ReadableStream` paths
  - `readXls`
  - `readXlsb`
- ODS: the option doesn't apply. ODS stores dates as ISO text (`office:date-value`), so there is no serial to return. `dateSystem` is marked n/a for ODS for the same reason. The `docs/PARITY.md` "Read options, per reader" table, the JSDoc table on `ReadOptions` and the option's doc comment all say so, and a test checks that `readOds` still returns a `Date`.
- `openXlsx` → `saveXlsx`: `saveXlsx` writes a serial back as a plain number, so its date format survives only with `readStyles: true`. The `openXlsx` doc comment and `docs/PARITY.md` say so, and a test covers the round trip, serial 60 included.
- `docs/PARITY.md`:
  - The "Serial 60 collapses onto 59" section now names the option as the escape hatch.
  - The per-reader table has a `dates` row.
- README: the option is in the `readXlsx` options example, and there is a short example after the serial-60 note.
- `test/read-dates-serial.test.ts` reads the Excel-written `.xls`/`.xlsb`/`.xlsx` fixtures off disk, so it is listed in `tsconfig.cli.json` (and excluded from `tsconfig.json`), as CONTRIBUTING asks.
- `test/xlsx-stream-parity.test.ts`: the buffered-vs-streaming sweep over the fixture corpus now also runs with `dates: "serial"`. Each reader has its own date branch, and #496 was a fix that reached only one of them.

Tests: 18 of the 25 new tests in `test/read-dates-serial.test.ts` fail against `main`. The 7 that pass already held on `main`:
- five check that the default, and an explicit `dates: "date"`, still give Dates in `readXlsx`, `streamXlsxRows`, `readXls` and `readXlsb`
- one checks that a `t="d"` bare time stays text
- one checks that ODS still returns a `Date` The new serial-mode case in the corpus sweep passes on `main` too, because there both readers ignore the option. It's there to catch drift later.

`bun run test` passes on Node 24.22.0 (241 files, 10,724 tests), and so does `TZ=Asia/Tokyo bun run test` (same counts).

### Checked with a file Excel wrote

This change only affects reading, so there is no output to open. Instead I saved a workbook from Excel for Mac and read it with `main` and with this branch. It had A3 `15.1.2024 12:00` formatted as a date and time, B3 `6:30` formatted `hh:mm`, and C3 the number `60` formatted as a date.

- **Default read:** unchanged on both. A3, B3 and C3 come back as `Date`s, with C3 as `1900-02-28`.
- **`dates: "serial"`:**
  - `main` ignores the option.
  - This branch returns A3 `45306.5`, B3 `0.2708333333333333` and C3 exactly `60`, the serials Excel stored.

The `.xls`/`.xlsb`/`.xlsx` expectations in the tests come from the Excel-written fixtures in `test/fixtures` (#464), with the serials counted by hand from `scripts/fixtures/make-fixtures.vbs`.

### Not verified

- Only the `.xlsx` reader was checked against a file saved in this session. `.xls` and `.xlsb` rely on the existing fixtures.
- `readXls` and `readXlsb` don't set `Workbook.dateSystem`. In serial mode, a caller reading a 1904 `.xls`/`.xlsb` therefore can't tell from the result which epoch the numbers count from. `readXlsx` does report it. That gap predates this PR and is left for a follow-up.
- There is no CLI flag for the option.
