# Excel check for the #594–#602 fix branches

This branch is not for upstream. It builds the same workbook with `main` and
with each fix branch, so the two can be opened side by side in real Excel
before the PRs go to productdevbook/hucre.

Needs [Bun](https://bun.sh) (`curl -fsSL https://bun.sh/install | bash`).
No `bun install` is needed: hucre has no runtime dependencies, and the
script imports each branch's `src/` directly.

```bash
git clone https://github.com/ronistrom/hucre.git && cd hucre
git checkout excel-check
excel-check/run.sh                      # → excel-check/out/*-before.xlsx, *-after.xlsx, CHECKLIST.md
excel-check/run.sh read ~/Desktop/from-excel.xlsx   # read an Excel-saved file with main and each read-side branch
```

Branches that are not on `origin` yet are skipped. Run it again after a
fix branch is updated; it fetches first.

`CHECKLIST.md` says what to look at in each pair, and how to make the
Excel-saved file for the read checks.
