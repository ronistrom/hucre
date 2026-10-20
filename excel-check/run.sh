#!/usr/bin/env bash
# Builds before/after workbooks for each fix branch, or reads an Excel-saved
# file with main and with each read-side fix branch. See README.md here.
#
#   excel-check/run.sh                 # write the before/after files to excel-check/out/
#   excel-check/run.sh read file.xlsx  # read an Excel-saved file with each branch
set -euo pipefail

repo=$(git rev-parse --show-toplevel)
here="$repo/excel-check"
remote=${REMOTE:-origin}
cache=$(mktemp -d)
trap 'rm -rf "$cache"' EXIT

# issue  branch                         slug
branches=(
  "594 fix/font-inherits-default        font-inherits-default"
  "595 fix/cf-color-argb                cf-color-argb"
  "596 fix/fit-to-page-implied          fit-to-page"
  "597 fix/xlfn-prefix                  xlfn-prefix"
  "598 fix/image-one-cell-anchor        image-size"
  "599 fix/formula-error-result         formula-error"
  "600 feat/read-dates-serial           dates-serial"
  "601 feat/row-style                   row-style"
  "602 fix/rich-text-shared-strings     rich-text-shared"
)

git fetch -q "$remote"

# Extract src/ of a commit once; print the directory.
src_of() {
  local rev
  rev=$(git rev-parse --short "$1")
  if [ ! -d "$cache/$rev" ]; then
    mkdir -p "$cache/$rev"
    git archive "$rev" src | tar -x -C "$cache/$rev"
  fi
  echo "$cache/$rev/src"
}

has_branch() { git rev-parse -q --verify "$remote/$1" >/dev/null; }

if [ "${1:-}" = "read" ]; then
  file=$(cd "$(dirname "$2")" && pwd)/$(basename "$2")
  base=$(src_of "$remote/main")
  for entry in "${branches[@]}"; do
    read -r issue branch _ <<<"$entry"
    case $issue in 599 | 600 | 601) ;; *) continue ;; esac
    has_branch "$branch" || { echo "#$issue: $remote/$branch not pushed yet, skipped"; continue; }
    echo "════ #$issue — before ($remote/main)"
    bun "$here/read.ts" --src "$base" --file "$file"
    echo "════ #$issue — after ($remote/$branch)"
    bun "$here/read.ts" --src "$(src_of "$remote/$branch")" --file "$file"
  done
  exit 0
fi

out="$here/out"
mkdir -p "$out"
for entry in "${branches[@]}"; do
  read -r issue branch slug <<<"$entry"
  [ "$issue" = 600 ] && continue # read-side only: use `run.sh read`
  has_branch "$branch" || { echo "#$issue: $remote/$branch not pushed yet, skipped"; continue; }
  base=$(git merge-base "$remote/main" "$remote/$branch")
  bun "$here/make.ts" --src "$(src_of "$base")" --issue "$issue" --out "$out/$issue-$slug-before.xlsx"
  bun "$here/make.ts" --src "$(src_of "$remote/$branch")" --issue "$issue" --out "$out/$issue-$slug-after.xlsx"
  echo "#$issue: $out/$issue-$slug-{before,after}.xlsx"
done
cp "$here/CHECKLIST.md" "$out/"
echo "Open $out/CHECKLIST.md and the files next to it in Excel."
