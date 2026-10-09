#!/bin/sh
# Download public OOXML test corpora (mostly files saved by real Excel/Word) into $1 (default ./corpus).
# Partial, sparse clones: only the matching files are fetched.
set -e
# On a Mac without the developer tools, /usr/bin/git is only an installer prompt; prefer a real git.
GIT=${GIT:-$(command -v "$HOME/.local/bin/git" || command -v git)}
DEST=${1:-corpus}
mkdir -p "$DEST"
fetch() {  # name url patterns...
  name=$1 url=$2; shift 2
  if [ ! -d "$DEST/$name/.git" ]; then
    "$GIT" clone --quiet --depth 1 --filter=blob:none --no-checkout "$url" "$DEST/$name" || exit 1
    "$GIT" -C "$DEST/$name" sparse-checkout set --no-cone "$@" || exit 1
    "$GIT" -C "$DEST/$name" checkout --quiet || exit 1
  fi
  echo "$name: $(find "$DEST/$name" -type f \( -name '*.xls[xm]' -o -name '*.docx' \) | wc -l | tr -d ' ') files"
}
fetch xlcalculator https://github.com/bradbase/xlcalculator.git '*.xlsx' '*.xlsm'
fetch formulas     https://github.com/vinci1it2000/formulas.git  '*.xlsx' '*.xlsm'
fetch pycel        https://github.com/dgorissen/pycel.git         '*.xlsx' '*.xlsm'
fetch python-docx  https://github.com/python-openxml/python-docx.git '*.docx'
fetch poi          https://github.com/apache/poi.git '/test-data/spreadsheet/*.xlsx' '/test-data/spreadsheet/*.xlsm' '/test-data/document/*.docx'
# SheetJS/test_files was also used originally, but GitHub has disabled that repository.
