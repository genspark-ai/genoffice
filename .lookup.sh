#!/bin/bash
# usage: .lookup.sh <file-glob-dir> <key1> [key2...]
# prints key<TAB>ja<TAB>ko<TAB>fr<TAB>de<TAB>es
DIR="$1"; shift
LANGS="ja ko fr de es"
printf "%s" "key"
for l in $LANGS; do printf "\t%s" "$l"; done
printf "\n"
for k in "$@"; do
  printf "%s" "$k"
  for l in $LANGS; do
    v=$(grep -h "^  '\?$k'\?: " "$DIR/$l.ts" 2>/dev/null | head -1 | sed -E "s/^  '?$k'?: //")
    printf "\t%s" "$v"
  done
  printf "\n"
done
