#!/usr/bin/env bash
# Ask Word how it paginates a .docx (thoughts ADR 029): open it in Microsoft
# Word, save it as PDF, and report per page its size and which numbered
# paragraphs it carries. Word is the oracle for DOCX standalone open the way
# InDesign is for IDML.
#
#   bash scripts/word-pagination-probe.sh <in.docx> <out.pdf>
#
# Office automation traps (see the corpus harness): Word can only be relied
# on to read and write inside its OWN sandbox container, so the file is
# staged there; quit gracefully (a SIGKILLed Word answers but cannot open);
# judge success by the PDF, never by AppleScript's return string.
set -euo pipefail
IN="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
OUT="$(cd "$(dirname "$2")" && pwd)/$(basename "$2")"
STAGE="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/paged-probe"
mkdir -p "$STAGE"
NAME="$(basename "$IN")"
cp "$IN" "$STAGE/$NAME"
PDF="$STAGE/${NAME%.docx}.pdf"
rm -f "$PDF"
osascript <<OSA || true
with timeout of 180 seconds
    tell application "Microsoft Word"
        activate
        delay 3
        open POSIX file "$STAGE/$NAME"
        delay 2
        save as active document file name "$PDF" file format format PDF
        close active document saving no
    end tell
end timeout
OSA
for i in $(seq 1 30); do [ -s "$PDF" ] && break; sleep 1; done
[ -s "$PDF" ] || { echo "Word produced no PDF (judge by the artifact)"; exit 1; }
cp "$PDF" "$OUT"
echo "==> $OUT"
