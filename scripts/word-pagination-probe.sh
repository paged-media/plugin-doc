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
# The document is addressed BY NAME, never as "active document": a long
# document is still opening when the next command runs, and "active
# document" then saved the PREVIOUS file under this one's name (a 61-page
# real document came back as three different PDFs, all of it). Wait until
# the PDF stops growing; a long document takes minutes.
osascript <<OSA || true
with timeout of 900 seconds
    tell application "Microsoft Word"
        activate
        delay 3
        open POSIX file "$STAGE/$NAME"
        delay 2
        set d to document "$NAME"
        save as d file name "$PDF" file format format PDF
        close d saving no
    end tell
end timeout
OSA
prev=-1
for i in $(seq 1 120); do
  size=$(stat -f %z "$PDF" 2>/dev/null || echo 0)
  [ "$size" -gt 0 ] && [ "$size" = "$prev" ] && break
  prev=$size; sleep 2
done
pdfinfo "$PDF" >/dev/null 2>&1 || { echo "Word produced no PDF (judge by the artifact)"; exit 1; }
cp "$PDF" "$OUT"
echo "==> $OUT"
