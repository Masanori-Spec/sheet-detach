# XLSX handoff validation

The test layers make different claims. A formula string or an openpyxl read alone
is not evidence that a spreadsheet application recalculates the workbook.

## Independent oracle

Run `npm run fixture`, install `requirements.txt`, then run
`python3 tools/oracle.py`.

The Python oracle does not import application code. It opens the source and
export with openpyxl and inspects their ZIP/XML directly. A deliberately narrow
recursive-descent parser evaluates actual formula text using numeric literals,
A1 references, parentheses, unary signs, arithmetic, and SUM. Unsupported syntax
fails instead of being guessed. The parser has separate precedence and rejection
self-tests. It never reads cached results to calculate an answer.

For all 12 combinations of B2 in `[0, 1, 8, 10]` and B3 in `[0, 2, 3]`, both
the source and exported formulas must match independently written expectations:

- C2 = 6 × B2 − 2
- C3 = 5 × B3
- C4 = 6 × B2 + 5 × B3 − 2

The initial values are 46, 10, 56. Changing B2 to 10 and B3 to 3 must produce
58, 15, 73. Source formula caches must all be the deliberately wrong value 999.
The exported values must be absent or correct, never stale. Tests also require
all three Plan result cells to remain formulas with their original local
dependencies, a retained SUM, unchanged non-formula cells and styles, automatic
recalculation settings, one worksheet payload, valid manifest/relationships, and
no Constants/Derived sheet names or orphan worksheet parts in the handoff.
Eight intentionally damaged exports must be rejected: flattened formulas,
wrong arithmetic, stale caches, leaked removed-sheet names, orphan sheets,
dangling relationships, changed cell styles, and missing worksheet manifests.

Evidence: `test-results/oracle/evidence.json`. This evidence explicitly reports
that it is not native-application recalculation.

## Hosted LibreOffice Calc consumer

The hosted Ubuntu 22.04 workflow must install `libreoffice-calc` from the official
Ubuntu apt repositories, run the browser tests first, then run
`python3 tools/libreoffice.py`. The input is strictly the actual browser download
at `test-results/browser/actual-artifacts/handoff.xlsx`; there is no generated
fixture fallback, missing-binary skip, or local native-launch mode.

The harness opens a byte-identical copy of that download in LibreOffice, saves
XLSX, and opens/saves that result again using a new isolated profile. It then
authors disposable input-change copies with openpyxl, preserving formulas and
clearing formula caches, and repeats native save/reopen for all 12 states. Every
native invocation uses a fresh temporary profile configured to recalculate OOXML
formulas on load. openpyxl only authors inputs and reads the consumer's saved
caches; it does not calculate expected results or populate result caches.

All 26 native conversions must succeed. Both the first save and reopened save
are checked for numeric results, live formulas, correct inputs, and a sole Plan
sheet. The original browser download remains untouched. JSON evidence records
the actual LibreOffice version, input and output SHA-256 hashes, conversion
results, values, formulas, and every state. Failed or incomplete runs write
failure/partial evidence and exit nonzero.

Evidence: `test-results/native/evidence.json` and the saved/reopened workbooks in
its `artifacts_directory`. The [verified hosted run](https://github.com/Masanori-Spec/sheet-detach/actions/runs/37215341186)
passed all 26 conversions in LibreOffice 7.3.7.2; see
[exact evidence and scope](VERIFICATION.md). `python3 tools/libreoffice.py --validate` is safe for
local syntax and configuration checks, never launches LibreOffice, and cannot
produce successful native-recalculation evidence. This does not establish
Microsoft Excel compatibility or general Excel formula support.

References:

- [LibreOffice command-line options](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)
- [Calc formula recalculation options](https://help.libreoffice.org/latest/en-US/text/shared/optionen/01060900.html)
- [LibreOffice Calc configuration schema](https://github.com/LibreOffice/core/blob/master/officecfg/registry/schema/org/openoffice/Office/Calc.xcs), `Formula/Load/OOXMLRecalcMode`: 0 means always recalculate
