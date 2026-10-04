# Verified behavior and remaining limits

## Exact hosted evidence

[Run 37215341186](https://github.com/Masanori-Spec/sheet-detach/actions/runs/37215341186) passed all three jobs on 2026-10-04 at source commit
`af79b6f78f6f0d5000afb452348e7f4d23f16218`. Its 61 source files were matched to the reviewed local tree by
Git blob hashes. Later evidence/documentation commits do not change that
application or its verification harness. Check the repository Actions page for
the final head's separate run.

- 132 Node tests pass independently on Node 22 and 24
- Python/openpyxl/ZIP oracle: 12 closed-form input states and 8 deliberately
  damaged-export rejections, without importing the application evaluator
- 15 browser checks in sandboxed Chromium: staged imports, keyboard controls,
  stale-result invalidation, reset, repeated import, dependency-return guards,
  unsupported/malformed/oversized inputs and actual XLSX/CSV/JSON/TXT downloads
- Network observation: one local GET, no external request or workbook upload;
  no captured console errors
- JA/EN desktop and 390px mobile screenshots plus blocked-state layouts inspected
- Real JA/EN Chromium print PDFs: two A4 pages each, independently checked with
  Poppler for required formula, ledger and scope text; all four rendered pages
  inspected for clipping and blank pages
- Actual browser-downloaded XLSX passed the independent oracle
- LibreOffice 7.3.7.2 30(Build:2), official Ubuntu 22.04 distribution package:
  26 real native conversions using fresh isolated profiles and the SVP headless
  backend, covering initial save/reopen and 12 changed-input save/reopen states

## Actual editable workbook results

Only Plan remains and C2/C3/C4 remain formulas. The browser-downloaded XLSX SHA-256
is `5468fb273a0ba0fed74909e82f118bd98934a9619d22bba68d01fe4df97f19d2`.
Its bytes were preserved throughout native testing.

- B2=8, B3=2: C2=46, C3=10, C4=56 after native save and reopen
- B2=10, B3=3: C2=58, C3=15, C4=73 after native save and reopen
- All combinations of B2 in 0, 1, 8, 10 and B3 in 0, 2, 3 matched the separate
  closed-form expectation before and after reopening

openpyxl authored disposable input changes and cleared formula caches. It did
not calculate the consumer results or manufacture their caches. Those values
were read from LibreOffice's saved workbooks.

## Evidence files

- [Browser evidence](evidence/hosted/browser-evidence.json)
- [Independent actual-download oracle](evidence/hosted/actual-download-oracle.json)
- [Native conversion and state evidence](evidence/hosted/native-evidence.json)
- [Japanese print PDF](evidence/hosted/handoff-review-ja.pdf) and
  [English print PDF](evidence/hosted/handoff-review-en.pdf)
- Full saved/reopened native workbooks are in the linked run's
  `sheetdetach-browser-and-libreoffice-evidence` artifact. Its ZIP SHA-256 is
  `ed4748d0479366bb8e89ac9492e43b82d9d4c5b8559bf69d6f9c700c20fa2a24`.

## Limits

Microsoft Excel-native behavior remains unverified. Passing these bounded
fixtures does not establish general XLSX compatibility, exact decimal/financial
precision, demand, time savings, novelty or patentability. Formula fill, structural
edits, changes to omitted sheets and unsupported workbook features remain outside
the contract. This is not a confidential-data redaction tool.

The first hosted native attempt failed because the harness requested X11 without
a display. Switching its isolated test backend to SVP resolved that issue. The
passing evidence above supersedes the earlier incomplete runs.
