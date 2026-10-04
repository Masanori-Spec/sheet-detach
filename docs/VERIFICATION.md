# Verification status at the source freeze

## Executed locally

- 132 Node tests pass: formula parsing/arithmetic, all 12 sample input states,
  stale source caches, dependency-return witness, multiple retained sheets,
  quoted/apostrophe/Japanese names, repeated references, numeric precedence,
  zero/negative/decimal values, kept/omitted SUM behavior, blocked unsupported
  syntax, cycle witnesses, 1,500-node DAG, package features and limits, recursive XML
  allowlists, malformed payloads/attributes, namespace bindings, style cross-
  references, actual-byte ZIP inflation limits, repeated-range/lineage budgets,
  and an independently produced plain openpyxl input
- Independent Python/openpyxl/ZIP oracle passes all 12 states without importing
  application code or calculating from caches
- 8 intentionally damaged exports are rejected by the independent oracle
- Python oracle parser self-tests and native harness --validate pass
- Standalone bundle builds; repeated XLSX exports are byte-identical

## Authored, not yet executed at this freeze

- Sandboxed hosted Chromium: JA/EN desktop/mobile, keyboard handling, import,
  selection staging, stale invalidation, guard witness, reimport and downloads
- Independent oracle against the actual browser-downloaded XLSX
- Distribution-provided LibreOffice Calc: actual download, initial recalculation,
  12 input states, fresh-profile native save/reopen (26 conversions total)

The workflow must pass on the published exact head before these stages can be
reported as successful. Keep this status distinction when publishing evidence.
Microsoft Excel-native testing remains unperformed even after LibreOffice passes.
Formula snapshots alone do not establish recalculation in a consumer.

The second source review identified and closed fail-closed/XML and resource-
budget gaps before publication. Earlier 89-test evidence and archive hashes are
superseded by this source snapshot. The source ZIP manifest supplies its hash.
