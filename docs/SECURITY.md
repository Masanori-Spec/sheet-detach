# Data and safety boundaries

The generated standalone HTML bundles its runtime dependencies. It makes no
network requests, uses no analytics or local storage, and does not upload files.
Files are processed in browser memory. Downloads happen only on explicit button
activation. Changing the input or selection discards stale previews/downloads.

The XLSX reader permits only documented plain parts. Macros, embedded objects,
external relationships, tables, drawings, unsupported worksheet features and
rich text fail closed. ZIP entry metadata, local/central consistency and descriptors are checked
before sequential bounded pako decompression. Actual emitted bytes are capped
during inflation and CRCs are checked in the same pass. JSZip is export-only.
Strict lexical and namespace-aware SAX checks bound XML depth/nodes before
xml-js allocates a tree; recursive feature allowlists then validate retained
content. DTD/entities, duplicate attributes, conflicting cell payloads and
forbidden XML controls are rejected. Expanded references and lineage evidence
have separate cumulative limits. Oversized files are rejected before the UI
calls File.arrayBuffer.
CSV string fields with formula-like prefixes receive an apostrophe; XLSX text
labels remain inline strings even when they begin with `=`.

This is not a confidentiality scrubber. Frozen values are intentionally included
in the workbook, while the ledger and recipe expose omitted-source lineage.
Cell styles, custom number formats and optional theme metadata are preserved.
Review every handoff file before sharing it. Browser memory is not guaranteed to
be securely erased. Very large valid inputs can still take time within the
profile bounds; no performance SLA or adversarial-security certification exists.

Hosted CI uses a sandboxed Chromium browser and isolated LibreOffice profiles.
No user desktop/profile or personal workbook is used. Native LibreOffice is
restricted to GitHub-hosted Linux CI. Microsoft Excel is not tested. No secrets,
paid services, persistent permissions or cloud backend are required.
