# Bounded transform

1. Inspect the ZIP directory before decompression. Reject unsafe/duplicate paths,
   encryption, ZIP64, multipart archives, unexpected tails and size violations.
   Match central/local headers and descriptors. Inflate one entry at a time
   through bounded pako chunks, checking actual output bytes and CRC before
   accepting text; declared sizes are never the sole decompression limit.
2. Apply a package-part/worksheet-feature allowlist. Resolve workbook relationship
   IDs to worksheets. Reject duplicate attributes, forbidden XML 1.0 characters and declarations.
   Run a namespace-aware SAX depth/node-budget pass before constructing an XML
   tree, then recursively validate allowed children/attributes, cell-payload
   alternatives and style/layout references. Parse XML without DTDs/entities. Load plain labels, numeric
   cells and formula text; ignore formula result caches.
3. Tokenize and parse supported formulas into an AST. Preserve source spans of
   scalar reference nodes. Bind sheet names case-insensitively and normalize A1
   addresses. Reject unsupported syntax rather than guessing.
4. Build the cell-dependency graph. Validate operands and numeric ranges, find
   cycles with an iterative traversal, and evaluate the numeric DAG in
   topological order. Expanded reference visits and unique edges are limited
   during construction, including repeated SUM ranges. Arithmetic evaluation
   and AST walks are iterative.
5. For every retained formula, keep references to retained sheets unchanged.
   Reject any range crossing the boundary. For an omitted scalar reference,
   traverse all transitive dependencies; if any reaches a retained cell, report
   the exact path and block the entire export. Cache shared omitted lineage,
   bound traversal work, and count every emitted evidence node/path cell and
   serialized ledger character, including repeated reference occurrences.
6. Freeze each independent omitted scalar as a parenthesized finite literal.
   Replacement locations come from bound AST nodes, never regex matching of
   formula text. Untouched formula slices retain their original spelling,
   spacing and anchors. Repeated references create distinct ledger rows.
7. Rebuild the XLSX package with selected worksheet payloads, relationship IDs
   and content-type declarations. Convert retained plain shared strings to
   inline strings so omitted shared labels do not leak through that table.
   Keep supported worksheet layout, styles and optional theme. Set automatic
   full recalculation and write independently recomputed formula caches.

Unsupported features block import/export; there is no whole-cell-value fallback.
All kept formulas appear in the preview even if no reference is frozen. The
recipe/ledger record the original formula, replacement, each frozen occurrence,
source cell, recomputed value and dependency lineage. CSV strings are guarded
against spreadsheet formula injection with a leading apostrophe where needed;
JSON retains exact formula text.

The source bytes and parsed worksheet trees are never modified. ZIP output uses
fixed timestamps and deterministic entry ordering for repeatable exports. A
selection/import/reset invalidates the UI's previous analysis and downloads.
