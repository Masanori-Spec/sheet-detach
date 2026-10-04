# Existing workflows and the bounded difference

Research date: 2026-10-04. This is a product-workflow comparison, not a patent
search or proof of demand. Existing software already splits sheets and converts
formulas to values.

- **Ablebits Split Workbook** documents preserving referenced source worksheets
  when exporting sheets with formulas. SheetDetach instead offers a restricted
  scalar-only boundary rewrite, with dependency-return blocking.
  https://www.ablebits.com/docs/excel-split-workbook/
- **ASAP Utilities Export worksheets** documents formula, value and displayed-
  value export choices. The reviewed page does not document SheetDetach's mixed
  operand-level dependency guard. This does not prove the feature is absent from
  every product/version or workflow.
  https://www.asap-utilities.com/tools-detail.php?lang=en-us&tip=233&utilities=196
- **Microsoft Excel** supports manually selecting part of a formula and pressing
  F9 to replace that part with its result. SheetDetach applies a bounded,
  reviewable batch transform across selected sheets and records each occurrence.
  https://support.microsoft.com/en-us/excel/replace-a-formula-with-its-result-in-excel
- **Microsoft OOXML documentation** explains that a formula cell's value records
  its last-calculation cache. SheetDetach does not use those caches as truth.
  https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas

The potential use case is handing off an ordinary planning template while
retaining editable local assumptions and formulas. Its usefulness remains to
be validated with actual users. No novel-algorithm, unique-product, patentability,
financial-model reliability, customer demand or measured time-savings claim is
made. No customer outreach, university access or paid purchase was performed.
