# Third-party software

Runtime dependencies are pinned in package-lock.json. JSZip uses its MIT option.
Transitive runtime notices are included in third-party/. No project license is
assigned by this document.

- core-util-is 1.0.3: MIT
- immediate 3.0.6: MIT
- inherits 2.0.4: ISC
- isarray 1.0.0: MIT
- jszip 3.10.1: (MIT OR GPL-3.0-or-later)
- lie 3.3.0: MIT
- pako 1.0.11: (MIT AND Zlib)
- process-nextick-args 2.0.1: MIT
- readable-stream 2.3.8: MIT
- safe-buffer 5.1.2: MIT
- sax 1.6.1: BlueOak-1.0.0
- setimmediate 1.0.5: MIT
- string_decoder 1.1.1: MIT
- util-deprecate 1.0.2: MIT
- xml-js 1.6.11: MIT

Build/test-only tools: esbuild 0.25.11 (MIT), Playwright 1.56.0 (Apache-2.0),
openpyxl 3.1.5 (MIT), and hosted distribution LibreOffice (MPL/LGPL terms).
See their upstream packages for complete notices.

The standard DrawingML theme test fixture in fixtures/default-theme.xml is copied
from openpyxl 3.1.5 (MIT). It is used only for import/preservation tests.
