# Native import diagnostics

- For file-based OLE insertion, supply typed numeric/string arguments rather than `System.Reflection.Missing` in width/height positions. The bundled script contains the tested call.
- Chinese paths produced ChemDraw read errors on the tested installation. The renderer copies each molecule into a unique ASCII TEMP path before embedding. Preserve originals. On another machine with a non-ASCII TEMP path, use an accessible ASCII directory for that command's TEMP; do not change global system settings.
- ProgIDs depend on installation. This machine has used `ChemDraw_x64.Document.6.0`; accept the ChemDraw/CSChem family rather than hardcoding this one ID. The current renderer delegates activation to PowerPoint and does not require `ChemDraw.Application`.
- Saving can emit duplicated OLE nodes in XML compatibility branches. Count unique relationship IDs per slide, not every `oleObj` node. Check actual contents and preview in addition to ProgID.
- A clipboard paste failure is not necessarily a copying error. Cleanup exceptions can obscure the primary error. File-based import is preferred where it works; do not repeatedly send the user through manual copying without diagnosing the log.
- Execution-policy, desktop-session or COM-registration failures are environment failures, not proof a molecule is invalid. Keep logs and source files, state what could not be tested, and do not relabel a bitmap substitute as native completion.
- Output objects can be native yet chemically wrong or empty. Compare exported previews to the original source, including the most crowded labels and special bonds.
