# Native molecular generation and annotation

Use this branch for known molecular graphs, SMILES-to-CDXML generation, ChemDraw native export, and synthetic OCSR datasets. The method transfers structure identity through native objects and measured geometry. It does not recover an unknown molecular graph from an arbitrary PNG; use the reconstruction branch for that task.

## Select the existing implementation

Inspect the owning project's current release pointer, entrypoints, annotation schema, source manifest, and rendering policy before changing a generator. In projects using `tools/chemdraw_dataset/`, useful seams include:

| Concern | Implementation to inspect |
| --- | --- |
| Current version, immutable runs | `CURRENT.json`, `pipeline.py`, `bulk_run.py` |
| Source graph and CDXML objects | `runtime/baseline_rules.py`, `runtime/native_draw.py` |
| Dictionary and dual-port definitions | `runtime/v10_catalog.py`, `twoport_abbreviations.py` |
| Native batching, fonts, pixel anchors | `batch_test.py`, `native_labels.py`, `visible_symbols.py` |
| Source H and repeated units | `explicit_hydrogens.py`, `repeat_units.py` |
| Source chemistry and local repair | `runtime/validation.py`, `repair_study.py`, `repair_policy.py` |
| Final scene coordinates and delivery | `production_scenes.py`, `parallel_scenes.py`, `bulk_run.py` |

These are implementation locators, not dependencies bundled with this skill. On another machine, locate the corresponding code or implement and validate the same contracts. Read actual configuration instead of copying a historical version number, probability table, bond-class numbering, or benchmark result into a new project.

## Graph-to-image contract

1. Parse and validate the source graph. Record immutable source identity and atom provenance before collapsing groups or adding depiction variants. Keep the original source separate from any approved Markush, repeat, or stereo generalization.
2. Generate 2D coordinates, then write native CDXML atoms, bonds, text runs, fragment definitions, and approved document graphics. RDKit supplies chemical graph operations/layout; ChemDraw supplies native depiction and export.
3. Open the file in ChemDraw, save native CDXML, and export SVG. Read back actual node IDs, text boxes, font/color tables, and graphics. Bind objects through explicit ID maps; native import/save can normalize tables and regenerate helper IDs.
4. Determine the document-to-SVG transform from independent control markers or another verified native calibration. Markers must be excluded from delivered images and target labels. Check registration residual and integer rounding separately from semantic anchor correctness.
5. Rasterize the native SVG, preserving its actual fonts. For text outlining, resolve the exact font family, weight, italic style, and glyph metrics; an unsupported font needs an explicit resolution, not silent fallback. Apply a saved crop/padding/scale matrix to every target coordinate and complete text box.
6. Validate native chemistry, target JSON, actual visible glyphs, and final-image geometry before publishing a sample. A successful `SaveAs` or a parseable SMILES proves neither complete visual quality nor annotation correctness.

CPU-side graph construction, CDXML writing, SVG rasterization, and JSON assembly can run on a server when their dependencies and fonts are available. The desktop COM exporter described here runs on Windows with a working registered ChemDraw installation. An official server rendering service is a separate integration requiring verification; do not assume this COM workflow provides one. PowerPoint is needed only for the PPT/OLE branch.

## Minimal native export

This example exports an existing chemical CDXML. It does not implement structure generation, annotation, or batch supervision. Run it only after confirming no other task owns the native controller. Use a new output directory.

```python
from pathlib import Path
import pythoncom
import win32com.client

source = Path("molecule.cdxml").resolve(strict=True)
output = Path("export_native").resolve()
output.mkdir(exist_ok=False)

pythoncom.CoInitialize()
try:
    # Verified with a locally registered 64-bit ChemDraw installation.
    # Resolve the installed ProgID if another supported installation differs.
    app = win32com.client.Dispatch("ChemDraw_x64.Application")
    doc = app.Documents.Open(str(source))
    try:
        doc.SaveAs(str(output / "native.cdxml"), "text/xml")
        doc.SaveAs(str(output / "molecule.svg"), "image/svg+xml")
    finally:
        doc.Close(2)
finally:
    pythoncom.CoUninitialize()
```

Close only the document opened by this operation. Preserve the user's application and documents. Reuse the project's tested normalization/save sequence where it needs additional roundtrips; do not replace that sequence solely with this minimal example.

## Native semantics that must survive rendering

### Abbreviations and generic labels

- Match chemically defined fragments with explicit attachment roots. Keep hidden atoms/bonds in a native `NodeType="Fragment"` and connect external bonds through its `ExternalConnectionPoint` nodes. A printed label alone is not a chemical definition.
- Read the rendered glyph order for the annotation symbol. `OnPr` and `nPrO`, or `NBoc` and `BocN`, may share an expansion while having different visible labels. Preserve a canonical definition/port map separately from the displayed symbol; arbitrary character reversal is not a chemical rule.
- A complete label remains one target node. Anchor it at the actual connection-atom glyph, not the center of the entire text box. Store the box of all characters independently, including superscripts, subscripts, and associated charge symbols.
- Generic `R`, `Ar`, `X`, or `Y` labels need explicit wildcard/placeholder semantics where appropriate. Distinguish a generic class from a definite chemical fragment or an element with the same spelling. Use the project's vocabulary; ordinary OH/HO labels may denote an O atom with attached H, not an abbreviation node.

### Two-port nitrogen abbreviations

For qualified ring-N abbreviations such as NBz, NBoc, NTs, NMe, and their displayed aliases:

- Keep two native external connection points. Both attach to the same internal N; explicitly pair the outer bond ordering with the inner connection order (`BondOrdering` and `ConnectionOrder`). The JSON has one visible label node and two external edges.
- Check hidden expansion and N valence independently of text. Benzoyl Bz and benzyl Bn have distinct definitions. For example, `*N(C(=O)c1ccccc1)*` describes a two-port benzoylated N, while `*N(Cc1ccccc1)*` describes a benzylated N.
- Choose an outward label placement from the incoming bonds, then measure what ChemDraw actually displays. Select the N glyph explicitly during localization: a string such as `NiPr` must not be mistaken for a nickel atom by a generic element-token matcher.
- Reuse the scope established by native prototypes. The validated neutral, achiral ring-N route does not establish correctness for charged/stereogenic N or linkers such as COO whose external ports attach to different internal atoms. Extend those types with their own mapping and native roundtrip evidence.

### Source-specified hydrogen and stereo

When the dataset policy requires source-written H as independent nodes, materialize only those hydrogens, retaining parent/source mapping, isotope, charge, and radical information. Ordinary implicit H need not become separate nodes. Protect required source H and their parents from R replacement or abbreviation collapse.

Inspect bracket-written H counts before stereo cleanup as well as after parsing: cleanup can remove a written H count when a nominal chiral center is actually nonstereogenic. Preserve the written hydrogen without inventing chirality. Existing D/T atoms must not be duplicated.

Canonical comparison may normalize ordinary H for chemistry, so add a separate gate requiring the expected H node, its parent bond, and a visible native H/D/T label with measured bounds. Preserve the project's retained stereo and approved unspecified states when re-wedging. A new explicit H must not silently restore a center deliberately generalized by an earlier depiction rule.

### Repeated fragments and special bonds

- A numeric repeat count equals the actual source units folded. Variable `n/m/x/y/z` labels express unspecified count; the original finite chain is provenance, not proof that the generalized structure is identical to that finite molecule.
- Write native repeated-group membership and crossing-bond references. Use the appropriate `MultipleGroup`/`RepeatCount` or `SRU`/`SRULabel` representation, and independently verify the JSON repeat membership and semantics.
- For repeat counter object tags, changing only the text `p` may be overridden by native automatic positioning. A validated route uses `PositioningType="absolute"` with `PositioningOffset`, refines placement from measured glyph bounds, and translates that absolute offset when moving a batch cell.
- Keep counter search bounded and verify clearance after native save. No acceptable placement is a sample-local failure; preserve the counter's meaning and isolate the sample if approved repair candidates fail.
- Preserve the distinction between constant-width `Bold`, tapered `WedgeBegin`/`WedgedHashBegin`, constant-width `Hash`, ordinary `Dash`, and `Wavy`. Edge-class numbers belong to the project's schema. Converting a directional stereo bond to an unspecified depiction requires synchronized chemical, CDXML, JSON, and validation semantics.

For typography runs, bond clipping, and geometry scaling, read [cdxml-style.md](cdxml-style.md). Use a complete style profile rather than changing only line width. Keep literature-inspired profiles distinguished from verified official journal templates.

## Native batching without font inheritance errors

Maintain one active ChemDraw controller across tasks. Inspect the existing run/controller ownership before starting another. CPU postprocessing can use a bounded worker pool without opening COM connections in its workers.

Group native documents by the full typography and drawing signature: actual font family, label/caption size and face, bond length, line/bold widths, spacing, margins, and relevant display settings. Rebind font and color IDs against the saved document tables, since native saves can renumber them. Mixing incompatible document defaults can silently change label sizes even if individual bond widths are explicit.

For each merged cell, map every ID and reference, translate coordinates and absolute object-tag offsets, and preserve ownership of fragments, brackets, and auxiliary graphics. Split using those maps and native measurements. ViewBox cropping alone does not prove correct membership or annotation.

Benchmark end-to-end time per selected source and per accepted sample, including native opening/saving, splitting, font outlining, rasterization, chemistry checks, scenes, and writing. Report native-only timing separately. Larger batches and more CPU workers do not imply proportional speedup.

## Quality, isolation, and recoverable runs

Publish only samples passing the agreed chemical and geometric gates: expected versus native/JSON representative structure, retained stereo, actual visible symbol, required H, complete label boxes, visible bond strokes, spacing, and final coordinates. Inspect representative images as well; automated acceptance is not expert chemical audit or evidence of training improvement.

Keep sample-local layout/data exceptions inside the sample boundary, including native charge/mark/repeat preparation before export. Exclude failed preparations from subsequent native export, record the ID/stage/reason, and continue healthy neighbors. An all-failed group still needs explicit per-sample terminal records instead of an empty native document. Preserve system failures such as disk IO, out-of-memory, or controller errors as run-level failures requiring appropriate recovery.

Freeze source order, recipe/seed, code hashes, and deadline before a long run. Retain checkpoints and accepted-file hashes. On restart, verify completed commits and resume the first uncommitted block. For Windows sharing violations on atomically replaced activity/checkpoint files, use bounded read/write retries and keep the previous valid snapshot; distinguish a temporary file-access race from invalid data.

Identify a process by command line, creation time, and ownership as well as PID; PIDs can be reused. Confirm that no native controller remains before a safe restart. Preserve the original deadline and record recovery evidence. A code fix in a frozen run needs an authorized, backed-up migration with new hashes, unchanged-source checks, and regression evidence; a retry is not permission to edit frozen code or bypass integrity checks.

## Deliverable and validation boundary

The usual target triplet is editable CDXML, final-size PNG, and compact target JSON (`symbols`, `coords`, `bonds`, plus approved repeat fields). Keep source/definition mapping, actual complete labels, final coordinate matrices, and non-target ownership in sidecar metadata or a manifest.

For a scene containing context text, adjacent fragments, frames, or other clutter, apply the same recorded transformation to target pixels and labels, and keep non-target objects out of the molecular target graph. State that a target-only CDXML does not reconstruct every background object in the final screenshot. Cap inspection overlays according to the task rather than generating one for every sample.

Pair files through the accepted manifest, not a directory glob while generation is active. Publish a compact triplet only after validated paired promotion; preserve the recovery journal if an interrupted move could otherwise mix images and JSON. Separate isolated sources from accepted training data.

Completion evidence should state actual selected/accepted/isolated/pending counts, file and hash verification, render/annotation checks, and visual inspection scope. A minimal COM example, successful export, or small coverage pilot must not be described as a complete portable dataset generator or proof of corpus-wide success.
