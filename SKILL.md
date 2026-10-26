---
name: chemdraw-native-ppt
description: "用 ChemDraw 原生对象绘制和修改分子，从 SMILES 构造可编辑 CDXML，通过 Windows COM 批量导出分子图并生成 OCSR 标注；也将化学文献图片或 PDF 图重建为含原生分子 OLE 的可编辑 PPT。"
---

# ChemDraw 原生分子绘制、标注与可编辑 PPT

Choose the deliverable from the user's task: editable CDXML/CDX, molecular images with annotations, or an editable PPT. For PPT reconstruction, deliver native ChemDraw OLE molecules, native PPT text/graphics, a standalone reproducible project, and verified previews. A molecule-generation task does not require PowerPoint or OLE. Use the requested output directory; otherwise prefer Downloads on Windows.

## Route the task

- SMILES-to-CDXML, native molecular image export, synthetic OCSR annotation, or batch generation: read [native-generation-and-annotation.md](references/native-generation-and-annotation.md). Use the owning project's validated generator when present; preserve its source identities, schema, style policy, and frozen-run boundaries.
- Image/PDF figure reconstruction: inspect the supplied image/PDF first. Read [reconstruction.md](references/reconstruction.md), inventory its elements, and reconstruct its actual molecules. Initialize with `scripts/new_project.py --mode empty --destination <new-directory>`.
- Existing project: inspect `project.json`, source image and molecule files. Preserve approved chemistry and styling. A self-authored ethanol/benzene demonstration is available with `--mode example`; read [example.md](references/example.md). Never substitute it for a new image.
- PPT rendering/importing: read [project-format.md](references/project-format.md). The renderer supports arbitrary nonzero molecule counts and multiple slides. A completed project runs through its `start.cmd` without the installed skill.

## Image-to-project workflow

1. Identify molecule IDs/source regions, captions, yields, arrows, variable groups and special graphics. Retain source pixels; record bounding boxes as `[left, top, right, bottom]` in pixels with exclusive right/bottom. For PDFs retain page/figure anchors. Use available PDF/image tools for extraction and locate their runtime rather than assuming an environment.
2. Recover each molecular graph from the visible evidence. Available specialist OCSR can propose candidates; compare atom identities, connectivity, ring substitution, charge, R-groups, isotopes and stereochemistry against each crop. RDKit parseability is not image correctness. Record unresolved features without inventing obscured connectivity. If source evidence is insufficient, ask a focused question for the affected feature and progress on other objects.
3. Write editable CDXML or obtain genuine CDX from ChemDraw. Separate structure from depiction: Bold means a constant-width bond; solid/hashed wedges are directional styles. Inspect every special bond in repeated series. Use native text runs for bold/subscript/superscript, preserve label direction, and check expandability of chemically meaningful abbreviations.
4. Compose other content using manifest `elements` or an independently generated PPT template. Charts need source data; crystallographic reconstruction needs structural data. Photographs/microscopy/crystallographic renderings without reconstruction data require recorded raster exceptions. If fully vector output is essential and evidence is missing, name the remaining exception rather than claiming complete vectorization.
5. Run `powershell.exe -NoProfile -STA -File <project>/build.ps1`. Requires working local Windows ChemDraw and Microsoft PowerPoint. The script imports from ASCII temporary paths, saves a new PPT, verifies saved OLE objects and exports previews. It does not recognize images or prove chemistry.
6. Inspect every slide preview and dense molecular regions. Fix clipping, spacing, labels, special bonds, Ts gaps and source-specific details. Check `validation.json`, then verify interactive native editing where desktop access permits. Report structural OLE checks, visual inspection and interactive editing separately; a correct ProgID alone does not prove correct content.
7. Deliver the PPT and standalone project including reference, sources, manifest and launcher. State any unresolved chemistry/raster exceptions. Edit projects rather than writing generated files into the installed skill.

## Styling changes

Use immutable baseline sources with explicit transformations; read [cdxml-style.md](references/cdxml-style.md). Geometry changes and fit-to-box scaling interact: for 10% shorter skeletons with unchanged fonts, scale coordinates by 0.9 and retain calibrated `nativeScale`. If no scale exists, calibrate using the recorded native dimensions/scale from a first render. Check rendered N–Ts gaps: visible line length differs from atom-center distance because labels clip the line.

Example font/line widths and Ts settings are calibrated for one figure, not universal defaults. ChemDraw Clean Up Structure can normalize geometry but is not screenshot matching and may change wedge placement; inspect after cleanup.

## Operating boundaries

Use existing authorized tools/credentials. This skill adds no OCSR model, recognition API or automatic network upload. If native desktop export is unavailable, prepare the project and label export untested, rather than substituting molecule images. Keep source files and previous outputs. Do not routinely modify execution policies, registry associations or installed applications.

Read [troubleshooting.md](references/troubleshooting.md) for import errors. Load only references relevant to the current branch.
