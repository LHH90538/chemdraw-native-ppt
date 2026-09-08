---
name: chemdraw-native-ppt
description: "将化学文献图片或 PDF 图拆解为可编辑 PPT，分子重建为 ChemDraw 原生 CDXML/CDX 与嵌入式 OLE，文字和箭头重建为 PPT 对象；也用于修改 ChemDraw 图的键型、上下标和一键制作程序。"
---

# 化学文献图转 ChemDraw 原生 PPT

Deliver an editable reconstruction, a standalone reproducible project, and verified previews. Molecules are native ChemDraw OLE, not SVG, traced PPT lines, image previews alone, or Package attachments. Other text, arrows and simple illustrations are native PPT shapes or editable ChemDraw graphics. Use the requested output directory; otherwise prefer Downloads on Windows.

## Route the task

- New figure: inspect the supplied image/PDF first. Read [reconstruction.md](references/reconstruction.md), inventory its elements, and reconstruct its actual molecules. Initialize with `scripts/new_project.py --mode empty --destination <new-directory>`.
- Existing project: inspect `project.json`, source image and molecule files. Preserve approved chemistry and styling. A self-authored ethanol/benzene demonstration is available with `--mode example`; read [example.md](references/example.md). Never substitute it for a new image.
- Rendering/importing: read [project-format.md](references/project-format.md). The renderer supports arbitrary nonzero molecule counts and multiple slides. A completed project runs through its `start.cmd` without the installed skill.

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