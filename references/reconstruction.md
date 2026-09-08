# New literature figure

Create a project per figure. Retain the supplied reference and PDF page anchor. Inventory every visible element before reconstructing: molecule, caption/condition, identifier/yield, arrow/line, chart, or image requiring separate treatment. A compact `evidence.json` can hold `{id, source, page, bbox, kind, transcription, chemicalReview, uncertainties}`. The bounding-box convention is `[left,top,right,bottom]`, exclusive right/bottom, in source image pixels.

For each molecule, reconstruct the graph then the drawing. Use visible atom/bond evidence; an OCSR candidate or SMILES is not a validated reconstruction. Compare each ring attachment, heteroatom, double bond, charge, isotope, abbreviation, variable group and stereochemical depiction. Keep ambiguous details recorded until resolved. For repeat-series figures derive a shared core only after verifying it is actually shared, then independently check every substituent and special bond.

Use ChemDraw-authored CDXML/CDX where available. Programmatic CDXML must include unique object IDs, valid atom/bond references, native text styling, and sensible coordinates. ChemDraw's native preview is the rendering authority. If a chemistry toolkit is available, graph round-trip checks supplement visual review; they do not establish correspondence to a blurry image.

Expand labels such as Ph/Ts/CO2H into valid internal fragments with an ExternalConnectionPoint when meaningful editability requires it. Visible abbreviations alone can look correct while being chemically incomplete. R-group semantics and unsupported Markush features must remain explicit. Do not infer missing 3D geometry from a crystallographic picture.

Recreate text and simple arrows with the manifest's native `elements`. For complex layouts, charts or tables, build a template using the available presentation tooling, then supply its path. Keep molecular regions empty of OLE objects in the template. A source image in the reference folder is not part of the rendered output unless explicitly intended.

If a photograph or a plot without recoverable data remains raster, record its location and reason in `rasterExceptions`, and disclose it in delivery. Ordinary molecules cannot use this exception as a shortcut. Delivering a native ChemDraw object with a wrong graph is also a failure.

The bundled scripts do not include an image recognition service. Codex performs image inspection/reconstruction with available tools. This skill requires neither a particular external provider nor new credentials.
