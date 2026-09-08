# Depiction, typography and geometry

Native `Display="Bold"` has constant width; `WedgeBegin` and `WedgedHashBegin` encode different directional depictions. Preserve bond endpoint direction and chemistry when making a requested visual correction. Never classify all bold-looking colored bonds as wedges. The correct style is determined per bond from the supplied source, not by color alone.

CDXML `<s>` runs: face 1 bold, 32 subscript, 64 superscript, 65 bold superscript. Use ordinary digit text with these styles instead of Unicode subscript glyphs. Verify in ChemDraw after import. Set explicit label display/alignment for R superscripts if automatic layout reverses the label. Color indices start with reserved entries; inspect the actual color table instead of assuming a color ID globally.

Atom labels and PPT compound numbers are separate objects. Changing label font does not bold the PPT number. Record typography for each source region; preserve font weight, not just font family.

Coordinate scaling with constant label font changes the skeleton-to-label ratio. To shorten bonds 10%, scale geometry 0.9, preserve font sizes/line widths, and freeze the calibrated PPT `nativeScale`. Include text anchor coordinates and attached graphics in a consistent transform. A subsequent fit-to-box pass otherwise enlarges the molecule again.

N–Ts visible gaps depend on glyph width, label alignment, margins and node distance. Do not repeatedly increase/decrease an arbitrary factor without viewing the result. Move the collapsed-group anchor and its label consistently; if its internal fragment moves, keep the external connection point anchored to the parent structure. Regenerate from a fixed baseline to prevent cumulative changes.

A resin bead can use native filled circles with explicit center and axes, layered for shading. Keep it inside the ChemDraw document if it should edit with the catalyst. It is an illustrative carrier bead, not reconstructed 3D coordinates. Do not use image generation to draw chemical bonds or manufacture molecular evidence.

Changing `BondLength` metadata alone may not move existing coordinates. Editing a generated `molecules_styled` file may be overwritten by the next preparation pass; inspect the project's generator and edit the authoritative baseline/parameters.
