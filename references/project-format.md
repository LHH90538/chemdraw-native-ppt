# Standalone project contract

Create with Python: `python <skill>/scripts/new_project.py --mode empty --destination <new-folder>`.
For the self-authored demonstration use `--mode example`. Existing nonempty destinations are rejected. Each project receives `build.ps1`, `start.cmd` and `project.json`. The example also contains its own source molecules. The skill folder is not a runtime dependency.

Run on Windows with usable desktop ChemDraw/PowerPoint: `powershell.exe -NoProfile -STA -File <project>/build.ps1 -OutputDirectory <output-folder>`. Default output is the current Windows user's Downloads. `-StepDelayMs 450` controls visible molecule placement; it is not bond-by-bond drawing. Python is needed for initialization only, not double-click rebuilding.

## Manifest

```json
{
  "name": "My_Figure",
  "canvasWidth": 960, "canvasHeight": 540,
  "slideWidth": 960, "slideHeight": 540, "slideCount": 1,
  "template": null,
  "molecules": [{
    "name": "compound_1", "file": "molecules/compound_1.cdxml", "slide": 1,
    "left": 50, "top": 100, "width": 220, "height": 190,
    "chemicalReview": "pending",
    "bondChecks": [{"id": "31", "display": "Bold"}]
  }],
  "elements": [{
    "type": "text", "name": "caption_1", "slide": 1,
    "left": 50, "top": 310, "width": 220, "height": 30,
    "text": "1 (80%)", "font": "Arial", "size": 14,
    "runs": [{"start": 1, "length": 1, "bold": true}]
  }],
  "textEdits": [], "rasterExceptions": []
}
```

This schema illustration is not a molecule fixture: create the named CDXML and use its actual bond IDs. `molecules` must be nonempty and names unique. All slide indices are one-based. All molecule/element placement coordinates use `canvasWidth`/`canvasHeight`; text sizes and line widths are PowerPoint points. Without a template, `slideWidth`/`slideHeight` are PPT page sizes in points. Template page size takes precedence. Templates must have no preexisting OLE molecules and may not have more slides than `slideCount`.

Molecule `nativeScale` is optional. Omit for initial fit-to-box placement; a positive value overrides fitting and preserves physical rendering scale during geometry edits. The renderer centers the result in its box and records actual scale/native width/height. Calibrate after visual inspection. Large scales may overflow the box; review previews, including slide edges.

`bondChecks` validates specific CDXML bond Display values before embedding. It is a regression guard, not stereochemical validation. CDX is accepted but cannot use XML-specific guards. `chemicalReview` is copied into the report, not certified by the renderer.

Supported native elements: `text`, `line`, `rectangle`, `ellipse`. Common fields: `slide,left,top,width,height,name`. Line width/height describe signed endpoint offsets. Text fields: `text,font,size,bold,color,align` (`center` optional), and `runs`. Run offsets use one-based PowerPoint character positions and lengths; supported run flags are `bold,subscript,superscript,color`. Other shapes support `color,lineWidth,fill`; a line supports `arrow:true`. Colors are six-digit RGB strings. Use a template for advanced shapes, charts, gradients, pictures or rich text beyond this interface.

`textEdits` applies bold to existing template text: `{slide,pattern,start,length,expectedMatches}`. Pattern is a regex. Explicit match count prevents silently formatting nothing or the wrong repeated labels.

`prepareScript` optionally names a project-local PowerShell generator. Inspect it before using an unfamiliar project. It runs before source checks; it should regenerate from immutable baselines. The synthetic example needs no preparation script.

`rasterExceptions` records `{id,slide,reason}`. It is a disclosure list, not an image insertion API. Embed permitted raster content in the template and review it. This field does not automatically detect undeclared raster material.

## Outputs and acceptance

Timestamped PPT goes to the requested directory. `work/<timestamp>/` holds `build.log`, `candidate.pptx`, `previews/slide_N.png`, and `validation.json`. Failures retain `PARTIAL_NOT_COMPLETE.pptx` if possible and return a nonzero exit status.

The renderer verifies unique ChemDraw OLE references per slide, accounting for compatibility branches. Automated checks pass only after native import and saving. Inspect every PNG; record the agent's visual/chemical review in a separate project `review.md`. The generated report intentionally leaves visual and interactive-edit checks pending. Validate interactive editing separately when desktop access is available. Do not turn pending fields into success without the actual check.
