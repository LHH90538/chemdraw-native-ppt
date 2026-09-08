# Windows PowerShell 5.1 / STA, installed Microsoft PowerPoint and ChemDraw.
[CmdletBinding()]
param(
    [string]$Project = '',
    [string]$OutputDirectory = (Join-Path $env:USERPROFILE 'Downloads'),
    [ValidateRange(0,5000)][int]$StepDelayMs = 450
)
$ErrorActionPreference = 'Stop'
if (-not $Project) { $Project = Join-Path $PSScriptRoot 'project.json' }
$root = Split-Path -Parent ([IO.Path]::GetFullPath($Project))
$stamp = (Get-Date -Format 'yyyyMMdd_HHmmss_fff')
$work = Join-Path $root ('work/' + $stamp)
[void](New-Item -ItemType Directory -Path $work -Force)
$logPath = Join-Path $work 'build.log'
function Log($message) {
    $line = (Get-Date -Format 'HH:mm:ss') + ' ' + $message
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}
function LocalFile($relative) {
    if ([string]::IsNullOrWhiteSpace($relative)) { throw 'Empty project file path.' }
    $path = [IO.Path]::GetFullPath((Join-Path $root $relative))
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw ('Missing file: '+$path) }
    return $path
}
function RGB($hex) {
    if (-not $hex) { return 0 }
    $s=([string]$hex).TrimStart('#')
    if ($s -notmatch '^[0-9A-Fa-f]{6}$') { throw ('Invalid RGB color: '+$hex) }
    return [Convert]::ToInt32($s.Substring(0,2),16) + 256*[Convert]::ToInt32($s.Substring(2,2),16) + 65536*[Convert]::ToInt32($s.Substring(4,2),16)
}
$deck=$null
try {
    $cfg=Get-Content -LiteralPath $Project -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($cfg.prepareScript) { & (LocalFile $cfg.prepareScript) }
    $items=@($cfg.molecules)
    if ($items.Count -lt 1) { throw 'No molecules configured. Reconstruct the supplied figure and populate project.json first.' }
    if ([double]$cfg.canvasWidth -le 0 -or [double]$cfg.canvasHeight -le 0 -or [int]$cfg.slideCount -lt 1) { throw 'Invalid canvas or slide count.' }
    $seen=@{}
    foreach($item in $items) {
        if (-not $item.name -or $seen.ContainsKey([string]$item.name)) { throw 'Molecule names must be nonempty and unique.' }
        $seen[[string]$item.name]=$true
        if ([int]$item.slide -lt 1 -or [int]$item.slide -gt [int]$cfg.slideCount) { throw ('Invalid slide: '+$item.name) }
        if ([double]$item.width -le 0 -or [double]$item.height -le 0) { throw ('Invalid molecule box: '+$item.name) }
        $source=LocalFile $item.file
        if ([IO.Path]::GetExtension($source) -notmatch '^\.(cdxml|cdx)$') { throw ('Expected CDXML/CDX: '+$source) }
        if ($source.EndsWith('.cdxml',[StringComparison]::OrdinalIgnoreCase)) {
            [xml]$cx=Get-Content -LiteralPath $source -Raw -Encoding UTF8
            if ($cx.DocumentElement.Name -ne 'CDXML' -or $cx.SelectNodes('//fragment/n').Count -eq 0) { throw ('Empty or invalid CDXML: '+$item.name) }
            foreach($expected in @($item.bondChecks)) {
                if ($null -eq $expected) { continue }
                $bond=$cx.SelectSingleNode('//b[@id="'+$expected.id+'"]')
                if ($null -eq $bond -or $bond.GetAttribute('Display') -ne $expected.display) { throw ('Bond style mismatch: '+$item.name+' / '+$expected.id) }
            }
        }
    }
    [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
    $stage=Join-Path ([IO.Path]::GetTempPath()) ('ChemDrawPPT_'+[guid]::NewGuid().ToString('N'))
    if ($stage -match '[^\x00-\x7F]') { throw 'An ASCII temporary path is required for this ChemDraw import route. Configure an accessible ASCII TEMP directory for this command.' }
    [void](New-Item -ItemType Directory -Path $stage)
    $pp=New-Object -ComObject PowerPoint.Application
    $pp.Visible=-1
    if ($cfg.template) { $deck=$pp.Presentations.Open((LocalFile $cfg.template),0,-1,-1) }
    else {
        $deck=$pp.Presentations.Add(-1)
        $deck.PageSetup.SlideWidth=[single]$cfg.slideWidth
        $deck.PageSetup.SlideHeight=[single]$cfg.slideHeight
    }
    while($deck.Slides.Count -lt [int]$cfg.slideCount) { [void]$deck.Slides.Add($deck.Slides.Count+1,12) }
    if ($deck.Slides.Count -ne [int]$cfg.slideCount) { throw 'Template slide count differs from project slideCount.' }
    foreach($slide in $deck.Slides) {
        foreach($shape in $slide.Shapes) { if ($shape.Type -eq 7 -or $shape.Type -eq 10) { throw 'Template must not contain existing OLE objects. Put molecules in the manifest.' } }
    }
    $sx=[double]$deck.PageSetup.SlideWidth/[double]$cfg.canvasWidth
    $sy=[double]$deck.PageSetup.SlideHeight/[double]$cfg.canvasHeight
    $window=$deck.Windows.Item(1);$window.Activate()
    foreach($e in @($cfg.elements)) {
        if ($null -eq $e) { continue }
        if ([int]$e.slide -lt 1 -or [int]$e.slide -gt [int]$cfg.slideCount) { throw 'Invalid element slide.' }
        $slide=$deck.Slides.Item([int]$e.slide)
        $x=[single]($e.left*$sx);$y=[single]($e.top*$sy);$w=[single]($e.width*$sx);$h=[single]($e.height*$sy)
        switch($e.type) {
            'text' {
                $shape=$slide.Shapes.AddTextbox(1,$x,$y,$w,$h)
                $shape.TextFrame.MarginLeft=0;$shape.TextFrame.MarginRight=0;$shape.TextFrame.MarginTop=0;$shape.TextFrame.MarginBottom=0
                $tr=$shape.TextFrame.TextRange;$tr.Text=[string]$e.text
                $tr.Font.Name=if($e.font){[string]$e.font}else{'Arial'}
                $fontSizeValue=if($e.size){$e.size}else{12}
                $tr.Font.Size=[single]$fontSizeValue
                $tr.Font.Color.RGB=RGB $e.color
                if ($e.bold) { $tr.Font.Bold=-1 }
                if ($e.align -eq 'center') { $tr.ParagraphFormat.Alignment=2 }
                foreach($run in @($e.runs)) {
                    if ($null -eq $run) {continue}
                    if ($run.start -lt 1 -or $run.length -lt 1 -or ($run.start+$run.length-1) -gt $tr.Length) { throw 'Invalid text run range.' }
                    $part=$tr.Characters([int]$run.start,[int]$run.length)
                    if ($run.bold) {$part.Font.Bold=-1}
                    if ($run.subscript) {$part.Font.Subscript=-1}
                    if ($run.superscript) {$part.Font.Superscript=-1}
                    if ($run.color) {$part.Font.Color.RGB=RGB $run.color}
                }
            }
            'line' { $shape=$slide.Shapes.AddLine($x,$y,[single]($x+$w),[single]($y+$h)); if($e.arrow){$shape.Line.EndArrowheadStyle=3} }
            'rectangle' { $shape=$slide.Shapes.AddShape(1,$x,$y,$w,$h) }
            'ellipse' { $shape=$slide.Shapes.AddShape(9,$x,$y,$w,$h) }
            default { throw ('Unsupported native element: '+$e.type) }
        }
        if ($e.type -ne 'text') {
            $shape.Line.ForeColor.RGB=RGB $e.color
            $weightValue=if($e.lineWidth){$e.lineWidth}else{1}
            $shape.Line.Weight=[single]$weightValue
            if ($e.type -ne 'line') { if($e.fill){$shape.Fill.Solid();$shape.Fill.ForeColor.RGB=RGB $e.fill}else{$shape.Fill.Visible=0} }
        }
        if ($e.name) {$shape.Name=[string]$e.name}
    }
    foreach($edit in @($cfg.textEdits)) {
        if($null -eq $edit){continue}
        $matchedCount=0
        foreach($shape in $deck.Slides.Item([int]$edit.slide).Shapes) {
            if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $tr=$shape.TextFrame.TextRange
                if($tr.Text -match [string]$edit.pattern) {
                    if($edit.start -lt 1 -or $edit.length -lt 1 -or ($edit.start+$edit.length-1) -gt $tr.Length){throw 'Invalid text edit range.'}
                    $tr.Characters([int]$edit.start,[int]$edit.length).Font.Bold=-1;$matchedCount++
                }
            }
        }
        if($matchedCount -ne [int]$edit.expectedMatches){throw ('Text edit match count: '+$edit.pattern+' / '+$matchedCount)}
    }
    $results=@();$i=0
    foreach($item in $items) {
        $i++;$slide=$deck.Slides.Item([int]$item.slide)
        $window.View.GotoSlide([int]$item.slide)
        $source=LocalFile $item.file
        $oleFile=Join-Path $stage ($i.ToString()+[IO.Path]::GetExtension($source))
        Copy-Item -LiteralPath $source -Destination $oleFile
        $ole=$slide.Shapes.AddOLEObject(0,0,-1,-1,'',[string]$oleFile,0,'',0,'',0)
        $prog=[string]$ole.OLEFormat.ProgID
        if($ole.Type -ne 7 -or $prog -notmatch 'ChemDraw|CSChem'){throw ('Not a ChemDraw embedded object: '+$item.name+' / '+$prog)}
        $nw=[double]$ole.Width;$nh=[double]$ole.Height
        if($nw -le 0 -or $nh -le 0){throw 'Empty OLE preview.'}
        $tw=[double]$item.width*$sx;$th=[double]$item.height*$sy
        $scale=if($item.nativeScale){[double]$item.nativeScale}else{[Math]::Min($tw/$nw,$th/$nh)}
        if($scale -le 0){throw 'Invalid nativeScale.'}
        $ole.LockAspectRatio=0;$ole.Width=$nw*$scale;$ole.Height=$nh*$scale;$ole.LockAspectRatio=-1
        $ole.Left=[double]$item.left*$sx+($tw-$ole.Width)/2;$ole.Top=[double]$item.top*$sy+($th-$ole.Height)/2
        $ole.Name='ChemDraw_'+$item.name
        $ole.AlternativeText='Native ChemDraw: '+$item.name
        $results+=@{name=$item.name;slide=$item.slide;progId=$prog;nativeWidth=$nw;nativeHeight=$nh;scale=$scale;source=$item.file;chemicalReview=$item.chemicalReview}
        Log ($i.ToString()+'/'+$items.Count+': '+$item.name+' ['+$prog+']')
        $window.Activate();Start-Sleep -Milliseconds $StepDelayMs
    }
    $draft=Join-Path $work 'candidate.pptx';$deck.SaveAs($draft,24)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($draft)
    try {
        $saved=0
        for($s=1;$s -le [int]$cfg.slideCount;$s++) {
            $entry=$zip.GetEntry('ppt/slides/slide'+$s+'.xml')
            $reader=[IO.StreamReader]::new($entry.Open());try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
            $objects=@($xml.SelectNodes('//*[local-name()="oleObj"]') | Group-Object { $_.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships') } | ForEach-Object {$_.Group[0]})
            $expected=@($items | Where-Object {[int]$_.slide -eq $s}).Count
            if($objects.Count -ne $expected){throw ('Saved OLE count mismatch on slide '+$s)}
            foreach($obj in $objects){if($obj.progId -notmatch 'ChemDraw|CSChem'){throw 'Saved non-ChemDraw object.'}}
            $saved+=$objects.Count
        }
    } finally {$zip.Dispose()}
    $prefix=([string]$cfg.name -replace '[^a-zA-Z0-9_-]','_');if(-not $prefix){$prefix='Chemical_Figure'}
    $target=Join-Path $OutputDirectory ($prefix+'_'+$stamp+'.pptx')
    $deck.SaveAs($target,24)
    $preview=Join-Path $work 'previews';[void](New-Item -ItemType Directory -Path $preview)
    for($s=1;$s -le $deck.Slides.Count;$s++) {$deck.Slides.Item($s).Export((Join-Path $preview ('slide_'+$s+'.png')),'PNG',1600,[int](1600*$deck.PageSetup.SlideHeight/$deck.PageSetup.SlideWidth))}
    $report=@{pptx=$target;oleValidated=$true;moleculeCount=$saved;molecules=$results;visualReview='pending';interactiveEditing='not_tested';rasterExceptions=@($cfg.rasterExceptions);previews=$preview}
    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $work 'validation.json') -Encoding UTF8
    Log ('SUCCESS: '+$target);Log ('REVIEW PREVIEWS: '+$preview)
} catch {
    Log ('FAILED: '+$_.Exception.Message)
    Log ('POSITION: '+$_.InvocationInfo.PositionMessage)
    if($null -ne $deck){try{$deck.SaveAs((Join-Path $work 'PARTIAL_NOT_COMPLETE.pptx'),24)}catch{}}
    exit 1
}
