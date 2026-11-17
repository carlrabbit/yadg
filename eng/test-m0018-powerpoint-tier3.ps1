param(
    [string] $TemplatePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/powerpoint-origin.pptx'),
    [string] $SourcePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/review-source.md'),
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/runs/powerpoint-origin-generated.pptx'),
    [string] $EvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/powerpoint-tier3-evidence.json'),
    [string] $ConfigPath,
    [switch] $Yolo,
    [switch] $SkipBuild,
    [string] $Milestone = 'M0018',
    [string] $Gate = 'VAL-06'
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$template = [IO.Path]::GetFullPath($TemplatePath)
$source = [IO.Path]::GetFullPath($SourcePath)
$output = [IO.Path]::GetFullPath($OutputPath)
$evidence = [IO.Path]::GetFullPath($EvidencePath)
$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.dll'
$provenancePath = Join-Path (Split-Path -Parent $template) 'powerpoint-origin.provenance.json'
$provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json -DateKind String
if ((Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$provenance.sha256) { throw 'PowerPoint-origin fixture bytes do not match their provenance hash.' }
if (-not (Test-Path -LiteralPath $cli -PathType Leaf)) { throw "Build the Release CLI before Tier 3: $cli" }
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Presentation source is missing: $source" }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $output),(Split-Path -Parent $evidence) | Out-Null

$revision = (& git -C $repo rev-parse HEAD).Trim()
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToUpperInvariant()
$configFullPath = if ($ConfigPath) { [IO.Path]::GetFullPath($ConfigPath) } else { $null }
$configHash = if ($configFullPath) { (Get-FileHash -LiteralPath $configFullPath -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
$templateHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant()
$app = $null
$presentation = $null
$templatePresentation = $null
$reviewCandidatePresentation = $null
$appWasRunning = $false
$failure = $null
$applicationVersion = $null
$inspect = $null
$formattingValidation = $null
$result = 'failed'

function Get-M0018ShapeByName($document, [int] $slideIndex, [string] $name) {
    foreach ($shape in $document.Slides.Item($slideIndex).Shapes) {
        if ([string]$shape.Name -ceq $name) { return $shape }
    }
    throw "PowerPoint slide $slideIndex is missing expected shape '$name'."
}

function Get-M0018Paragraph($shape, [string] $expectedText) {
    foreach ($paragraph in $shape.TextFrame.TextRange.Paragraphs()) {
        if ([string]$paragraph.Text.Trim() -ceq $expectedText -or ([string]$paragraph.Text.Trim()).StartsWith($expectedText, [StringComparison]::Ordinal)) { return $paragraph }
    }
    throw "PowerPoint shape '$($shape.Name)' is missing formatting exemplar/text '$expectedText'."
}

function Get-M0018TextFormatting($paragraph) {
    $font = $paragraph.Font
    $format = $paragraph.ParagraphFormat
    $bullet = $format.Bullet
    return [ordered]@{
        fontName = [string]$font.Name
        fontSize = [double]$font.Size
        bold = [int]$font.Bold
        italic = [int]$font.Italic
        underline = [int]$font.Underline
        colorRgb = [int]$font.Color.RGB
        alignment = [int]$format.Alignment
        indentLevel = [int]$paragraph.IndentLevel
        bulletVisible = [int]$bullet.Visible
        bulletType = [int]$bullet.Type
        bulletCharacter = [int]$bullet.Character
    }
}

function Compare-M0018Formatting([string]$role, $templateFormatting, $generatedFormatting) {
    foreach ($key in $templateFormatting.Keys) {
        if ([string]$templateFormatting[$key] -cne [string]$generatedFormatting[$key]) {
            throw "PowerPoint generated $role formatting differs from its native template example for '$key': template='$($templateFormatting[$key])', generated='$($generatedFormatting[$key])'."
        }
    }
    return [ordered]@{ role = $role; result = 'matched-template-example'; template = $templateFormatting; generated = $generatedFormatting }
}

function Get-M0018TableCellFormatting($slide, [int] $row, [int] $column) {
    foreach ($shape in $slide.Shapes) {
        try {
            if ($shape.HasTable -eq -1) {
                return Get-M0018TextFormatting $shape.Table.Cell($row, $column).Shape.TextFrame.TextRange.Paragraphs(1)
            }
        } catch {}
    }
    throw "PowerPoint slide $($slide.SlideIndex) has no native table cell at ($row,$column)."
}

function Get-M0018TableGeometry($slide) {
    foreach ($shape in $slide.Shapes) {
        try {
            if ($shape.HasTable -eq -1) {
                $rowHeights = @(); for ($row = 1; $row -le $shape.Table.Rows.Count; $row++) { $rowHeights += [double]$shape.Table.Rows.Item($row).Height }
                $columnWidths = @(); for ($column = 1; $column -le $shape.Table.Columns.Count; $column++) { $columnWidths += [double]$shape.Table.Columns.Item($column).Width }
                return [ordered]@{ left=[double]$shape.Left; top=[double]$shape.Top; width=[double]$shape.Width; height=[double]$shape.Height; rowHeights=$rowHeights; columnWidths=$columnWidths }
            }
        } catch {}
    }
    throw "PowerPoint slide $($slide.SlideIndex) has no measurable native editable table."
}

try {
    if (-not $SkipBuild) {
        $arguments = @($cli, 'presentation', 'check', '--template', $template, '--source', $source)
        if ($ConfigPath) { $arguments += @('--config', [IO.Path]::GetFullPath($ConfigPath)) }
        if ($Yolo) { $arguments += '--yolo' }
        & dotnet @arguments
        if ($LASTEXITCODE -ne 0) { throw 'PowerPoint-origin fixture failed presentation preflight.' }
        $buildArguments = @($cli, 'presentation', 'build', '--template', $template, '--source', $source, '--output', $output)
        if ($ConfigPath) { $buildArguments += @('--config', [IO.Path]::GetFullPath($ConfigPath)) }
        if ($Yolo) { $buildArguments += '--yolo' }
        & dotnet @buildArguments
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $output -PathType Leaf)) { throw 'PowerPoint-origin presentation build failed.' }
    } elseif (-not (Test-Path -LiteralPath $output -PathType Leaf)) { throw "Existing generated candidate is missing: $output" }

    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $appWasRunning = $true }
    catch { $app = New-Object -ComObject PowerPoint.Application }
    $app.Visible = -1
    $process = Get-Process POWERPNT | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if (-not $process) { throw 'The real Microsoft PowerPoint window did not start.' }
    $applicationVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path).ProductVersion
    $presentation = $app.Presentations.Open($output, -1, 0, -1)
    if (-not $presentation) { throw 'PowerPoint did not return an opened presentation document.' }
    $templatePresentation = $app.Presentations.Open($template, -1, 0, 0)
    if (-not $templatePresentation) { throw 'PowerPoint did not open the canonical template for formatting comparison.' }

    $titles = [Collections.Generic.List[string]]::new()
    $staticCount = 0
    $pictures = 0
    $tables = 0
    $textShapeCount = 0
    foreach ($slide in $presentation.Slides) {
        $slideTitle = $null
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -eq 'Native template footer') { $staticCount++ }
            if ($shape.Type -eq 13 -or $shape.Type -eq 11 -or ($Milestone -eq 'M0020' -and $shape.Type -eq 28)) { $pictures++ }
            try { if ($shape.HasTable -eq -1) { $tables++ } } catch {}
            try {
                if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $textShapeCount++
                    $text = [string]$shape.TextFrame.TextRange.Text
                    if ($text.Contains('{{')) { throw 'Generated PowerPoint-origin output contains an unresolved control marker.' }
                    if ($shape.Name -like 'YADG title*') { $slideTitle = $text.Trim() }
                }
            } catch {
                if ($_.Exception.Message -like '*unresolved control marker*') { throw }
            }
        }
        if ([string]::IsNullOrWhiteSpace($slideTitle)) { throw 'An expected generated slide title is missing.' }
        $titles.Add($slideTitle)
    }
    $expectedTitles = if ($Milestone -eq 'M0019') { 'Narrative overview|Audience choices|Architecture map|Measures|YOLO recovery' } else { 'Narrative overview|Audience choices|Architecture map|Measures' }
    $expectedSlideCount = if ($Milestone -eq 'M0019') { 5 } else { 4 }
    if ($presentation.Slides.Count -ne $expectedSlideCount -or $staticCount -ne $expectedSlideCount -or $pictures -lt 1 -or $tables -ne 1 -or $textShapeCount -lt 4) { throw "Opened PowerPoint output did not retain expected native content (slides=$($presentation.Slides.Count)/$expectedSlideCount, footers=$staticCount, pictures=$pictures, tables=$tables, textShapes=$textShapeCount)." }
    if ($titles -join '|' -cne $expectedTitles) { throw "PowerPoint slide order differs from the source: $($titles -join ' | ')" }
    if ($Milestone -eq 'M0020') {
        $svgShapes = @($presentation.Slides | ForEach-Object { $_.Shapes } | Where-Object { $_.Type -eq 28 })
        if ($svgShapes.Count -ne 1 -or [string]$svgShapes[0].AlternativeText -cne 'Nested service diagram') { throw 'PowerPoint did not reopen exactly one SVG vector picture with the expected caption metadata.' }
    }
    $designCount = $presentation.Designs.Count
    $layoutCount = $presentation.Designs.Item(1).SlideMaster.CustomLayouts.Count
    if ($designCount -lt 1 -or $layoutCount -lt 1) { throw 'The opened output has no native design/master/layout relationship.' }

    $templateTitle = Get-M0018Paragraph (Get-M0018ShapeByName $templatePresentation 1 'YADG title content') '{{title}}'
    $outputTitle = Get-M0018Paragraph (Get-M0018ShapeByName $presentation 1 'YADG title content') 'Narrative overview'
    $templateBody = Get-M0018ShapeByName $templatePresentation 1 'YADG content region'
    $outputBody = Get-M0018ShapeByName $presentation 1 'YADG content region'
    $formatComparisons = [Collections.Generic.List[object]]::new()
    $formatComparisons.Add((Compare-M0018Formatting 'title' (Get-M0018TextFormatting $templateTitle) (Get-M0018TextFormatting $outputTitle)))
    $formatComparisons.Add((Compare-M0018Formatting 'paragraph' (Get-M0018TextFormatting (Get-M0018Paragraph $templateBody 'ordinary paragraph example')) (Get-M0018TextFormatting (Get-M0018Paragraph $outputBody 'This paragraph was authored in the shared M0018 compatibility source.'))))
    $templateHeading = Get-M0018Paragraph $templateBody 'content heading example'
    $generatedHeading = Get-M0018Paragraph $outputBody 'Design intent'
    if ([double]$templateHeading.Font.Size -ge [double]$templateTitle.Font.Size -or [double]$generatedHeading.Font.Size -ge [double]$outputTitle.Font.Size) { throw 'A body content heading must remain below the slide-title role in native font size.' }
    $formatComparisons.Add((Compare-M0018Formatting 'heading' (Get-M0018TextFormatting $templateHeading) (Get-M0018TextFormatting $generatedHeading)))
    $formatComparisons.Add((Compare-M0018Formatting 'nested-list' (Get-M0018TextFormatting (Get-M0018Paragraph $templateBody 'list formatting example')) (Get-M0018TextFormatting (Get-M0018Paragraph $outputBody 'A nested list remains editable'))))
    $templateTableHeader = Get-M0018TableCellFormatting $templatePresentation.Slides.Item(4) 1 2
    $outputTableHeader = Get-M0018TableCellFormatting $presentation.Slides.Item(4) 1 2
    $formatComparisons.Add((Compare-M0018Formatting 'table-header' $templateTableHeader $outputTableHeader))
    $templateTableBody = Get-M0018TableCellFormatting $templatePresentation.Slides.Item(4) 2 1
    $outputTableBody = Get-M0018TableCellFormatting $presentation.Slides.Item(4) 2 1
    $formatComparisons.Add((Compare-M0018Formatting 'table-body' $templateTableBody $outputTableBody))
    $templateTableGeometry = Get-M0018TableGeometry $templatePresentation.Slides.Item(4)
    $outputTableGeometry = Get-M0018TableGeometry $presentation.Slides.Item(4)
    if ($templateTableGeometry.left -ne $outputTableGeometry.left -or $templateTableGeometry.top -ne $outputTableGeometry.top -or $templateTableGeometry.width -ne $outputTableGeometry.width -or ($templateTableGeometry.columnWidths -join ',') -cne ($outputTableGeometry.columnWidths -join ',')) { throw 'Generated PowerPoint table changed the prototype-owned frame origin/width or column widths.' }
    if ($outputTableGeometry.rowHeights.Count -ne 3 -or $outputTableGeometry.rowHeights[0] -ne $templateTableGeometry.rowHeights[0] -or $outputTableGeometry.rowHeights[1] -ne $templateTableGeometry.rowHeights[1] -or $outputTableGeometry.rowHeights[2] -ne $templateTableGeometry.rowHeights[1]) { throw 'Generated PowerPoint table did not preserve the header/body row heights while duplicating the body row.' }
    if ($outputTableGeometry.left -lt 0 -or $outputTableGeometry.top -lt 0 -or $outputTableGeometry.left + $outputTableGeometry.width -gt $presentation.PageSetup.SlideWidth -or $outputTableGeometry.top + $outputTableGeometry.height -gt $presentation.PageSetup.SlideHeight) { throw 'Generated PowerPoint table frame extends outside the native slide bounds.' }
    $formattingValidation = [ordered]@{ result = 'passed'; source = 'Microsoft PowerPoint COM formatting comparison against the canonical PowerPoint-origin template'; comparisons = @($formatComparisons) }
    $slideRelationships = [Collections.Generic.List[object]]::new()
    foreach ($slide in $presentation.Slides) {
        $titleShape = $null
        $placeholderTypes = [Collections.Generic.List[int]]::new()
        foreach ($shape in $slide.Shapes) {
            try { $placeholderType = [int]$shape.PlaceholderFormat.Type; $placeholderTypes.Add($placeholderType); if ($shape.Name -like 'YADG title*') { $titleShape = $shape } } catch {}
        }
        if (-not $titleShape -or [int]$titleShape.PlaceholderFormat.Type -ne 1) { throw "PowerPoint output slide $($slide.SlideIndex) lost its native title placeholder role." }
        $prototypeName = [string]$titleShape.Name -replace '^YADG title ', ''
        $prototypeSlide = $null
        foreach ($candidateSlide in $templatePresentation.Slides) {
            foreach ($candidateShape in $candidateSlide.Shapes) { if ($candidateShape.Name -ceq "YADG title $prototypeName") { $prototypeSlide = $candidateSlide; break } }
            if ($prototypeSlide) { break }
        }
        if (-not $prototypeSlide) { throw "PowerPoint output slide $($slide.SlideIndex) cannot be tied to native prototype '$prototypeName'." }
        if ([string]$slide.CustomLayout.Name -cne [string]$prototypeSlide.CustomLayout.Name -or [string]$slide.Design.SlideMaster.Name -cne [string]$prototypeSlide.Design.SlideMaster.Name) { throw "PowerPoint output slide $($slide.SlideIndex) changed its native layout/master relationship." }
        $slideRelationships.Add([ordered]@{ slide=$slide.SlideIndex; prototype=$prototypeName; layoutName=[string]$slide.CustomLayout.Name; masterName=[string]$slide.Design.SlideMaster.Name; placeholderTypes=@($placeholderTypes) })
    }
    if ($Milestone -eq 'M0019') {
        $placeholderText = @($presentation.Slides | ForEach-Object { $_.Shapes } | Where-Object { $_.Name -eq 'YADG image region' } | ForEach-Object { $_.TextFrame.TextRange.Text }) -join "`n"
        if (-not $placeholderText.Contains('YADG image unavailable')) { throw 'M0019 PowerPoint output lost its visible YOLO image placeholder.' }
    }
    $renderRoot = Join-Path (Split-Path -Parent $evidence) 'renderings/powerpoint-origin'
    $templateRenderRoot = Join-Path $renderRoot 'prototype'
    $outputRenderRoot = Join-Path $renderRoot 'candidate'
    New-Item -ItemType Directory -Force -Path $templateRenderRoot, $outputRenderRoot | Out-Null
    $nativeRenders = [ordered]@{ application='Microsoft PowerPoint'; applicationVersion=$applicationVersion; prototype=@(); candidate=@() }
    foreach ($slide in $templatePresentation.Slides) {
        $path = Join-Path $templateRenderRoot "slide-$($slide.SlideIndex).png"
        $slide.Export($path, 'PNG', 1440, 810)
        $nativeRenders.prototype += [ordered]@{ slide=$slide.SlideIndex; path=[IO.Path]::GetRelativePath($repo, $path).Replace('\','/'); sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() }
    }
    foreach ($slide in $presentation.Slides) {
        $path = Join-Path $outputRenderRoot "slide-$($slide.SlideIndex).png"
        $slide.Export($path, 'PNG', 1440, 810)
        $nativeRenders.candidate += [ordered]@{ slide=$slide.SlideIndex; path=[IO.Path]::GetRelativePath($repo, $path).Replace('\','/'); sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() }
    }
    $inspect = [ordered]@{ slideTitles=@($titles); slideCount=$presentation.Slides.Count; slideRelationships=@($slideRelationships); staticFooterCount=$staticCount; editablePictureCount=$pictures; nativeTableCount=$tables; tableGeometry=[ordered]@{ prototype=$templateTableGeometry; candidate=$outputTableGeometry; candidateWithinSlideBounds=$true }; editableTextShapeCount=$textShapeCount; designCount=$designCount; nativeLayoutCount=$layoutCount; formattingValidation=$formattingValidation; nativeRenderings=$nativeRenders }

    $reviewCandidatePath = Join-Path $repo 'artifacts/review/evidence/M0018/review-sample.pptx'
    if ($Milestone -eq 'M0018' -and (Test-Path -LiteralPath $reviewCandidatePath -PathType Leaf)) {
        $reviewCandidatePresentation = $app.Presentations.Open($reviewCandidatePath, -1, 0, -1)
        if (-not $reviewCandidatePresentation -or $reviewCandidatePresentation.Slides.Count -ne 4) { throw 'The designated Impress-origin review candidate did not open as a four-slide presentation in PowerPoint.' }
        $reviewCandidateCheck = [ordered]@{
            path = [IO.Path]::GetRelativePath($repo, $reviewCandidatePath)
            sha256 = (Get-FileHash -LiteralPath $reviewCandidatePath -Algorithm SHA256).Hash.ToUpperInvariant()
            result = 'opened-without-repair'
            slideCount = $reviewCandidatePresentation.Slides.Count
        }
        $reviewCandidatePresentation.Close()
        $reviewCandidatePresentation = $null
    }
    $result = 'passed'
} catch {
    $failure = $_.Exception.Message
} finally {
    $presentationClosed = $true
    if ($reviewCandidatePresentation) { try { $reviewCandidatePresentation.Close() } catch { $presentationClosed = $false } }
    if ($presentation) { try { $presentation.Close() } catch { $presentationClosed = $false } }
    if ($templatePresentation) { try { $templatePresentation.Close() } catch { $presentationClosed = $false } }
    $applicationQuit = $null
    if ($app -and -not $appWasRunning) { $applicationQuit = $true; try { $app.Quit() } catch { $applicationQuit = $false } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $record = [ordered]@{
        milestone = $Milestone; gate = $Gate; originPath = 'PowerPoint-origin template -> YADG -> Microsoft PowerPoint'; result = $result; failure = $failure
        repositoryRevision = $revision; application = 'Microsoft PowerPoint'; applicationVersion = $applicationVersion
        platform = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
        templateOriginApplication = 'Microsoft PowerPoint'; templateOriginVersion = [string]$provenance.applicationVersion
        templateCreationMethod = [string]$provenance.creationMethod; templateCreatedAtUtc = [string]$provenance.createdAtUtc
        syntheticAndRedistributionSafe = [bool]$provenance.syntheticAndRedistributionSafe
        templatePath = [IO.Path]::GetRelativePath($repo, $template); templateSha256 = $templateHash
        sourcePath = [IO.Path]::GetRelativePath($repo, $source); sourceSha256 = $sourceHash
        configPath = if ($configFullPath) { [IO.Path]::GetRelativePath($repo, $configFullPath) } else { $null }; configSha256 = $configHash; yolo = [bool]$Yolo
        outputPath = [IO.Path]::GetRelativePath($repo, $output)
        outputSha256 = if (Test-Path -LiteralPath $output -PathType Leaf) { (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
        openResult = if ($result -eq 'passed') { 'opened-without-repair' } else { 'failed' }
        inspection = $inspect; formattingValidation = $formattingValidation; reviewCandidatePowerPointOpen = $reviewCandidateCheck
        presentationClosed = $presentationClosed; applicationOwned = (-not $appWasRunning); applicationQuit = $applicationQuit
    }
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $evidence -Encoding utf8
}
if ($result -ne 'passed') { throw "M0018 PowerPoint Tier 3 failed: $failure. Failure receipt: $evidence" }
Write-Output "M0018 PowerPoint-origin Tier 3 passed: $evidence"
