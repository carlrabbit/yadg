param(
    [string] $TemplatePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/powerpoint-origin.pptx'),
    [string] $SourcePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/review-source.md'),
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/runs/powerpoint-origin-generated.pptx'),
    [string] $EvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/powerpoint-tier3-evidence.json')
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
$templateHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant()
$app = $null
$presentation = $null
$reviewCandidatePresentation = $null
$appWasRunning = $false
$failure = $null
$applicationVersion = $null
$inspect = $null
$result = 'failed'
try {
    $arguments = @($cli, 'presentation', 'check', '--template', $template, '--source', $source)
    & dotnet @arguments
    if ($LASTEXITCODE -ne 0) { throw 'PowerPoint-origin fixture failed presentation preflight.' }
    & dotnet $cli presentation build --template $template --source $source --output $output
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $output -PathType Leaf)) { throw 'PowerPoint-origin presentation build failed.' }

    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $appWasRunning = $true }
    catch { $app = New-Object -ComObject PowerPoint.Application }
    $app.Visible = -1
    $process = Get-Process POWERPNT | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    if (-not $process) { throw 'The real Microsoft PowerPoint window did not start.' }
    $applicationVersion = [Diagnostics.FileVersionInfo]::GetVersionInfo($process.Path).ProductVersion
    $presentation = $app.Presentations.Open($output, -1, 0, -1)
    if (-not $presentation) { throw 'PowerPoint did not return an opened presentation document.' }

    $titles = [Collections.Generic.List[string]]::new()
    $staticCount = 0
    $pictures = 0
    $tables = 0
    $textShapeCount = 0
    foreach ($slide in $presentation.Slides) {
        $slideTitle = $null
        foreach ($shape in $slide.Shapes) {
            if ($shape.Name -eq 'Native template footer') { $staticCount++ }
            if ($shape.Type -eq 13 -or $shape.Type -eq 11) { $pictures++ }
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
    if ($presentation.Slides.Count -ne 4 -or $staticCount -ne 4 -or $pictures -lt 1 -or $tables -ne 1 -or $textShapeCount -lt 4) { throw 'Opened PowerPoint output did not retain the four native slides, footers, text, picture and table.' }
    if ($titles -join '|' -cne 'Narrative overview|Audience choices|Architecture map|Measures') { throw "PowerPoint slide order differs from the source: $($titles -join ' | ')" }
    $designCount = $presentation.Designs.Count
    $layoutCount = $presentation.Designs.Item(1).SlideMaster.CustomLayouts.Count
    if ($designCount -lt 1 -or $layoutCount -lt 1) { throw 'The opened output has no native design/master/layout relationship.' }
    $inspect = [ordered]@{ slideTitles=@($titles); slideCount=$presentation.Slides.Count; staticFooterCount=$staticCount; editablePictureCount=$pictures; nativeTableCount=$tables; editableTextShapeCount=$textShapeCount; designCount=$designCount; nativeLayoutCount=$layoutCount }

    $reviewCandidatePath = Join-Path $repo 'artifacts/review/evidence/M0018/review-sample.pptx'
    if (Test-Path -LiteralPath $reviewCandidatePath -PathType Leaf) {
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
    $applicationQuit = $null
    if ($app -and -not $appWasRunning) { $applicationQuit = $true; try { $app.Quit() } catch { $applicationQuit = $false } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $record = [ordered]@{
        milestone = 'M0018'; gate = 'VAL-06'; result = $result; failure = $failure
        repositoryRevision = $revision; application = 'Microsoft PowerPoint'; applicationVersion = $applicationVersion
        platform = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
        templateOriginApplication = 'Microsoft PowerPoint'; templateOriginVersion = [string]$provenance.applicationVersion
        templateCreationMethod = [string]$provenance.creationMethod; templateCreatedAtUtc = [string]$provenance.createdAtUtc
        syntheticAndRedistributionSafe = [bool]$provenance.syntheticAndRedistributionSafe
        templatePath = [IO.Path]::GetRelativePath($repo, $template); templateSha256 = $templateHash
        sourcePath = [IO.Path]::GetRelativePath($repo, $source); sourceSha256 = $sourceHash
        outputPath = [IO.Path]::GetRelativePath($repo, $output)
        outputSha256 = if (Test-Path -LiteralPath $output -PathType Leaf) { (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
        openResult = if ($result -eq 'passed') { 'opened-without-repair' } else { 'failed' }
        inspection = $inspect; reviewCandidatePowerPointOpen = $reviewCandidateCheck
        presentationClosed = $presentationClosed; applicationOwned = (-not $appWasRunning); applicationQuit = $applicationQuit
    }
    $record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $evidence -Encoding utf8
}
if ($result -ne 'passed') { throw "M0018 PowerPoint Tier 3 failed: $failure. Failure receipt: $evidence" }
Write-Output "M0018 PowerPoint-origin Tier 3 passed: $evidence"
