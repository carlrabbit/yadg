param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0019'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bunVersion = [string](Get-Content (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json).bun.version
$bun = Join-Path ([IO.Path]::GetTempPath()) "yadg-m0008-tools/bun-$bunVersion-windows-x64/bun.exe"
if (-not (Test-Path -LiteralPath $bun -PathType Leaf)) { & (Join-Path $PSScriptRoot 'test-m0008-tier3.ps1') -CacheRoot (Split-Path -Parent (Split-Path -Parent $bun)) }
if ((& $bun --version).Trim() -ne $bunVersion) { throw "M0019 PPTX Tier 3 requires pinned Bun $bunVersion." }
$tools = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
$mermaidPackage = [string]$tools.mermaidCli.package
$mermaidVersion = [string]$tools.mermaidCli.version
$sourceFixture = Join-Path $repo 'tests/fixtures/m0018/review-source.md'
$source = Join-Path $EvidenceRoot 'review-source.md'
$config = Join-Path $EvidenceRoot 'authoring.yaml'
$runner = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'test-m0019-run-mermaid.ps1')).Replace('\','/')
$runnerArgument = '      - "' + $runner + '"'
$content = Get-Content -LiteralPath $sourceFixture -Raw
$content = $content.Replace('This paragraph was authored in the shared M0018 compatibility source.','This paragraph was authored in the shared M0018 compatibility source. Review period: {{value:period}}. ~~legacy~~ [support link](https://example.test/support).')
$mermaid = @'
```mermaid {#generated-flow caption="Generated flow"}
flowchart LR
  A --> B
```
'@
$recovery = @'
# YOLO recovery {#text-image}

## Explanation {#text}

This slide demonstrates explicit recovery.

## Missing image {#image}

![Missing image](images/not-present.png)
'@
$content = [regex]::Replace($content,'!\[[^\]]*\]\([^\)]+\)',$mermaid.Trim())
$content += "`n`n" + $recovery.Trim() + "`n"
New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
Set-Content -LiteralPath $source -Encoding utf8 -Value $content
Set-Content -LiteralPath $config -Encoding utf8 -Value @('yadg:','  version: 1','values:','  period: "M0019 review"','markdown:','  codeInline: ignore','producers:','  mermaid:','    executable: powershell.exe','    arguments:','      - -NoProfile','      - -File',$runnerArgument)

$globalPath = Join-Path $repo 'global.json'
if (Test-Path -LiteralPath $globalPath) { throw 'Refusing to replace an existing repository global.json.' }
Set-Content -LiteralPath $globalPath -Value '{"sdk":{"version":"10.0.403","rollForward":"disable","allowPrerelease":false}}' -Encoding ascii
$oldBun = $env:YADG_M0019_BUN
$candidate = Join-Path $EvidenceRoot 'deck-impress.pptx'
$impressReceipt = Join-Path $EvidenceRoot 'impress-tier3-evidence.json'
try {
    $env:YADG_M0019_BUN = $bun
    & (Join-Path $PSScriptRoot 'test-m0018-tier3.ps1') -SourcePath $source -CandidatePath $candidate -EvidencePath $impressReceipt -ConfigPath $config -Yolo -AutomatedReviewOnly
    if ($LASTEXITCODE -ne 0) { throw 'Real Impress did not open the M0019-generated review deck successfully.' }

    $powerPointExe = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/POWERPNT.EXE'
    if (-not (Test-Path -LiteralPath $powerPointExe -PathType Leaf)) { throw 'M0019 Tier 3 requires the real Microsoft PowerPoint desktop application.' }
    $powerPointVersion = (Get-Item -LiteralPath $powerPointExe).VersionInfo.FileVersion
    $appWasRunning = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $null
    $result = 'failed'
    $failure = $null
    try {
        $presentation = $app.Presentations.Open($candidate, -1, 0, 0)
        if (-not $presentation -or $presentation.Slides.Count -ne 5) { throw 'PowerPoint did not reopen the exact five-slide M0019 candidate.' }
        $pictures = 0; $text = [Collections.Generic.List[string]]::new(); $tables = 0
        foreach ($slide in $presentation.Slides) {
            foreach ($shape in $slide.Shapes) {
                if ([int]$shape.Type -eq 13 -or [int]$shape.Type -eq 11) { $pictures++ }
                if ([int]$shape.HasTable -eq -1) { $tables++ }
                if ([int]$shape.HasTextFrame -eq -1 -and [int]$shape.TextFrame.HasText -eq -1) { $text.Add([string]$shape.TextFrame.TextRange.Text) }
            }
        }
        if ($pictures -lt 1 -or $tables -lt 1 -or -not ($text -join "`n").Contains('YADG image unavailable') -or ($text -join "`n").Contains('{{')) { throw 'PowerPoint inspection did not find native pictures, table structure, a visible YOLO placeholder, and cleared template/source tags.' }
        $result = 'passed'
    } catch { $failure = $_.Exception.Message } finally {
        if ($presentation) { try { $presentation.Close() } catch { $result = 'failed'; $failure = 'PowerPoint presentation cleanup failed.' } }
        if (-not $appWasRunning -and $app) { try { $app.Quit() } catch { $result = 'failed'; $failure = 'PowerPoint process cleanup failed.' } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    if ($result -ne 'passed') { throw "PowerPoint Tier 3 failed: $failure" }
    $impress = Get-Content -LiteralPath $impressReceipt -Raw | ConvertFrom-Json -DateKind String
    $record = [ordered]@{
        milestone='M0019'; gate='VAL-05'; result='passed'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim()
        templateOrigin='LibreOffice Impress'; templatePath='tests/fixtures/m0018/impress-origin.pptx'; templateSha256=(Get-FileHash -LiteralPath (Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx') -Algorithm SHA256).Hash.ToUpperInvariant()
        sourcePath='artifacts/review/evidence/M0019/review-source.md'; sourceSha256=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToUpperInvariant()
        configPath='artifacts/review/evidence/M0019/authoring.yaml'; configSha256=(Get-FileHash -LiteralPath $config -Algorithm SHA256).Hash.ToUpperInvariant()
        candidatePath='artifacts/review/evidence/M0019/deck-impress.pptx'; candidateSha256=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToUpperInvariant()
        impressVersion=[string]$impress.applicationVersion; impressReceiptSha256=(Get-FileHash -LiteralPath $impressReceipt -Algorithm SHA256).Hash.ToUpperInvariant()
        powerPointVersion=$powerPointVersion; powerPointOpen='opened-without-repair'; slideCount=5; nativePictureCount=$pictures; nativeTableCount=$tables; yoloPlaceholderVisible=$true
        mermaidPackage="$mermaidPackage@$mermaidVersion"; profileCleaned=[bool]$impress.profileCleaned
    }
    $record | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'pptx-tier3-evidence.json') -Encoding utf8
    Write-Output 'M0019 real Impress and PowerPoint Tier 3 reopen checks passed.'
} finally {
    Remove-Item -LiteralPath $globalPath -Force -ErrorAction SilentlyContinue
    if ($null -eq $oldBun) { Remove-Item Env:YADG_M0019_BUN -ErrorAction SilentlyContinue } else { $env:YADG_M0019_BUN = $oldBun }
}
