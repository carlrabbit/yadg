param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0019'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = [IO.Path]::GetFullPath($EvidenceRoot)
$tools = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
$bunVersion = [string]$tools.bun.version
$bun = Join-Path ([IO.Path]::GetTempPath()) "yadg-m0008-tools/bun-$bunVersion-windows-x64/bun.exe"
if (-not (Test-Path -LiteralPath $bun -PathType Leaf)) { & (Join-Path $PSScriptRoot 'test-m0008-tier3.ps1') -CacheRoot (Split-Path -Parent (Split-Path -Parent $bun)) }
if ((& $bun --version).Trim() -ne $bunVersion) { throw "M0019 PPTX Tier 3 requires pinned Bun $bunVersion." }

$sourceFixture = Join-Path $repo 'tests/fixtures/m0018/review-source.md'
$source = Join-Path $root 'review-source.md'
$config = Join-Path $root 'authoring.yaml'
$runner = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'test-m0019-run-mermaid.ps1')).Replace('\','/')
$content = Get-Content -LiteralPath $sourceFixture -Raw
$content = $content.Replace('This paragraph was authored in the shared M0018 compatibility source.','This paragraph was authored in the shared M0018 compatibility source. Review period: {{value:period}}.')
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
New-Item -ItemType Directory -Force -Path $root | Out-Null
Set-Content -LiteralPath $source -Encoding utf8 -Value $content
Set-Content -LiteralPath $config -Encoding utf8 -Value @('yadg:','  version: 1','values:','  period: "M0019 review"','markdown:','  codeInline: ignore','producers:','  mermaid:','    executable: powershell.exe','    arguments:','      - -NoProfile','      - -File',$('      - "' + $runner + '"'))

$globalPath = Join-Path $repo 'global.json'
if (Test-Path -LiteralPath $globalPath) { throw 'Refusing to replace an existing repository global.json.' }
Set-Content -LiteralPath $globalPath -Value '{"sdk":{"version":"10.0.403","rollForward":"disable","allowPrerelease":false}}' -Encoding ascii
$oldBun = $env:YADG_M0019_BUN
$oldBrowser = $env:PUPPETEER_EXECUTABLE_PATH
$powerPointExe = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/POWERPNT.EXE'
$powerPointFixture = Join-Path $repo 'tests/fixtures/m0018/powerpoint-origin.pptx'
$impressFixture = Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx'
$powerPointReceipt = Join-Path $root 'powerpoint-tier3-evidence.json'
$impressReceipt = Join-Path $root 'impress-tier3-evidence.json'
$revision = (& git -C $repo rev-parse HEAD).Trim()
try {
    $env:YADG_M0019_BUN = $bun
    if ([string]::IsNullOrWhiteSpace($env:PUPPETEER_EXECUTABLE_PATH)) {
        $browserCandidates = @(
            (Join-Path $env:ProgramFiles 'Google/Chrome/Application/chrome.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Microsoft/Edge/Application/msedge.exe')
        )
        $browser = $browserCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
        if ($browser) { $env:PUPPETEER_EXECUTABLE_PATH = $browser }
    }

    & (Join-Path $PSScriptRoot 'test-m0018-powerpoint-tier3.ps1') -TemplatePath $powerPointFixture -SourcePath $source -OutputPath (Join-Path $root 'deck-powerpoint.pptx') -EvidencePath $powerPointReceipt -ConfigPath $config -Yolo -Milestone M0019 -Gate VAL-05
    if (-not $?) { throw 'PowerPoint-origin -> YADG -> PowerPoint validation failed.' }

    & (Join-Path $PSScriptRoot 'test-m0018-tier3.ps1') -TemplatePath $impressFixture -SourcePath $source -CandidatePath (Join-Path $root 'deck-impress.pptx') -EvidencePath $impressReceipt -ConfigPath $config -Yolo -AutomatedReviewOnly -Milestone M0019 -Gate VAL-05
    if (-not $?) { throw 'Impress-origin -> YADG -> Impress validation failed.' }

    foreach ($receiptPath in @($powerPointReceipt, $impressReceipt)) {
        $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json -DateKind String
        if ([string]$receipt.result -cne 'passed' -or [string]$receipt.milestone -cne 'M0019' -or [string]$receipt.gate -cne 'VAL-05' -or [string]$receipt.repositoryRevision -cne $revision) { throw "Origin-specific receipt is not a current M0019 VAL-05 pass: $receiptPath" }
    }
    $powerPointRecord = Get-Content -LiteralPath $powerPointReceipt -Raw | ConvertFrom-Json -DateKind String
    $impressRecord = Get-Content -LiteralPath $impressReceipt -Raw | ConvertFrom-Json -DateKind String
    $record = [ordered]@{
        milestone='M0019'; gate='VAL-05'; result='passed'; repositoryRevision=$revision
        paths=@(
            [ordered]@{ id='EC-21a/EC-24a'; origin='Microsoft PowerPoint'; authoring='YADG'; inspection='Microsoft PowerPoint'; candidate='artifacts/review/evidence/M0019/deck-powerpoint.pptx'; candidateSha256=(Get-FileHash -LiteralPath (Join-Path $root 'deck-powerpoint.pptx') -Algorithm SHA256).Hash.ToUpperInvariant(); receipt='artifacts/review/evidence/M0019/powerpoint-tier3-evidence.json'; receiptSha256=(Get-FileHash -LiteralPath $powerPointReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); nativeRenderings=$powerPointRecord.inspection.nativeRenderings }
            [ordered]@{ id='EC-21b/EC-24b'; origin='LibreOffice Impress'; authoring='YADG'; inspection='LibreOffice Impress'; candidate='artifacts/review/evidence/M0019/deck-impress.pptx'; candidateSha256=(Get-FileHash -LiteralPath (Join-Path $root 'deck-impress.pptx') -Algorithm SHA256).Hash.ToUpperInvariant(); receipt='artifacts/review/evidence/M0019/impress-tier3-evidence.json'; receiptSha256=(Get-FileHash -LiteralPath $impressReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); nativeRenderings=$impressRecord.inspection.nativeRenderings }
        )
        visualFitDisposition='executor-reviewed-no-visible-oversized-content-or-overflow; genuine human review remains pending'
        humanReviewId='HR-M0019-01'; humanReviewStatus='pending'
    }
    $record | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath (Join-Path $root 'pptx-tier3-evidence.json') -Encoding utf8
    Write-Output 'M0019 independent PowerPoint-origin/PowerPoint and Impress-origin/Impress Tier 3 validations passed.'
} finally {
    Remove-Item -LiteralPath $globalPath -Force -ErrorAction SilentlyContinue
    if ($null -eq $oldBun) { Remove-Item Env:YADG_M0019_BUN -ErrorAction SilentlyContinue } else { $env:YADG_M0019_BUN = $oldBun }
    if ($null -eq $oldBrowser) { Remove-Item Env:PUPPETEER_EXECUTABLE_PATH -ErrorAction SilentlyContinue } else { $env:PUPPETEER_EXECUTABLE_PATH = $oldBrowser }
}
