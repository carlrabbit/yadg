$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$evidenceRoot = Join-Path $repo 'artifacts/review/evidence/M0018'
$candidate = Join-Path $evidenceRoot 'review-sample.pptx'
$source = Join-Path $repo 'tests/fixtures/m0018/review-source.md'
$image = Join-Path $repo 'tests/fixtures/m0018/images/native-office-sample.png'
$impressTemplate = Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx'
$powerPointReceiptPath = Join-Path $evidenceRoot 'powerpoint-tier3-evidence.json'
$impressReceiptPath = Join-Path $evidenceRoot 'impress-tier3-evidence.json'
$paths = @($candidate, $source, $image, $impressTemplate, $powerPointReceiptPath, $impressReceiptPath)
foreach ($path in $paths) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Cannot create M0018 review manifest; required evidence is missing: $path" } }
$revision = (& git -C $repo rev-parse HEAD).Trim()
$powerPointReceipt = Get-Content -LiteralPath $powerPointReceiptPath -Raw | ConvertFrom-Json -DateKind String
$impressReceipt = Get-Content -LiteralPath $impressReceiptPath -Raw | ConvertFrom-Json -DateKind String
if ([string]$powerPointReceipt.repositoryRevision -cne $revision -or [string]$powerPointReceipt.result -cne 'passed' -or [string]$powerPointReceipt.gate -cne 'VAL-06') { throw 'Current PowerPoint-origin Tier 3 receipt is absent, stale or failed.' }
if ([string]$impressReceipt.repositoryRevision -cne $revision -or [string]$impressReceipt.result -cne 'passed' -or [string]$impressReceipt.gate -cne 'VAL-03') { throw 'Current Impress-origin Tier 3 receipt is absent, stale or failed.' }

function Get-Hash([string] $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() }
function Get-Relative([string] $path) { [IO.Path]::GetRelativePath($repo, $path).Replace('\', '/') }

$manifest = [ordered]@{
    schemaVersion = 1
    milestone = 'M0018'
    repositoryRevision = $revision
    candidatePath = (Get-Relative $candidate)
    candidateSha256 = (Get-Hash $candidate)
    sourcePath = (Get-Relative $source)
    sourceSha256 = (Get-Hash $source)
    imagePath = (Get-Relative $image)
    imageSha256 = (Get-Hash $image)
    impressTemplatePath = (Get-Relative $impressTemplate)
    impressTemplateSha256 = (Get-Hash $impressTemplate)
    powerPointTier3 = [ordered]@{
        path = (Get-Relative $powerPointReceiptPath)
        sha256 = (Get-Hash $powerPointReceiptPath)
    }
    impressTier3 = [ordered]@{
        path = (Get-Relative $impressReceiptPath)
        sha256 = (Get-Hash $impressReceiptPath)
    }
    officeVersions = [ordered]@{
        powerPointTier3 = [string]$powerPointReceipt.applicationVersion
        impressTier3 = [string]$impressReceipt.applicationVersion
    }
    humanReviewId = 'HR-M0018-01'
}
$manifestPath = Join-Path $evidenceRoot 'review-manifest.json'
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding utf8
& (Join-Path $PSScriptRoot 'verify-m0018-review-manifest.ps1')
Write-Output "M0018 review manifest created and verified: $manifestPath"
