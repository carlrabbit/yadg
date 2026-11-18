param([switch] $AllowPendingHumanReview)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = Join-Path $repo 'artifacts/review/evidence/M0020'
$manifestPath = Join-Path $root 'review-manifest.json'
function Hash([string] $path) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "M0020 review evidence is missing: $path" }; (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() }
function ResolveRepo([string] $path) { Join-Path $repo $path }
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw 'M0020 review manifest is missing.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -DateKind String
$revision = (& git -C $repo rev-parse HEAD).Trim()
if ($manifest.milestone -cne 'M0020' -or $manifest.humanReviewId -cne 'HR-M0020-01' -or [int]$manifest.candidateCount -ne 2 -or $manifest.repositoryRevision -cne $revision) { throw 'M0020 review manifest must bind two candidates, HR-M0020-01 and the current source revision.' }
if (@($manifest.candidates).Count -ne 2 -or @($manifest.candidates | Select-Object -ExpandProperty originApplication -Unique).Count -ne 2) { throw 'M0020 manifest must contain exactly one PowerPoint and one Impress candidate.' }
foreach ($candidate in $manifest.candidates) {
    if ($candidate.format -cne 'PPTX' -or $candidate.nativeApplication -cne $candidate.originApplication) { throw 'M0020 candidates must be origin-matched PPTX artifacts.' }
    if ((Hash (ResolveRepo $candidate.path)) -cne $candidate.sha256) { throw "M0020 candidate is stale: $($candidate.path)" }
    if ((Hash (ResolveRepo $candidate.fixture.path)) -cne $candidate.fixture.sha256 -or (Hash (ResolveRepo $candidate.fixture.provenancePath)) -cne $candidate.fixture.provenanceSha256) { throw "M0020 origin fixture/provenance is stale: $($candidate.fixture.path)" }
    foreach ($receipt in @($candidate.tier3Receipt,$candidate.tier4Receipt)) { if ((Hash (ResolveRepo $receipt.path)) -cne $receipt.sha256) { throw "M0020 candidate receipt is stale: $($receipt.path)" } }
    if ([string]::IsNullOrWhiteSpace([string]$candidate.applicationVersion) -or [int]$candidate.svg.svgPartCount -lt 1 -or [int]$candidate.svg.svgPictureReferences -lt 1) { throw "M0020 candidate lacks native runtime or SVG-picture evidence: $($candidate.path)" }
    if (@($candidate.nativeRenderings).Count -eq 0) { throw "M0020 candidate has no slide-by-slide native renderings: $($candidate.path)" }
    foreach ($rendering in $candidate.nativeRenderings) { if ((Hash (ResolveRepo $rendering.path)) -cne $rendering.sha256) { throw "M0020 native rendering is missing or stale: $($rendering.path)" } }
}
foreach ($receipt in @($manifest.tier3Receipt,$manifest.tier4Receipt)) { if ((Hash (ResolveRepo $receipt.path)) -cne $receipt.sha256) { throw "M0020 gate receipt is stale: $($receipt.path)" } }
$tier3 = Get-Content -LiteralPath (ResolveRepo $manifest.tier3Receipt.path) -Raw | ConvertFrom-Json -DateKind String
$tier4 = Get-Content -LiteralPath (ResolveRepo $manifest.tier4Receipt.path) -Raw | ConvertFrom-Json -DateKind String
if ($tier3.result -cne 'passed' -or $tier3.gate -cne 'VAL-03' -or $tier3.repositoryRevision -cne $revision -or $tier4.result -cne 'passed' -or $tier4.gate -cne 'VAL-04' -or $tier4.repositoryRevision -cne $revision) { throw 'M0020 review requires current successful origin-matched VAL-03 and installed-consumer VAL-04 receipts.' }
if ((Hash (ResolveRepo $manifest.package.path)) -cne $manifest.package.sha256 -or $tier4.packageSha256 -cne $manifest.package.sha256) { throw 'M0020 manifest does not bind the exact installed consumer package.' }
$recordPath = Join-Path $repo '.review/records/HR-M0020-01.md'
if (-not (Test-Path -LiteralPath $recordPath -PathType Leaf)) {
    if (-not $AllowPendingHumanReview) { throw 'HR-M0020-01 is pending genuine human review.' }
    Write-Output 'M0020 candidate and evidence manifest is current; HR-M0020-01 remains pending human approval.'
    exit 0
}
$record = Get-Content -LiteralPath $recordPath -Raw
$manifestHash = Hash $manifestPath
$required = @('milestone: M0020','reviewId: HR-M0020-01','reviewClass: artifact-quality','decision: approved','status: approved',"repositoryRevision: $revision",'candidateManifestPath: artifacts/review/evidence/M0020/review-manifest.json',"candidateManifestSha256: $manifestHash",'powerPointArtifact: artifacts/review/evidence/M0020/deck-powerpoint-svg.pptx','impressArtifact: artifacts/review/evidence/M0020/deck-impress-svg.pptx','tier4Receipt: artifacts/review/evidence/M0020/tier4-evidence.json',"tier4ReceiptSha256: $((Hash (ResolveRepo $manifest.tier4Receipt.path)))",'evidence: artifacts/review/evidence/M0020/review-manifest.json')
foreach ($line in $required) { if ($record -notmatch "(?m)^$([regex]::Escape($line))\s*$") { throw "HR-M0020-01 is not bound to '$line'." } }
foreach ($candidate in $manifest.candidates) {
    $label = if ($candidate.originApplication -eq 'Microsoft PowerPoint') { 'powerPoint' } else { 'impress' }
    $lines = @("${label}ArtifactSha256: $($candidate.sha256)","${label}Tier3Receipt: $($candidate.tier3Receipt.path)","${label}Tier3ReceiptSha256: $($candidate.tier3Receipt.sha256)")
    foreach ($line in $lines) { if ($record -notmatch "(?m)^$([regex]::Escape($line))\s*$") { throw "HR-M0020-01 is not bound to '$line'." } }
}
if ($record -notmatch '(?m)^reviewer:\s*\S.+$' -or $record -notmatch '(?m)^reviewedAtUtc:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z\s*$') { throw 'HR-M0020-01 requires the actual reviewer and UTC review timestamp.' }
Write-Output 'HR-M0020-01 is approved and bound to both current vector candidates and evidence.'
