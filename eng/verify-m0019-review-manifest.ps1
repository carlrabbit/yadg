param([switch] $AllowPendingHumanReview)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$evidenceRoot = Join-Path $repo 'artifacts/review/evidence/M0019'
$manifestPath = Join-Path $evidenceRoot 'review-manifest.json'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw 'M0019 review manifest is missing.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -DateKind String
$revision = (& git -C $repo rev-parse HEAD).Trim()
if ([string]$manifest.milestone -cne 'M0019' -or [int]$manifest.schemaVersion -ne 1 -or [string]$manifest.humanReviewId -cne 'HR-M0019-01' -or [string]$manifest.repositoryRevision -cne $revision) { throw 'M0019 review manifest does not identify the current milestone, review request and repository revision.' }

$expected = @('artifacts/review/evidence/M0019/finalized-word.docx','artifacts/review/evidence/M0019/deck-impress.pptx')
$candidates = @($manifest.candidates)
$candidatePaths = @($candidates | ForEach-Object { [string]$_.path } | Sort-Object)
$expectedPaths = @($expected | Sort-Object)
if ($candidates.Count -ne 2 -or $candidatePaths[0] -cne $expectedPaths[0] -or $candidatePaths[1] -cne $expectedPaths[1]) { throw 'M0019 manifest must bind exactly the designated finalized DOCX and PPTX candidates.' }
foreach ($candidate in $candidates) {
    $path = Join-Path $repo ([string]$candidate.path)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "M0019 candidate is missing: $($candidate.path)" }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$candidate.sha256) { throw "M0019 candidate hash is stale: $($candidate.path)" }
}

foreach ($name in @('wordSource','wordTemplate','presentationSource','directConfig','impressTemplate')) {
    $input = $manifest.inputs.$name
    $inputPath = Join-Path $repo ([string]$input.path)
    if ($null -eq $input -or -not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "M0019 manifest input '$name' is missing." }
    $expectedHash = if ($name -eq 'wordSource') { [string]$input.sourceDefinitionSha256 } else { [string]$input.sha256 }
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne $expectedHash) { throw "M0019 manifest input '$name' hash is stale." }
}
$packagePath = Join-Path $repo ([string]$manifest.tool.packagePath)
if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf) -or (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$manifest.tool.packageSha256) { throw 'M0019 installed-tool package is missing or stale.' }
foreach ($application in @('word','writer','impress','powerPoint')) { if ([string]::IsNullOrWhiteSpace([string]$manifest.applications.$application)) { throw "M0019 manifest is missing the $application application version." } }

foreach ($gate in @('VAL-03','VAL-04','VAL-05')) {
    $receipt = $manifest.tier3Receipts.$($gate.Replace('-',''))
    if ($null -eq $receipt -or [string]::IsNullOrWhiteSpace([string]$receipt.path)) { throw "M0019 review manifest has no $gate receipt." }
    $receiptPath = Join-Path $repo ([string]$receipt.path)
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf) -or (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$receipt.sha256) { throw "M0019 $gate receipt is missing or its hash is stale." }
    $record = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json -DateKind String
    if ([string]$record.milestone -cne 'M0019' -or [string]$record.gate -cne $gate -or [string]$record.result -cne 'passed' -or [string]$record.repositoryRevision -cne $revision) { throw "M0019 $gate receipt is not a current successful receipt." }
}

$reviewRecordPath = Join-Path $repo '.review/records/HR-M0019-01.md'
if (-not (Test-Path -LiteralPath $reviewRecordPath -PathType Leaf)) {
    if (-not $AllowPendingHumanReview) { throw 'HR-M0019-01 is still pending genuine human review.' }
    Write-Output 'M0019 candidates and automated receipts are current; HR-M0019-01 remains pending human review.'
    exit 0
}
$review = Get-Content -LiteralPath $reviewRecordPath -Raw
$manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToUpperInvariant()
foreach ($line in @(
    'milestone: M0019','reviewId: HR-M0019-01','reviewClass: artifact-quality','decision: approved','status: approved',
    "repositoryRevision: $revision",'candidateManifestPath: artifacts/review/evidence/M0019/review-manifest.json',
    "candidateManifestSha256: $manifestHash",'wordArtifact: artifacts/review/evidence/M0019/finalized-word.docx',
    "wordArtifactSha256: $($candidates | Where-Object path -eq 'artifacts/review/evidence/M0019/finalized-word.docx' | Select-Object -ExpandProperty sha256)",
    'deckArtifact: artifacts/review/evidence/M0019/deck-impress.pptx',
    "deckArtifactSha256: $($candidates | Where-Object path -eq 'artifacts/review/evidence/M0019/deck-impress.pptx' | Select-Object -ExpandProperty sha256)",
    "evidence: artifacts/review/evidence/M0019/review-manifest.json"
)) { if ($review -notmatch "(?m)^$([regex]::Escape($line))\s*$") { throw "M0019 human decision is not bound to '$line'." } }
if ($review -notmatch '(?m)^reviewer:\s*\S.+$' -or $review -notmatch '(?m)^reviewedAtUtc:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z\s*$') { throw 'M0019 approval must name the actual reviewer and UTC review time.' }
Write-Output 'HR-M0019-01: approved and bound to both current artifacts.'
