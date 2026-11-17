param([switch] $AllowPendingHumanReview)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$evidenceRoot = Join-Path $repo 'artifacts/review/evidence/M0019'
$manifestPath = Join-Path $evidenceRoot 'review-manifest.json'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw 'M0019 review manifest is missing.' }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -DateKind String
$revision = (& git -C $repo rev-parse HEAD).Trim()
if ([string]$manifest.milestone -cne 'M0019' -or [int]$manifest.schemaVersion -ne 2 -or [int]$manifest.candidateCount -ne 3 -or [string]$manifest.humanReviewId -cne 'HR-M0019-01' -or [string]$manifest.repositoryRevision -cne $revision) { throw 'M0019 manifest must bind three candidates, the current milestone/revision and HR-M0019-01.' }

$expected = @('artifacts/review/evidence/M0019/deck-impress.pptx','artifacts/review/evidence/M0019/deck-powerpoint.pptx','artifacts/review/evidence/M0019/finalized-word.docx')
$candidates = @($manifest.candidates)
$candidatePaths = @($candidates | ForEach-Object { [string]$_.path } | Sort-Object)
$expectedPaths = @($expected | Sort-Object)
if ($candidates.Count -ne 3 -or [string]::Join('|', $candidatePaths) -cne [string]::Join('|', $expectedPaths)) { throw 'M0019 manifest must bind exactly the designated Word, PowerPoint-origin PPTX, and Impress-origin PPTX artifacts.' }
foreach ($candidate in $candidates) {
    $path = Join-Path $repo ([string]$candidate.path)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "M0019 candidate is missing: $($candidate.path)" }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$candidate.sha256) { throw "M0019 candidate hash is stale: $($candidate.path)" }
}

foreach ($name in @('wordSource','wordTemplate','presentationSource','directConfig','powerpointTemplate','impressTemplate')) {
    $input = $manifest.inputs.$name
    $inputPath = Join-Path $repo ([string]$input.path)
    if ($null -eq $input -or -not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw "M0019 manifest input '$name' is missing." }
    $expectedHash = if ($name -eq 'wordSource') { [string]$input.sourceDefinitionSha256 } else { [string]$input.sha256 }
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne $expectedHash) { throw "M0019 manifest input '$name' hash is stale." }
    if ($name -in @('powerpointTemplate','impressTemplate')) {
        $provenancePath = Join-Path $repo ([string]$input.provenancePath)
        if (-not (Test-Path -LiteralPath $provenancePath -PathType Leaf) -or (Get-FileHash -LiteralPath $provenancePath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$input.provenanceSha256) { throw "M0019 $name provenance is missing or stale." }
    }
}
$packagePath = Join-Path $repo ([string]$manifest.tool.packagePath)
if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf) -or (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$manifest.tool.packageSha256) { throw 'M0019 installed-tool package is missing or stale.' }
foreach ($application in @('word','writer','impress','powerPoint')) { if ([string]::IsNullOrWhiteSpace([string]$manifest.applications.$application)) { throw "M0019 manifest is missing the $application application version." } }

foreach ($gate in @('VAL03','VAL04','VAL05','VAL05PowerPoint','VAL05Impress')) {
    $receipt = $manifest.tier3Receipts.$gate
    if ($null -eq $receipt -or [string]::IsNullOrWhiteSpace([string]$receipt.path)) { throw "M0019 review manifest has no receipt $gate." }
    $receiptPath = Join-Path $repo ([string]$receipt.path)
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf) -or (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$receipt.sha256) { throw "M0019 receipt $gate is missing or its hash is stale." }
    $record = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json -DateKind String
    $expectedGate = if ($gate -like 'VAL05*') { 'VAL-05' } else { $gate.Insert(3,'-') }
    if ([string]$record.milestone -cne 'M0019' -or [string]$record.gate -cne $expectedGate -or [string]$record.result -cne 'passed' -or [string]$record.repositoryRevision -cne $revision) { throw "M0019 $gate receipt is not a current successful receipt." }
    if ($gate -in @('VAL05PowerPoint','VAL05Impress')) {
        if ([string]$record.sourceSha256 -cne [string]$manifest.inputs.presentationSource.sha256 -or [string]$record.configSha256 -cne [string]$manifest.inputs.directConfig.sha256) { throw "M0019 $gate receipt does not use the manifest-bound presentation source and config." }
        $expectedTemplate = if ($gate -eq 'VAL05PowerPoint') { $manifest.inputs.powerpointTemplate } else { $manifest.inputs.impressTemplate }
        if ([string]$record.templateSha256 -cne [string]$expectedTemplate.sha256) { throw "M0019 $gate receipt does not use its manifest-bound native-origin template." }
    }
}

foreach ($candidate in @($candidates | Where-Object { $_.format -eq 'PPTX' })) {
    $receiptPath = Join-Path $repo ([string]$candidate.nativeTier3Receipt.path)
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf) -or (Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$candidate.nativeTier3Receipt.sha256) { throw "M0019 candidate's native Tier 3 evidence is missing or stale: $($candidate.path)" }
    foreach ($render in @($candidate.nativeRenderings)) {
        $renderPath = Join-Path $repo ([string]$render.path)
        if (-not (Test-Path -LiteralPath $renderPath -PathType Leaf) -or (Get-FileHash -LiteralPath $renderPath -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$render.sha256) { throw "M0019 native rendering is missing or stale: $($render.path)" }
    }
    $expectedPrototypeSlides = 4
    $expectedCandidateSlides = 5
    $prototypeSlides = @($candidate.nativeRenderings | Where-Object { [string]$_.kind -in @('prototype-slide','prototypeSlideImages') } | ForEach-Object { [int]$_.slide } | Sort-Object -Unique)
    $candidateSlides = @($candidate.nativeRenderings | Where-Object { [string]$_.kind -in @('candidate-slide','candidateSlideImages') } | ForEach-Object { [int]$_.slide } | Sort-Object -Unique)
    if ($prototypeSlides.Count -ne $expectedPrototypeSlides -or $candidateSlides.Count -ne $expectedCandidateSlides) { throw "M0019 $($candidate.originApplication) review evidence must contain prototype and candidate renderings for every slide." }
    foreach ($index in 1..$expectedPrototypeSlides) { if ($prototypeSlides -notcontains $index) { throw "M0019 $($candidate.originApplication) prototype slide $index has no native rendering." } }
    foreach ($index in 1..$expectedCandidateSlides) { if ($candidateSlides -notcontains $index) { throw "M0019 $($candidate.originApplication) candidate slide $index has no native rendering." } }
}

$reviewRecordPath = Join-Path $repo '.review/records/HR-M0019-01.md'
if (-not (Test-Path -LiteralPath $reviewRecordPath -PathType Leaf)) {
    if (-not $AllowPendingHumanReview) { throw 'HR-M0019-01 is still pending genuine human review of all three artifacts.' }
    Write-Output 'M0019 candidates and origin-specific render evidence are current; HR-M0019-01 remains pending human review.'
    exit 0
}
$review = Get-Content -LiteralPath $reviewRecordPath -Raw
$manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToUpperInvariant()
$wordCandidate = $candidates | Where-Object path -eq 'artifacts/review/evidence/M0019/finalized-word.docx'
$powerPointCandidate = $candidates | Where-Object path -eq 'artifacts/review/evidence/M0019/deck-powerpoint.pptx'
$impressCandidate = $candidates | Where-Object path -eq 'artifacts/review/evidence/M0019/deck-impress.pptx'
foreach ($line in @(
    'milestone: M0019','reviewId: HR-M0019-01','reviewClass: artifact-quality','decision: approved','status: approved',
    "repositoryRevision: $revision",'candidateManifestPath: artifacts/review/evidence/M0019/review-manifest.json',
    "candidateManifestSha256: $manifestHash",
    'wordArtifact: artifacts/review/evidence/M0019/finalized-word.docx',"wordArtifactSha256: $($wordCandidate.sha256)",
    'powerPointArtifact: artifacts/review/evidence/M0019/deck-powerpoint.pptx',"powerPointArtifactSha256: $($powerPointCandidate.sha256)",
    'impressArtifact: artifacts/review/evidence/M0019/deck-impress.pptx',"impressArtifactSha256: $($impressCandidate.sha256)",
    "powerPointTier3Receipt: $($manifest.tier3Receipts.VAL05PowerPoint.path)","powerPointTier3ReceiptSha256: $($manifest.tier3Receipts.VAL05PowerPoint.sha256)",
    "impressTier3Receipt: $($manifest.tier3Receipts.VAL05Impress.path)","impressTier3ReceiptSha256: $($manifest.tier3Receipts.VAL05Impress.sha256)",
    'evidence: artifacts/review/evidence/M0019/review-manifest.json'
)) { if ($review -notmatch "(?m)^$([regex]::Escape($line))\s*$") { throw "M0019 human decision is not bound to '$line'." } }
if ($review -notmatch '(?m)^reviewer:\s*\S.+$' -or $review -notmatch '(?m)^reviewedAtUtc:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z\s*$') { throw 'M0019 approval must name the actual reviewer and UTC review time.' }
Write-Output 'HR-M0019-01: approved and bound to all three current artifacts and native-path evidence.'
