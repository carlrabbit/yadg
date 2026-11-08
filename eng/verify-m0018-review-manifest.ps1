param([switch] $AllowPendingHumanReview)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = Join-Path $repo 'artifacts/review/evidence/M0018'
$manifestPath = Join-Path $root 'review-manifest.json'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "M0018 review manifest is missing: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -DateKind String
$revision = (& git -C $repo rev-parse HEAD).Trim()
if ([string]$manifest.milestone -cne 'M0018' -or [int]$manifest.schemaVersion -ne 1 -or [string]$manifest.humanReviewId -cne 'HR-M0018-01') { throw 'M0018 review manifest has an unsupported milestone, schema version or review request.' }
if ([string]$manifest.candidatePath -cne 'artifacts/review/evidence/M0018/review-sample.pptx') { throw 'M0018 review manifest must identify the one canonical review-sample.pptx.' }
if ([string]$manifest.repositoryRevision -cne $revision) { throw 'M0018 review manifest does not bind current repository HEAD.' }

function Get-RepoHash([string] $relativePath) {
    $path = Join-Path $repo $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "M0018 manifest input is missing: $relativePath" }
    return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()
}

$candidateHash = Get-RepoHash ([string]$manifest.candidatePath)
if ($candidateHash -cne [string]$manifest.candidateSha256) { throw 'M0018 candidate hash does not match its review manifest.' }
$sourceHash = Get-RepoHash ([string]$manifest.sourcePath)
if ($sourceHash -cne [string]$manifest.sourceSha256) { throw 'M0018 source hash does not match its review manifest.' }
$imageHash = Get-RepoHash ([string]$manifest.imagePath)
if ($imageHash -cne [string]$manifest.imageSha256) { throw 'M0018 source image hash does not match its review manifest.' }
$templateHash = Get-RepoHash ([string]$manifest.impressTemplatePath)
if ($templateHash -cne [string]$manifest.impressTemplateSha256) { throw 'M0018 Impress-origin template hash does not match its review manifest.' }
$templateProvenancePath = Join-Path $repo 'tests/fixtures/m0018/impress-origin.provenance.json'
$templateProvenance = Get-Content -LiteralPath $templateProvenancePath -Raw | ConvertFrom-Json -DateKind String
if ([string]$templateProvenance.sha256 -cne $templateHash -or [string]$manifest.impressTemplatePath -cne ([string]$templateProvenance.path).Replace('\', '/')) { throw 'M0018 review manifest does not identify the canonical Impress-origin fixture.' }

$receiptSpecs = @(
    @{ key='powerPointTier3'; path='artifacts/review/evidence/M0018/powerpoint-tier3-evidence.json'; gate='VAL-06'; application='Microsoft PowerPoint' },
    @{ key='impressTier3'; path='artifacts/review/evidence/M0018/impress-tier3-evidence.json'; gate='VAL-03'; application='LibreOffice Impress' }
)
foreach ($spec in $receiptSpecs) {
    $entry = $manifest.$($spec.key)
    if ([string]$entry.path -cne $spec.path) { throw "M0018 manifest must bind the canonical $($spec.key) receipt." }
    $receiptHash = Get-RepoHash $spec.path
    if ($receiptHash -cne [string]$entry.sha256) { throw "M0018 $($spec.key) receipt hash is stale." }
    $receiptPath = Join-Path $repo $spec.path
    $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json -DateKind String
    if ([string]$receipt.milestone -cne 'M0018' -or [string]$receipt.gate -cne $spec.gate -or [string]$receipt.result -cne 'passed' -or [string]$receipt.repositoryRevision -cne $revision -or [string]$receipt.application -cne $spec.application) { throw "M0018 $($spec.key) receipt does not represent a current successful origin-specific Tier 3 run." }
    if ([string]::IsNullOrWhiteSpace([string]$receipt.applicationVersion) -or [string]$receipt.openResult -notin @('opened-without-repair','user-confirmed-open-without-repair')) { throw "M0018 $($spec.key) receipt does not establish a real Office open." }
    $formatting = if ($spec.key -eq 'powerPointTier3') { $receipt.formattingValidation } else { $receipt.inspection.formattingValidation }
    if ([string]$formatting.result -cne 'passed') { throw "M0018 $($spec.key) receipt does not establish deterministic formatting validation." }
    $formatRoles = @($formatting.comparisons | ForEach-Object { [string]$_.role })
    foreach ($role in @('title','paragraph','heading','table-header','table-body')) {
        if ($formatRoles -cnotcontains $role) { throw "M0018 $($spec.key) formatting receipt is missing the '$role' example comparison." }
    }
    if (-not ($formatRoles -ccontains 'list' -or $formatRoles -ccontains 'nested-list')) { throw "M0018 $($spec.key) formatting receipt is missing a representative list comparison." }
    if ($spec.key -eq 'impressTier3') {
        if (([string]$receipt.outputPath).Replace('\', '/') -cne [string]$manifest.candidatePath -or [string]$receipt.outputSha256 -cne $candidateHash -or [string]$receipt.templateSha256 -cne $templateHash -or [string]$receipt.sourceSha256 -cne $sourceHash) { throw 'M0018 Impress-origin receipt does not bind the exact review candidate, source and fixture.' }
        if (-not [bool]$receipt.profileCleaned) { throw 'M0018 Impress-origin receipt does not prove isolated profile cleanup.' }
    }
    if ($spec.key -eq 'powerPointTier3') {
        $powerPointTemplatePath = 'tests/fixtures/m0018/powerpoint-origin.pptx'
        $powerPointTemplateHash = Get-RepoHash $powerPointTemplatePath
        if (([string]$receipt.templatePath).Replace('\', '/') -cne $powerPointTemplatePath -or [string]$receipt.templateSha256 -cne $powerPointTemplateHash -or [string]$receipt.sourceSha256 -cne $sourceHash) { throw 'M0018 PowerPoint-origin receipt does not bind the canonical fixture and shared source.' }
        if ([bool]$receipt.presentationClosed -ne $true -or ([bool]$receipt.applicationOwned -and [bool]$receipt.applicationQuit -ne $true)) { throw 'M0018 PowerPoint-origin receipt does not establish Office cleanup.' }
        $reviewCandidateOpen = $receipt.reviewCandidatePowerPointOpen
        if (([string]$reviewCandidateOpen.path).Replace('\', '/') -cne [string]$manifest.candidatePath -or [string]$reviewCandidateOpen.sha256 -cne $candidateHash -or [string]$reviewCandidateOpen.result -cne 'opened-without-repair' -or [int]$reviewCandidateOpen.slideCount -ne 4) { throw 'PowerPoint did not reopen the exact four-slide review candidate.' }
    }
    if ([string]$manifest.officeVersions.$($spec.key) -cne [string]$receipt.applicationVersion) { throw "M0018 manifest $($spec.key) version does not match its receipt." }
}

$reviewRecordPath = Join-Path $repo '.review/records/HR-M0018-01.md'
if (Test-Path -LiteralPath $reviewRecordPath -PathType Leaf) {
    $reviewRecord = Get-Content -LiteralPath $reviewRecordPath -Raw
    $manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToUpperInvariant()
    foreach ($line in @(
        'milestone: M0018', 'reviewId: HR-M0018-01', 'decision: approved', 'status: approved',
        "repositoryRevision: $revision", "candidatePath: $($manifest.candidatePath)", "candidateSha256: $candidateHash",
        "candidateManifestPath: artifacts/review/evidence/M0018/review-manifest.json", "candidateManifestSha256: $manifestHash",
        "powerPointTier3EvidenceSha256: $($manifest.powerPointTier3.sha256)", "impressTier3EvidenceSha256: $($manifest.impressTier3.sha256)"
    )) { if ($reviewRecord -notmatch "(?m)^$([regex]::Escape($line))\s*$") { throw "M0018 human review record is not bound to '$line'." } }
    if ($reviewRecord -notmatch '(?m)^reviewer:\s*\S.+$' -or $reviewRecord -notmatch '(?m)^reviewedAtUtc:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z\s*$') { throw 'M0018 human review record must contain a named reviewer and UTC review time.' }
} elseif (-not $AllowPendingHumanReview) {
    Write-Output 'M0018 evidence manifest is current; HR-M0018-01 remains pending human review.'
}

Write-Output "M0018 review manifest verified: $manifestPath"
