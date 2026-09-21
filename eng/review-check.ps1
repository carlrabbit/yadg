param([Parameter(Mandatory=$true, Position=0)][string]$Milestone, [Parameter(ValueFromRemainingArguments=$true)][string[]]$RemainingArguments)
$ErrorActionPreference = 'Stop'
if ($Milestone -eq '--milestone' -and $RemainingArguments.Count -gt 0) { $Milestone = $RemainingArguments[0] }
$root = Split-Path -Parent $PSScriptRoot
$reviewId = if ($Milestone -eq 'M0005') { 'HR-M0005-01' } elseif ($Milestone -eq 'M0009') { 'HR-M0009-01' } elseif ($Milestone -eq 'M0010') { 'HR-M0010-01' } elseif ($Milestone -eq 'M0011') { 'HR-M0011-01' } else { throw "Unsupported milestone review context '$Milestone'." }
$pending = Join-Path $root ".review/pending/$reviewId.md"
$record = Join-Path $root ".review/records/$reviewId.md"
if (-not (Test-Path $pending)) { throw "Missing canonical review request: $pending" }
if (-not (Test-Path $record)) { Write-Error "$reviewId is still pending human review."; exit 2 }
$text = Get-Content $record -Raw
$required = if ($Milestone -eq 'M0005') { @('milestone: M0005','reviewId: HR-M0005-01','decision: approved','status: approved','reviewer:','repositoryRevision:','libreOfficeVersion:','evidence:') } elseif ($Milestone -eq 'M0009') { @('milestone: M0009','reviewId: HR-M0009-01','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','evidence:') } elseif ($Milestone -eq 'M0010') { @('milestone: M0010','reviewId: HR-M0010-01','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','libreOfficeVersion:','evidence:') } else { @('milestone: M0011','reviewId: HR-M0011-01','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','libreOfficeVersion:','evidence:') }
foreach ($item in $required) {
    if ($text -notmatch [regex]::Escape($item)) { Write-Error "Review record is missing '$item'."; exit 2 }
}
if ($text -match '(?m)^waiver:\s*true\s*$') { Write-Error "Review waiver is forbidden for $Milestone."; exit 2 }
if ($Milestone -eq 'M0009') {
    $artifact = Join-Path $root 'artifacts/review/evidence/M0009/finalized.docx'
    if (-not (Test-Path $artifact)) { Write-Error "M0009 review artifact is missing: $artifact"; exit 2 }
    $match = [regex]::Match($text, '(?m)^evidence:\s*SHA256:([0-9A-Fa-f]+)\s*$')
    $actual = (Get-FileHash $artifact -Algorithm SHA256).Hash
    if (-not $match.Success -or $match.Groups[1].Value.ToUpperInvariant() -ne $actual.ToUpperInvariant()) { Write-Error 'M0009 review evidence hash does not match the current finalized DOCX.'; exit 2 }
}
if ($Milestone -eq 'M0010') {
    $evidenceRoot = Join-Path $root 'artifacts/review/evidence/M0010'
    foreach ($artifact in @('word-origin-word.docx','libreoffice-origin-libreoffice.docx')) {
        if (-not (Test-Path (Join-Path $evidenceRoot $artifact))) { Write-Error "M0010 review artifact is missing: $artifact"; exit 2 }
    }
}
if ($Milestone -eq 'M0011') {
    $evidencePath = Join-Path $root 'artifacts/release/evidence/M0011/release-evidence.json'
    $package = Join-Path $root 'artifacts/package/Yadg.1.0.0.nupkg'
    if (-not (Test-Path $evidencePath)) { Write-Error "M0011 release evidence is missing: $evidencePath"; exit 2 }
    if (-not (Test-Path $package)) { Write-Error "M0011 release candidate package is missing: $package"; exit 2 }
    $evidence = Get-Content $evidencePath -Raw | ConvertFrom-Json
    $actual = (Get-FileHash $package -Algorithm SHA256).Hash.ToUpperInvariant()
    if ([string]$evidence.packageSha256 -ne $actual) { Write-Error 'M0011 release evidence package hash does not match the current package.'; exit 2 }
    $match = [regex]::Match($text, '(?m)^evidence:\s*SHA256:([0-9A-Fa-f]+)\s*$')
    if (-not $match.Success -or $match.Groups[1].Value.ToUpperInvariant() -ne $actual) { Write-Error 'M0011 approval hash does not match the current package.'; exit 2 }
    if ([string]$evidence.releaseVersion -ne '1.0.0' -or [string]$evidence.packageId -ne 'Yadg') { Write-Error 'M0011 release evidence identifies the wrong package.'; exit 2 }
}
Write-Output "$reviewId`: approved"
