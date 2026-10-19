param([Parameter(Mandatory=$true, Position=0)][string]$Milestone, [Parameter(ValueFromRemainingArguments=$true)][string[]]$RemainingArguments)
$ErrorActionPreference = 'Stop'
if ($Milestone -eq '--milestone' -and $RemainingArguments.Count -gt 0) { $Milestone = $RemainingArguments[0] }
$root = Split-Path -Parent $PSScriptRoot
$reviewId = if ($Milestone -eq 'M0005') { 'HR-M0005-01' } elseif ($Milestone -eq 'M0009') { 'HR-M0009-01' } elseif ($Milestone -eq 'M0010') { 'HR-M0010-01' } elseif ($Milestone -eq 'M0011') { 'HR-M0011-01' } elseif ($Milestone -eq 'M0013') { 'HR-M0013-01' } elseif ($Milestone -eq 'M0014') { 'HR-M0014-01' } elseif ($Milestone -eq 'M0015') { 'HR-M0015-01' } elseif ($Milestone -eq 'M0016') { 'HR-M0016-01' } else { throw "Unsupported milestone review context '$Milestone'." }
$pending = Join-Path $root ".review/pending/$reviewId.md"
$record = Join-Path $root ".review/records/$reviewId.md"
if (-not (Test-Path $pending)) { throw "Missing canonical review request: $pending" }
if (-not (Test-Path $record)) { Write-Error "$reviewId is still pending human review."; exit 2 }
$text = Get-Content $record -Raw
$required = if ($Milestone -eq 'M0005') { @('milestone: M0005','reviewId: HR-M0005-01','decision: approved','status: approved','reviewer:','repositoryRevision:','libreOfficeVersion:','evidence:') } elseif ($Milestone -eq 'M0009') { @('milestone: M0009','reviewId: HR-M0009-01','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','evidence:') } elseif ($Milestone -eq 'M0010') { @('milestone: M0010','reviewId: HR-M0010-01','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','libreOfficeVersion:','evidence:') } elseif ($Milestone -in @('M0011','M0014')) { @("milestone: $Milestone","reviewId: HR-$Milestone-01",'decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','libreOfficeVersion:','evidence:') } elseif ($Milestone -eq 'M0015') { @('milestone: M0015','reviewId: HR-M0015-01','reviewClass: release-readiness','decision: approved','status: approved','reviewer:','repositoryRevision:','wordVersion:','libreOfficeVersion:','packageSha256:','automatedEvidence: artifacts/release/evidence/M0015/release-evidence.json') } elseif ($Milestone -eq 'M0016') { @('milestone: M0016','reviewId: HR-M0016-01','reviewClass: artifact-quality','decision: approved','status: approved','reviewer:','reviewedAtUtc:','repositoryRevision:','wordVersion:','libreOfficeVersion:','tier3EvidenceSha256:','wordArtifactSha256:','libreOfficeArtifactSha256:','packageSha256:','automatedEvidence: artifacts/review/evidence/M0016/tier3-evidence.json','automatedConsumerEvidence: artifacts/review/evidence/M0016/tier4-evidence.json') } else { @('milestone: M0013','reviewId: HR-M0013-01','decision: approved','status: approved','reviewer:','repositoryRevision:','packageSha256:','automatedEvidence: artifacts/release/evidence/M0013/release-evidence.json') }
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
if ($Milestone -eq 'M0014') {
    $evidenceRoot = Join-Path $root 'artifacts/review/evidence/M0014'
    $evidencePath = Join-Path $evidenceRoot 'tier3-evidence.json'
    if (-not (Test-Path -LiteralPath $evidencePath)) { Write-Error 'M0014 Tier 3 evidence is missing.'; exit 2 }
    foreach ($artifact in @('word-diagram.docx','libreoffice-diagram.docx')) {
        if (-not (Test-Path (Join-Path $evidenceRoot $artifact))) { Write-Error "M0014 review artifact is missing: $artifact"; exit 2 }
    }
    $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
    $hash = (Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($text -notmatch "(?m)^evidence:\s*SHA256:$hash\s*$") { Write-Error 'M0014 approval is not bound to current Tier 3 evidence.'; exit 2 }
    $revisionMatch = [regex]::Match($text, '(?m)^repositoryRevision:\s*(\S+)\s*$')
    if (-not $revisionMatch.Success -or $revisionMatch.Groups[1].Value -ne [string]$evidence.repositoryRevision) { Write-Error 'M0014 review revision does not match Tier 3 evidence.'; exit 2 }
    foreach ($run in $evidence.runs) {
        $artifactPath = Join-Path $root ([string]$run.finalized)
        if (-not (Test-Path -LiteralPath $artifactPath)) { Write-Error "M0014 finalized evidence artifact is missing: $artifactPath"; exit 2 }
        if ((Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash -ne [string]$run.finalizedSha256) { Write-Error "M0014 finalized artifact hash does not match evidence: $artifactPath"; exit 2 }
        if ([string]$run.result -ne 'passed' -or [int]$run.svgPartCount -lt 3) { Write-Error 'M0014 Tier 3 evidence is incomplete.'; exit 2 }
    }
    if (@($evidence.runs | Where-Object { $_.renderer -eq 'word' -and $_.result -eq 'passed' }).Count -ne 1 -or @($evidence.runs | Where-Object { $_.renderer -eq 'libreoffice' -and $_.result -eq 'passed' }).Count -ne 1) { Write-Error 'M0014 Tier 3 evidence must include successful Word and LibreOffice runs.'; exit 2 }
}
if ($Milestone -eq 'M0011') {
    $evidencePath = Join-Path $root 'artifacts/release/evidence/M0011/release-evidence.json'
    $package = Join-Path $root 'artifacts/package/Yadg.1.0.0.nupkg'
    if (-not (Test-Path $evidencePath)) { Write-Error "M0011 release evidence is missing: $evidencePath"; exit 2 }
    if (-not (Test-Path $package)) { Write-Error "M0011 release candidate package is missing: $package"; exit 2 }
    $evidence = Get-Content $evidencePath -Raw | ConvertFrom-Json
    $actual = (Get-FileHash $package -Algorithm SHA256).Hash.ToUpperInvariant()
    if ([string]$evidence.packageSha256 -ne $actual) { Write-Error 'M0011 release evidence package hash does not match the current package.'; exit 2 }
    $match = [regex]::Match($text, '(?m)^packageSha256:\s*([0-9A-Fa-f]+)\s*$')
    if (-not $match.Success -or $match.Groups[1].Value.ToUpperInvariant() -ne $actual) { Write-Error 'M0011 approval package hash does not match the current package.'; exit 2 }
    $revisionMatch = [regex]::Match($text, '(?m)^repositoryRevision:\s*(\S+)\s*$')
    if (-not $revisionMatch.Success -or $revisionMatch.Groups[1].Value -ne [string]$evidence.repositoryRevision) { Write-Error 'M0011 approval revision does not match release evidence.'; exit 2 }
    if ($text -notmatch '(?m)^automatedEvidence:\s*artifacts/release/evidence/M0011/release-evidence\.json\s*$') { Write-Error 'M0011 approval is not bound to automated release evidence.'; exit 2 }
    if ([string]$evidence.releaseVersion -ne '1.0.0' -or [string]$evidence.packageId -ne 'Yadg') { Write-Error 'M0011 release evidence identifies the wrong package.'; exit 2 }
}
if ($Milestone -eq 'M0013') {
    $props = [xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)
    $version = [string]$props.Project.PropertyGroup.VersionPrefix
    $evidencePath = Join-Path $root 'artifacts/release/evidence/M0013/release-evidence.json'
    $package = Join-Path $root "artifacts/package/Yadg.$version.nupkg"
    if (-not (Test-Path $evidencePath)) { Write-Error "M0013 release evidence is missing: $evidencePath"; exit 2 }
    if (-not (Test-Path $package)) { Write-Error "M0013 release candidate package is missing: $package"; exit 2 }
    $evidence = Get-Content $evidencePath -Raw | ConvertFrom-Json
    $currentRevision = (& git -C $root rev-parse HEAD).Trim()
    if ([string]$evidence.repositoryRevision -ne $currentRevision) { Write-Error 'M0013 release evidence revision does not match current repository HEAD.'; exit 2 }
    $actual = (Get-FileHash $package -Algorithm SHA256).Hash.ToUpperInvariant()
    if ([string]$evidence.packageSha256 -ne $actual -or [string]$evidence.releaseVersion -ne $version -or [string]$evidence.packageId -ne 'Yadg') { Write-Error 'M0013 evidence does not bind the current package ID/version/hash.'; exit 2 }
    $match = [regex]::Match($text, '(?m)^packageSha256:\s*([0-9A-Fa-f]+)\s*$')
    if (-not $match.Success -or $match.Groups[1].Value.ToUpperInvariant() -ne $actual) { Write-Error 'M0013 approval package hash does not match the current package.'; exit 2 }
    $revisionMatch = [regex]::Match($text, '(?m)^repositoryRevision:\s*(\S+)\s*$')
    if (-not $revisionMatch.Success -or $revisionMatch.Groups[1].Value -ne [string]$evidence.repositoryRevision) { Write-Error 'M0013 approval revision does not match release evidence.'; exit 2 }
    if ($revisionMatch.Groups[1].Value -ne $currentRevision) { Write-Error 'M0013 human review revision does not match current repository HEAD.'; exit 2 }
    if ([string]$evidence.externalPublication -notmatch '^none') { Write-Error 'M0013 evidence does not confirm external publication exclusion.'; exit 2 }
}
if ($Milestone -eq 'M0015') {
    $props = [xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)
    $version = [string]$props.Project.PropertyGroup.VersionPrefix
    if ($version -ne '1.2.0') { Write-Error "M0015 central version is '$version', expected 1.2.0."; exit 2 }
    $evidencePath = Join-Path $root 'artifacts/release/evidence/M0015/release-evidence.json'
    $package = Join-Path $root "artifacts/package/Yadg.$version.nupkg"
    if (-not (Test-Path -LiteralPath $evidencePath) -or -not (Test-Path -LiteralPath $package)) { Write-Error 'M0015 release evidence or exact package is missing.'; exit 2 }
    $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
    $currentRevision = (& git -C $root rev-parse HEAD).Trim()
    $currentTree = (& git -C $root status --porcelain --untracked-files=all) -join "`n"
    if ([string]$evidence.repositoryRevision -ne $currentRevision -or [string]$text -notmatch "(?m)^repositoryRevision:\s*$([regex]::Escape($currentRevision))\s*$") { Write-Error 'M0015 review/evidence revision does not match current HEAD.'; exit 2 }
    if (-not [string]::IsNullOrWhiteSpace($currentTree) -or -not [string]::IsNullOrWhiteSpace([string]$evidence.workingTreeStatus)) { Write-Error 'M0015 release evidence and current repository tree must be clean.'; exit 2 }
    $actual = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToUpperInvariant()
    $hashMatch = [regex]::Match($text, '(?m)^packageSha256:\s*([0-9A-Fa-f]+)\s*$')
    if ([string]$evidence.releaseVersion -ne $version -or [string]$evidence.packageId -ne 'Yadg' -or [string]$evidence.packageSha256 -ne $actual -or -not $hashMatch.Success -or $hashMatch.Groups[1].Value.ToUpperInvariant() -ne $actual) { Write-Error 'M0015 review/evidence does not bind Yadg 1.2.0 and its current package hash.'; exit 2 }
    $wordMatch = [regex]::Match($text, '(?m)^wordVersion:\s*(\S+)\s*$')
    $libreMatch = [regex]::Match($text, '(?m)^libreOfficeVersion:\s*(\S+)\s*$')
    if (-not $wordMatch.Success -or $wordMatch.Groups[1].Value -ne [string]$evidence.wordVersion -or -not $libreMatch.Success -or $libreMatch.Groups[1].Value -ne [string]$evidence.libreOfficeVersion) { Write-Error 'M0015 review runtime versions do not match automated release evidence.'; exit 2 }
    if ([string]$evidence.externalPublication -ne 'none') { Write-Error 'M0015 evidence must state externalPublication: none.'; exit 2 }
    if ([string]$evidence.m0014Tier3.repositoryRevision -ne $currentRevision) { Write-Error 'M0015 evidence does not bind fresh M0014 Tier 3 to the release candidate revision.'; exit 2 }
    if ([string]$text -notmatch '(?m)^automatedEvidence:\s*artifacts/release/evidence/M0015/release-evidence\.json\s*$') { Write-Error 'M0015 review is not bound to its automated release evidence.'; exit 2 }
}
if ($Milestone -eq 'M0016') {
    $evidenceRoot = Join-Path $root 'artifacts/review/evidence/M0016'
    $evidencePath = Join-Path $evidenceRoot 'tier3-evidence.json'
    $consumerEvidencePath = Join-Path $evidenceRoot 'tier4-evidence.json'
    if (-not (Test-Path -LiteralPath $evidencePath) -or -not (Test-Path -LiteralPath $consumerEvidencePath)) { Write-Error 'M0016 Tier 3 or Tier 4 evidence is missing.'; exit 2 }
    $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
    $consumerEvidence = Get-Content -LiteralPath $consumerEvidencePath -Raw | ConvertFrom-Json
    $evidenceHash = (Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash.ToUpperInvariant()
    $evidenceLinePattern = '(?m)^evidence:\s*SHA256:' + [regex]::Escape($evidenceHash) + '\s*$'
    if ($text -notmatch "(?m)^tier3EvidenceSha256:\s*$evidenceHash\s*$" -or $text -notmatch $evidenceLinePattern) { Write-Error 'M0016 approval does not bind the exact Tier 3 evidence hash.'; exit 2 }
    if ([string]$evidence.milestone -ne 'M0016' -or [string]$evidence.repositoryRevision -ne (& git -C $root rev-parse HEAD).Trim()) { Write-Error 'M0016 Tier 3 evidence does not identify this milestone and current repository revision.'; exit 2 }
    if ([string]$evidence.wordVersion -ne ([regex]::Match($text, '(?m)^wordVersion:\s*(\S+)\s*$').Groups[1].Value) -or [string]$evidence.libreOfficeVersion -ne ([regex]::Match($text, '(?m)^libreOfficeVersion:\s*(\S+)\s*$').Groups[1].Value)) { Write-Error 'M0016 review runtime versions do not match Tier 3 evidence.'; exit 2 }
    foreach ($renderer in @('word','libreoffice')) {
        $run = @($evidence.runs | Where-Object { $_.renderer -eq $renderer -and $_.result -eq 'passed' })
        if ($run.Count -ne 1) { Write-Error "M0016 Tier 3 requires one passed $renderer run."; exit 2 }
        $artifactPath = Join-Path $root ([string]$run[0].finalized)
        if (-not (Test-Path -LiteralPath $artifactPath) -or (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash.ToUpperInvariant() -ne [string]$run[0].finalizedSha256) { Write-Error "M0016 finalized $renderer artifact hash does not match evidence."; exit 2 }
        $expectedPath = "artifacts/review/evidence/M0016/$renderer-finalized.docx"
        $expectedHash = [string]$run[0].finalizedSha256
        if ($text -notmatch "(?m)^${renderer}Artifact:\s*$([regex]::Escape($expectedPath))\s*$" -or $text -notmatch "(?m)^${renderer}ArtifactSha256:\s*$expectedHash\s*$") { Write-Error "M0016 approval is not bound to the exact finalized $renderer DOCX."; exit 2 }
    }
    $packagePath = Join-Path $root ([string]$consumerEvidence.packagePath)
    if ([string]$consumerEvidence.result -ne 'passed' -or -not (Test-Path -LiteralPath $packagePath)) { Write-Error 'M0016 packed-consumer evidence or package is incomplete.'; exit 2 }
    $packageHash = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($packageHash -ne [string]$consumerEvidence.packageSha256 -or $text -notmatch "(?m)^packageSha256:\s*$packageHash\s*$") { Write-Error 'M0016 approval/package evidence hash does not match the current package.'; exit 2 }
    if ([string]$consumerEvidence.packageVersion -ne '1.2.0' -or [string]$consumerEvidence.milestone -ne 'M0016') { Write-Error 'M0016 consumer evidence identifies the wrong package or milestone.'; exit 2 }
    if ($text -notmatch '(?m)^reviewer:\s*\S.+$' -or $text -notmatch '(?m)^reviewedAtUtc:\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z\s*$') { Write-Error 'M0016 approval must include a named reviewer and recorded UTC review time.'; exit 2 }
}
Write-Output "$reviewId`: approved"
