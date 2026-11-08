param(
    [ValidateSet('list','show','record')][string]$Command = 'list',
    [ValidateSet('M0005','M0009','M0010','M0011','M0014','M0015','M0016','M0017')][string]$Milestone = 'M0005',
    [ValidateSet('approved','changes-requested','rejected')][string]$Decision,
    [string]$Reviewer,
    [string]$RepositoryRevision,
    [string]$LibreOfficeVersion,
    [string]$WordVersion,
    [string]$EvidenceHash,
    [string]$PackageHash
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot; $id = if ($Milestone -eq 'M0005') { 'HR-M0005-01' } elseif ($Milestone -eq 'M0009') { 'HR-M0009-01' } elseif ($Milestone -eq 'M0010') { 'HR-M0010-01' } elseif ($Milestone -eq 'M0014') { 'HR-M0014-01' } elseif ($Milestone -in @('M0015','M0016','M0017')) { "HR-$Milestone-01" } else { 'HR-M0011-01' }; $pending = Join-Path $root ".review/pending/$id.md"; $record = Join-Path $root ".review/records/$id.md"
switch ($Command) {
    'list' {
        if (Test-Path $pending) { Get-Item $pending | Select-Object FullName,Length,LastWriteTime }
        if (Test-Path $record) { Get-Item $record | Select-Object FullName,Length,LastWriteTime }
        break
    }
    'show' {
        if (Test-Path $record) { Get-Content $record } elseif (Test-Path $pending) { Get-Content $pending } else { throw "No review request or record exists for $Milestone." }
        break
    }
    'record' {
        if ([string]::IsNullOrWhiteSpace($Decision) -or [string]::IsNullOrWhiteSpace($Reviewer)) { throw 'record requires -Decision and -Reviewer.' }
        if ($Milestone -notin @('M0011','M0015') -and [string]::IsNullOrWhiteSpace($EvidenceHash)) { throw 'record requires -EvidenceHash.' }
        if ($Milestone -eq 'M0005' -and [string]::IsNullOrWhiteSpace($LibreOfficeVersion)) { throw 'M0005 record requires -LibreOfficeVersion.' }
        if ($Milestone -in @('M0009','M0010','M0014') -and [string]::IsNullOrWhiteSpace($WordVersion)) { throw "$Milestone record requires -WordVersion." }
        if ($Milestone -in @('M0010','M0014') -and [string]::IsNullOrWhiteSpace($LibreOfficeVersion)) { throw "$Milestone record requires -LibreOfficeVersion." }
        if ($Milestone -eq 'M0011') {
            $evidencePath = Join-Path $root 'artifacts/release/evidence/M0011/release-evidence.json'
            $packagePath = Join-Path $root 'artifacts/package/Yadg.1.0.0.nupkg'
            if (-not (Test-Path -LiteralPath $evidencePath) -or -not (Test-Path -LiteralPath $packagePath)) { throw 'M0011 release evidence and current package are required before recording review.' }
            $releaseEvidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
            $PackageHash = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant()
            if ($PackageHash -ne ([string]$releaseEvidence.packageSha256).ToUpperInvariant()) { throw 'Current package does not match M0011 release evidence.' }
            $RepositoryRevision = ([string]$releaseEvidence.repositoryRevision)
            $WordVersion = ([string]$releaseEvidence.wordVersion)
            $LibreOfficeVersion = ([string]$releaseEvidence.libreOfficeVersion)
            $EvidenceHash = $PackageHash
        }
        if ($Milestone -eq 'M0015') {
            $evidencePath = Join-Path $root 'artifacts/release/evidence/M0015/release-evidence.json'
            $props = [xml](Get-Content -LiteralPath (Join-Path $root 'Directory.Build.props') -Raw)
            $releaseVersion = [string]$props.Project.PropertyGroup.VersionPrefix
            $packagePath = Join-Path $root "artifacts/package/Yadg.$releaseVersion.nupkg"
            if ($releaseVersion -ne '1.2.0' -or -not (Test-Path -LiteralPath $evidencePath) -or -not (Test-Path -LiteralPath $packagePath)) { throw 'M0015 requires its current 1.2.0 package and automated release evidence before review can be recorded.' }
            $releaseEvidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
            $RepositoryRevision = (& git -C $root rev-parse HEAD).Trim()
            $tree = (& git -C $root status --porcelain --untracked-files=all) -join "`n"
            if (-not [string]::IsNullOrWhiteSpace($tree) -or [string]$releaseEvidence.repositoryRevision -ne $RepositoryRevision -or -not [string]::IsNullOrWhiteSpace([string]$releaseEvidence.workingTreeStatus)) { throw 'M0015 release evidence must be bound to the current clean candidate revision.' }
            $PackageHash = (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant()
            if ([string]$releaseEvidence.releaseVersion -ne $releaseVersion -or [string]$releaseEvidence.packageId -ne 'Yadg' -or [string]$releaseEvidence.packageSha256 -ne $PackageHash -or [string]$releaseEvidence.externalPublication -ne 'none') { throw 'M0015 current package does not match exact automated release evidence.' }
            if ([string]$releaseEvidence.m0014Tier3.repositoryRevision -ne $RepositoryRevision) { throw 'Fresh M0014 Tier 3 evidence is not bound to the M0015 candidate.' }
            $WordVersion = [string]$releaseEvidence.wordVersion
            $LibreOfficeVersion = [string]$releaseEvidence.libreOfficeVersion
            $wordRun = @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'word' -and $_.result -eq 'passed' })[0]
            $libreOfficeRun = @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'libreoffice' -and $_.result -eq 'passed' })[0]
            $EvidenceHash = "SHA256:$PackageHash"
        }
        if ($Milestone -in @('M0016','M0017')) {
            $evidenceRoot = Join-Path $root "artifacts/review/evidence/$Milestone"
            $evidencePath = Join-Path $evidenceRoot 'tier3-evidence.json'
            $consumerEvidencePath = Join-Path $evidenceRoot 'tier4-evidence.json'
            if (-not (Test-Path -LiteralPath $evidencePath) -or -not (Test-Path -LiteralPath $consumerEvidencePath)) { throw "$Milestone requires successful real-renderer and packed-consumer evidence before review can be recorded." }
            $releaseEvidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
            $consumerEvidence = Get-Content -LiteralPath $consumerEvidencePath -Raw | ConvertFrom-Json
            if ([string]$releaseEvidence.milestone -ne $Milestone -or @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'word' -and $_.result -eq 'passed' }).Count -ne 1 -or @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'libreoffice' -and $_.result -eq 'passed' }).Count -ne 1) { throw "$Milestone Tier 3 evidence must contain one successful Word run and one successful LibreOffice run." }
            if ([string]$consumerEvidence.milestone -ne $Milestone -or [string]$consumerEvidence.result -ne 'passed') { throw "$Milestone packed-consumer evidence is incomplete." }
            $wordRun = @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'word' -and $_.result -eq 'passed' })[0]
            $libreOfficeRun = @($releaseEvidence.runs | Where-Object { $_.renderer -eq 'libreoffice' -and $_.result -eq 'passed' })[0]
            $RepositoryRevision = [string]$releaseEvidence.repositoryRevision
            $WordVersion = [string]$releaseEvidence.wordVersion
            $LibreOfficeVersion = [string]$releaseEvidence.libreOfficeVersion
            $tier3Hash = (Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash.ToUpperInvariant()
            if ($EvidenceHash -notmatch '^SHA256:[0-9A-Fa-f]{64}$' -or $EvidenceHash.Substring(7).ToUpperInvariant() -ne $tier3Hash) { throw "$Milestone EvidenceHash must identify the current Tier 3 evidence file." }
            if ([string]$releaseEvidence.repositoryRevision -ne (& git -C $root rev-parse HEAD).Trim()) { throw "$Milestone Tier 3 evidence must be bound to current repository HEAD." }
            foreach ($run in $releaseEvidence.runs) {
                $artifactPath = Join-Path $root ([string]$run.finalized)
                if (-not (Test-Path -LiteralPath $artifactPath) -or (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash.ToUpperInvariant() -ne [string]$run.finalizedSha256) { throw "$Milestone finalized artifact does not match Tier 3 evidence: $artifactPath" }
            }
            $PackageHash = [string]$consumerEvidence.packageSha256
            $packagePath = Join-Path $root ([string]$consumerEvidence.packagePath)
            if (-not (Test-Path -LiteralPath $packagePath) -or (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToUpperInvariant() -ne $PackageHash.ToUpperInvariant()) { throw "$Milestone packed-consumer evidence does not match the current package." }
        }
        if ([string]::IsNullOrWhiteSpace($RepositoryRevision)) { $RepositoryRevision = (git -C $root rev-parse HEAD).Trim() }
        if ($Decision -eq 'approved') { $confirmation = Read-Host "A human reviewer must type exactly APPROVE $Milestone"; if ($confirmation -cne "APPROVE $Milestone") { throw 'Human approval confirmation was not provided.' } }
        New-Item -ItemType Directory -Force -Path (Split-Path $record) | Out-Null
        $versionLine = if ($Milestone -eq 'M0005') { @("libreOfficeVersion: $LibreOfficeVersion") } elseif ($Milestone -eq 'M0009') { @("wordVersion: $WordVersion") } else { @("wordVersion: $WordVersion", "libreOfficeVersion: $LibreOfficeVersion") }
        $reviewClass = if ($Milestone -in @('M0011','M0015')) { 'release-readiness' } else { 'artifact-quality' }
        $reviewEvidence = if ($Milestone -eq 'M0011') { "SHA256:$PackageHash" } else { $EvidenceHash }
        $derivedLines = if ($Milestone -eq 'M0011') { @("packageSha256: $PackageHash", "automatedEvidence: artifacts/release/evidence/M0011/release-evidence.json") } elseif ($Milestone -eq 'M0015') { @("packageSha256: $PackageHash", "automatedEvidence: artifacts/release/evidence/M0015/release-evidence.json") } elseif ($Milestone -in @('M0016','M0017')) { @("automatedEvidence: artifacts/review/evidence/$Milestone/tier3-evidence.json", "tier3EvidenceSha256: $($EvidenceHash.Substring(7).ToUpperInvariant())", "wordArtifact: $($wordRun.finalized)", "wordArtifactSha256: $($wordRun.finalizedSha256)", "libreOfficeArtifact: $($libreOfficeRun.finalized)", "libreOfficeArtifactSha256: $($libreOfficeRun.finalizedSha256)", "packageSha256: $PackageHash", "automatedConsumerEvidence: artifacts/review/evidence/$Milestone/tier4-evidence.json") } else { @() }
        @("---", "milestone: $Milestone", "reviewId: $id", "reviewClass: $reviewClass", "status: $Decision", "decision: $Decision", "reviewer: $Reviewer", "reviewedAtUtc: $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))", "repositoryRevision: $RepositoryRevision", $versionLine, $derivedLines, "evidence: $reviewEvidence", "waiver: false", "---", "", "# $id Decision", "", "Recorded by the human reviewer through eng/review.ps1.") | Set-Content -Path $record -Encoding utf8
        Write-Output "Recorded $Decision for $id."
        break
    }
}
