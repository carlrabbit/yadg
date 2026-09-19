param(
    [ValidateSet('list','show','record')][string]$Command = 'list',
    [string]$Milestone = 'M0005',
    [ValidateSet('approved','changes-requested','rejected')][string]$Decision,
    [string]$Reviewer,
    [string]$RepositoryRevision,
    [string]$LibreOfficeVersion,
    [string]$EvidenceHash
)
$ErrorActionPreference = 'Stop'
if ($Milestone -ne 'M0005') { throw "Unsupported milestone review context '$Milestone'." }
$root = Split-Path -Parent $PSScriptRoot; $id = 'HR-M0005-01'; $pending = Join-Path $root ".review/pending/$id.md"; $record = Join-Path $root ".review/records/$id.md"
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
        if ([string]::IsNullOrWhiteSpace($Decision) -or [string]::IsNullOrWhiteSpace($Reviewer) -or [string]::IsNullOrWhiteSpace($LibreOfficeVersion) -or [string]::IsNullOrWhiteSpace($EvidenceHash)) { throw 'record requires -Decision, -Reviewer, -LibreOfficeVersion, and -EvidenceHash.' }
        if ([string]::IsNullOrWhiteSpace($RepositoryRevision)) { $RepositoryRevision = (git -C $root rev-parse HEAD).Trim() }
        if ($Decision -eq 'approved') { $confirmation = Read-Host 'A human reviewer must type exactly APPROVE M0005 to record approval'; if ($confirmation -cne 'APPROVE M0005') { throw 'Human approval confirmation was not provided.' } }
        New-Item -ItemType Directory -Force -Path (Split-Path $record) | Out-Null
        @("---", "milestone: M0005", "reviewId: HR-M0005-01", "reviewClass: artifact-quality", "status: $Decision", "decision: $Decision", "reviewer: $Reviewer", "repositoryRevision: $RepositoryRevision", "libreOfficeVersion: $LibreOfficeVersion", "evidence: $EvidenceHash", "waiver: false", "---", "", "# HR-M0005-01 Decision", "", "Recorded by the human reviewer through eng/review.ps1.") | Set-Content -Path $record -Encoding utf8
        Write-Output "Recorded $Decision for HR-M0005-01."
        break
    }
}
