param(
    [ValidateSet('list','show','record')][string]$Command = 'list',
    [ValidateSet('M0005','M0009')][string]$Milestone = 'M0005',
    [ValidateSet('approved','changes-requested','rejected')][string]$Decision,
    [string]$Reviewer,
    [string]$RepositoryRevision,
    [string]$LibreOfficeVersion,
    [string]$WordVersion,
    [string]$EvidenceHash
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot; $id = if ($Milestone -eq 'M0005') { 'HR-M0005-01' } else { 'HR-M0009-01' }; $pending = Join-Path $root ".review/pending/$id.md"; $record = Join-Path $root ".review/records/$id.md"
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
        if ([string]::IsNullOrWhiteSpace($Decision) -or [string]::IsNullOrWhiteSpace($Reviewer) -or [string]::IsNullOrWhiteSpace($EvidenceHash)) { throw 'record requires -Decision, -Reviewer, and -EvidenceHash.' }
        if ($Milestone -eq 'M0005' -and [string]::IsNullOrWhiteSpace($LibreOfficeVersion)) { throw 'M0005 record requires -LibreOfficeVersion.' }
        if ($Milestone -eq 'M0009' -and [string]::IsNullOrWhiteSpace($WordVersion)) { throw 'M0009 record requires -WordVersion.' }
        if ([string]::IsNullOrWhiteSpace($RepositoryRevision)) { $RepositoryRevision = (git -C $root rev-parse HEAD).Trim() }
        if ($Decision -eq 'approved') { $confirmation = Read-Host "A human reviewer must type exactly APPROVE $Milestone"; if ($confirmation -cne "APPROVE $Milestone") { throw 'Human approval confirmation was not provided.' } }
        New-Item -ItemType Directory -Force -Path (Split-Path $record) | Out-Null
        $versionLine = if ($Milestone -eq 'M0005') { "libreOfficeVersion: $LibreOfficeVersion" } else { "wordVersion: $WordVersion" }
        @("---", "milestone: $Milestone", "reviewId: $id", "reviewClass: artifact-quality", "status: $Decision", "decision: $Decision", "reviewer: $Reviewer", "repositoryRevision: $RepositoryRevision", $versionLine, "evidence: $EvidenceHash", "waiver: false", "---", "", "# $id Decision", "", "Recorded by the human reviewer through eng/review.ps1.") | Set-Content -Path $record -Encoding utf8
        Write-Output "Recorded $Decision for $id."
        break
    }
}
