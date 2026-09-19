param([Parameter(Mandatory=$true)][string]$Milestone)
$ErrorActionPreference = 'Stop'
if ($Milestone -ne 'M0005') { throw "Unsupported milestone review context '$Milestone'." }
$root = Split-Path -Parent $PSScriptRoot
$pending = Join-Path $root '.review/pending/HR-M0005-01.md'
$record = Join-Path $root '.review/records/HR-M0005-01.md'
if (-not (Test-Path $pending)) { throw "Missing canonical review request: $pending" }
if (-not (Test-Path $record)) { Write-Error 'HR-M0005-01 is still pending human review.'; exit 2 }
$text = Get-Content $record -Raw
foreach ($required in @('milestone: M0005','reviewId: HR-M0005-01','decision: approved','status: approved','reviewer:','repositoryRevision:','libreOfficeVersion:','evidence:')) {
    if ($text -notmatch [regex]::Escape($required)) { Write-Error "Review record is missing '$required'."; exit 2 }
}
if ($text -match '(?m)^waiver:\s*true\s*$') { Write-Error 'Review waiver is forbidden for M0005.'; exit 2 }
Write-Output 'HR-M0005-01: approved'
