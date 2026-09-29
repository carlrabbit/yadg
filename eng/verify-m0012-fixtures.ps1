$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$manifestPath = Join-Path $repo 'tests/fixtures/m0012/provenance.json'
if (-not (Test-Path -LiteralPath $manifestPath)) { throw "M0012 fixture provenance is missing: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ($manifest.fixtures.Count -lt 4) { throw 'M0012 requires Word-produced style, prototype, related-list recovery, and built-in-list recovery fixtures.' }
foreach ($required in @('related-list-recovery','builtin-list-recovery')) { if (-not ($manifest.fixtures | Where-Object { [string]$_.id -eq $required })) { throw "M0012 recovery fixture is missing: $required" } }
foreach ($fixture in $manifest.fixtures) {
    $path = Join-Path $repo ([string]$fixture.path)
    if (-not (Test-Path -LiteralPath $path)) { throw "M0012 fixture is missing: $path" }
    if ([string]$fixture.application -ne 'Microsoft Word' -or -not [bool]$fixture.synthetic -or -not [bool]$fixture.nonConfidential) { throw "M0012 fixture provenance is incomplete: $path" }
    foreach ($field in @('version','platform','method','date','sha256')) { if ([string]::IsNullOrWhiteSpace([string]$fixture.$field)) { throw "M0012 fixture provenance lacks '$field': $path" } }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne ([string]$fixture.sha256).ToUpperInvariant()) { throw "M0012 fixture hash mismatch: $path" }
}
Write-Output 'M0012 fixture provenance and hashes verified.'
