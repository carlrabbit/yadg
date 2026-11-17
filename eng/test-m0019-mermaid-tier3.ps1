param(
    [string] $CacheRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'yadg-m0008-tools')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$manifest = Get-Content (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
$bunVersion = [string]$manifest.bun.version
$toolDir = Join-Path $CacheRoot "bun-$bunVersion-windows-x64"
$bun = Join-Path $toolDir 'bun.exe'
if (-not (Test-Path -LiteralPath $bun -PathType Leaf)) {
    & (Join-Path $PSScriptRoot 'test-m0008-tier3.ps1') -CacheRoot $CacheRoot
    if ($LASTEXITCODE -ne 0) { throw 'The established pinned Bun/Mermaid setup failed.' }
}
if ((& $bun --version).Trim() -ne $bunVersion) { throw "M0019 Tier 3 requires pinned Bun $bunVersion." }

$globalPath = Join-Path $repo 'global.json'
if (Test-Path -LiteralPath $globalPath) { throw 'Refusing to replace an existing repository global.json.' }
Set-Content -LiteralPath $globalPath -Value '{"sdk":{"version":"10.0.403","rollForward":"disable","allowPrerelease":false}}' -Encoding ascii
$oldBun = $env:YADG_M0019_BUN
$oldReceipt = $env:YADG_M0019_PRODUCER_RECEIPT
$oldRevision = $env:YADG_M0019_REVISION
$oldBunVersion = $env:YADG_M0019_BUN_VERSION
$evidenceRoot = Join-Path $repo 'artifacts/review/evidence/M0019'
New-Item -ItemType Directory -Force -Path $evidenceRoot | Out-Null
$receiptPath = Join-Path $evidenceRoot 'mermaid-tier3-evidence.json'
try {
    $env:YADG_M0019_BUN = $bun
    $env:YADG_M0019_BUN_VERSION = $bunVersion
    $env:YADG_M0019_REVISION = (& git -C $repo rev-parse HEAD).Trim()
    $env:YADG_M0019_PRODUCER_RECEIPT = $receiptPath
    & dotnet test (Join-Path $repo 'tests/Yadg.IntegrationTests/Yadg.IntegrationTests.csproj') --no-restore --filter 'FullyQualifiedName~Tier3_real_bun_mermaid_producer_creates_native_pptx_picture'
    if ($LASTEXITCODE -ne 0) { throw 'M0019 real Mermaid PPTX producer integration failed.' }
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'M0019 Tier 3 test did not produce its automated evidence receipt.' }
    Write-Output "M0019 Tier 3 Mermaid passed: Bun $bunVersion, $($manifest.mermaidCli.package)@$($manifest.mermaidCli.version)."
} finally {
    if ($null -eq $oldBun) { Remove-Item Env:YADG_M0019_BUN -ErrorAction SilentlyContinue } else { $env:YADG_M0019_BUN = $oldBun }
    if ($null -eq $oldReceipt) { Remove-Item Env:YADG_M0019_PRODUCER_RECEIPT -ErrorAction SilentlyContinue } else { $env:YADG_M0019_PRODUCER_RECEIPT = $oldReceipt }
    if ($null -eq $oldRevision) { Remove-Item Env:YADG_M0019_REVISION -ErrorAction SilentlyContinue } else { $env:YADG_M0019_REVISION = $oldRevision }
    if ($null -eq $oldBunVersion) { Remove-Item Env:YADG_M0019_BUN_VERSION -ErrorAction SilentlyContinue } else { $env:YADG_M0019_BUN_VERSION = $oldBunVersion }
    Remove-Item -LiteralPath $globalPath -Force -ErrorAction SilentlyContinue
}
