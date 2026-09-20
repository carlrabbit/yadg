param(
    [string] $CacheRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'yadg-m0008-tools')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$manifest = Get-Content (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
$version = [string]$manifest.bun.version
$package = [string]$manifest.mermaidCli.package
$packageVersion = [string]$manifest.mermaidCli.version
if ($env:PROCESSOR_ARCHITECTURE -ne 'AMD64' -and $env:PROCESSOR_ARCHITEW6432 -ne 'AMD64') { throw 'M0008 Tier 3 currently supports Windows x64 only.' }

$toolDir = Join-Path $CacheRoot "bun-$version-windows-x64"
$bun = Join-Path $toolDir 'bun.exe'
if (-not (Test-Path -LiteralPath $bun)) {
    New-Item -ItemType Directory -Path $toolDir -Force | Out-Null
    $archive = Join-Path $toolDir "bun-$version.zip"
    $url = "https://github.com/oven-sh/bun/releases/download/bun-v$version/bun-windows-x64.zip"
    Invoke-WebRequest -Uri $url -OutFile $archive
    Expand-Archive -LiteralPath $archive -DestinationPath $toolDir -Force
    $nested = Get-ChildItem -LiteralPath $toolDir -Filter bun.exe -Recurse | Select-Object -First 1
    if ($null -eq $nested) { throw "Downloaded Bun archive did not contain bun.exe: $url" }
    if ($nested.FullName -ne $bun) { Copy-Item -LiteralPath $nested.FullName -Destination $bun -Force }
}
$reported = (& $bun --version).Trim()
if ($reported -ne $version) { throw "Downloaded Bun reported '$reported', expected '$version'." }
Write-Host "Tier 3 provenance: platform=windows-x64 bun=$reported mermaid=$package@$packageVersion"

$probe = Join-Path $toolDir 'probe.mmd'
$probePng = Join-Path $toolDir 'probe.png'
Set-Content -LiteralPath $probe -Value "flowchart LR`n    A --> B" -NoNewline
& $bun x --bun --package "$package@$packageVersion" mmdc -i $probe -o $probePng
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $probePng)) { throw 'Pinned Bun/Mermaid low-level probe failed.' }
Write-Host "Tier 3 low-level probe succeeded: $probePng"

$old = $env:YADG_M0008_BUN
try {
    $env:YADG_M0008_BUN = $bun
    dotnet test (Join-Path $repo 'tests/Yadg.IntegrationTests/Yadg.IntegrationTests.csproj') --no-restore --filter FullyQualifiedName~Tier3_real_bun_mermaid_path_authors_png_and_docx
    if ($LASTEXITCODE -ne 0) { throw 'Tier 3 real YADG Mermaid integration failed.' }
}
finally { $env:YADG_M0008_BUN = $old }
