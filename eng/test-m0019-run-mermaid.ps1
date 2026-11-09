param([string] $InputPath, [string] $OutputPath)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$bun = $env:YADG_M0019_BUN
if ([string]::IsNullOrWhiteSpace($bun) -or -not (Test-Path -LiteralPath $bun -PathType Leaf)) { throw 'M0019 pinned Bun is not configured in YADG_M0019_BUN.' }
$tools = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
& $bun x --bun --package "$($tools.mermaidCli.package)@$($tools.mermaidCli.version)" mmdc -i $InputPath -o $OutputPath
exit $LASTEXITCODE
