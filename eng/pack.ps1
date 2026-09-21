param(
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/package')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $repo 'Yadg.slnx'
$project = Join-Path $repo 'src/Yadg.Cli/Yadg.Cli.csproj'
$expected = Join-Path ([System.IO.Path]::GetFullPath($OutputPath)) 'Yadg.1.0.0.nupkg'

if (-not (Test-Path -LiteralPath $solution)) { throw "Solution is missing: $solution" }
if (-not (Test-Path -LiteralPath $project)) { throw "CLI project is missing: $project" }
$props = Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw
if ($props -notmatch '<VersionPrefix>\s*1\.0\.0\s*</VersionPrefix>') { throw 'Central VersionPrefix must be 1.0.0.' }

$output = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path $output | Out-Null
Get-ChildItem -LiteralPath $output -Filter 'Yadg.1.0.0.nupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $output -Filter 'Yadg.1.0.0.snupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force

& dotnet restore $solution
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed with exit code $LASTEXITCODE." }
& dotnet build $solution --configuration Release --no-restore
if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE." }
& dotnet pack $project --configuration Release --no-build --output $output
if ($LASTEXITCODE -ne 0) { throw "Pack failed with exit code $LASTEXITCODE." }

$packages = @(Get-ChildItem -LiteralPath $output -Filter '*.nupkg' -File)
if ($packages.Count -ne 1 -or $packages[0].FullName -ne $expected) {
    $names = ($packages | ForEach-Object Name) -join ', '
    throw "Pack must produce exactly one Yadg.1.0.0.nupkg in '$output'; found: $names"
}
Write-Output $expected
