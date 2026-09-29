param(
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/package')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $repo 'Yadg.slnx'
$project = Join-Path $repo 'src/Yadg.Cli/Yadg.Cli.csproj'
$propsPath = Join-Path $repo 'Directory.Build.props'

if (-not (Test-Path -LiteralPath $solution)) { throw "Solution is missing: $solution" }
if (-not (Test-Path -LiteralPath $project)) { throw "CLI project is missing: $project" }
$props = [xml](Get-Content -LiteralPath $propsPath -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Central VersionPrefix must be a stable semantic version; found '$version'." }
$expected = Join-Path ([System.IO.Path]::GetFullPath($OutputPath)) "Yadg.$version.nupkg"

$output = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path $output | Out-Null
Get-ChildItem -LiteralPath $output -Filter 'Yadg.*.nupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $output -Filter 'Yadg.*.snupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force

& dotnet restore $solution
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed with exit code $LASTEXITCODE." }
& dotnet build $solution --configuration Release --no-restore
if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE." }
& dotnet pack $project --configuration Release --no-build --output $output
if ($LASTEXITCODE -ne 0) { throw "Pack failed with exit code $LASTEXITCODE." }

$packages = @(Get-ChildItem -LiteralPath $output -Filter '*.nupkg' -File)
if ($packages.Count -ne 1 -or $packages[0].FullName -ne $expected) {
    $names = ($packages | ForEach-Object Name) -join ', '
    throw "Pack must produce exactly one Yadg.$version.nupkg in '$output'; found: $names"
}
Write-Output $expected
