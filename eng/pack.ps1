param(
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/package')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$project = Join-Path $repo 'src/Yadg.Cli/Yadg.Cli.csproj'
$expected = Join-Path ([System.IO.Path]::GetFullPath($OutputPath)) 'Yadg.1.0.0.nupkg'

if (-not (Test-Path -LiteralPath $project)) { throw "CLI project is missing: $project" }
$props = Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw
if ($props -notmatch '<VersionPrefix>\s*1\.0\.0\s*</VersionPrefix>') { throw 'Central VersionPrefix must be 1.0.0.' }

$candidates = @(
    'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe'
)
$msbuild = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $msbuild) {
    throw 'Full Visual Studio MSBuild with ResolveComReference is required; no supported MSBuild.exe was found. dotnet pack is not an allowed fallback.'
}

$output = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path $output | Out-Null
Get-ChildItem -LiteralPath $output -Filter 'Yadg.1.0.0.nupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $output -Filter 'Yadg.1.0.0.snupkg' -File -ErrorAction SilentlyContinue | Remove-Item -Force

& dotnet restore (Join-Path $repo 'Yadg.slnx')
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed with exit code $LASTEXITCODE." }
& $msbuild $project /t:Pack /p:Configuration=Release /p:Restore=false /p:PackageOutputPath=$output /p:NoPackageAnalysis=true /v:minimal
if ($LASTEXITCODE -ne 0) { throw "Full-MSBuild pack failed with exit code $LASTEXITCODE." }

$packages = @(Get-ChildItem -LiteralPath $output -Filter '*.nupkg' -File)
if ($packages.Count -ne 1 -or $packages[0].FullName -ne $expected) {
    $names = ($packages | ForEach-Object Name) -join ', '
    throw "Pack must produce exactly one Yadg.1.0.0.nupkg in '$output'; found: $names"
}

# The Windows implementation TFM is retained in the runtimeconfig and build
# output. The dotnet tool installer requires its implementation directory to use
# the platform-neutral tool TFM layout, so normalize only that package path
# after the full-MSBuild/COMReference pack has completed.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$normalized = "$expected.normalized"
$inputZip = [System.IO.Compression.ZipFile]::OpenRead($expected)
$outputZip = [System.IO.Compression.ZipFile]::Open($normalized, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($entry in $inputZip.Entries) {
        $name = $entry.FullName -replace '^tools/net10\.0-windows7\.0/', 'tools/net10.0/'
        $newEntry = $outputZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $inputStream = $entry.Open(); $outputStream = $newEntry.Open()
        try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose(); $inputStream.Dispose() }
    }
}
finally { $outputZip.Dispose(); $inputZip.Dispose() }
Move-Item -LiteralPath $normalized -Destination $expected -Force
Write-Output $expected
