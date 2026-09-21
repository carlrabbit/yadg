param(
    [Parameter(Mandatory = $true)][string] $Source,
    [string] $PackagePath
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Source)) { throw '-Source must be an explicit non-empty NuGet source.' }
if ([string]::IsNullOrWhiteSpace($PackagePath)) { $PackagePath = Join-Path $repo 'artifacts/package/Yadg.1.0.0.nupkg' }
$package = [System.IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw "Package does not exist: $package" }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
try {
    $nuspecEntry = $archive.Entries | Where-Object { $_.FullName -match '\.nuspec$' } | Select-Object -First 1
    if ($null -eq $nuspecEntry) { throw 'Package does not contain a nuspec manifest.' }
    $reader = New-Object System.IO.StreamReader($nuspecEntry.Open())
    try { [xml]$nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $metadata = $nuspec.package.metadata
    if ([string]$metadata.id -ne 'Yadg' -or [string]$metadata.version -ne '1.0.0') { throw "Expected package Yadg 1.0.0, found $($metadata.id) $($metadata.version)." }
}
finally { $archive.Dispose() }

$nugetArgs = @('nuget', 'push', $package, '--source', $Source)
if (-not [string]::IsNullOrWhiteSpace($env:NUGET_API_KEY)) { $nugetArgs += @('--api-key', $env:NUGET_API_KEY) }
& dotnet @nugetArgs
if ($LASTEXITCODE -ne 0) { throw "NuGet push failed with exit code $LASTEXITCODE." }
Write-Output "Pushed Yadg 1.0.0 to explicit source '$Source'."
