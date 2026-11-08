param(
    [Parameter(Mandatory=$true)][string] $TemplatePath,
    [Parameter(Mandatory=$true)][string[]] $SourcePath,
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$template = [IO.Path]::GetFullPath($TemplatePath)
$sources = @($SourcePath | ForEach-Object { [IO.Path]::GetFullPath($_) })
if (-not (Test-Path -LiteralPath $template -PathType Leaf) -or $sources.Count -eq 0 -or @($sources | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) { throw 'M0018 Tier 4 requires an existing PPTX prototype and one or more Markdown sources.' }
$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
$installedSdks = @(& dotnet --list-sdks)
$sdkVersion = $installedSdks | ForEach-Object { if ($_ -match '^\s*(10\.0\.\d+)\s') { $Matches[1] } } | Sort-Object { [version]$_ } -Descending | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace([string]$sdkVersion)) { throw 'M0018 Tier 4 requires an installed .NET 10 SDK.' }
$evidenceDirectory = [IO.Path]::GetFullPath($EvidenceRoot)
New-Item -ItemType Directory -Force -Path $evidenceDirectory | Out-Null
$evidenceTemplate = Join-Path $evidenceDirectory 'consumer-template.pptx'
Copy-Item -LiteralPath $template -Destination $evidenceTemplate -Force
$evidenceSources = @()
for ($index = 0; $index -lt $sources.Count; $index++) {
    $evidenceSource = Join-Path $evidenceDirectory ('consumer-source-{0:D2}.md' -f ($index + 1))
    Copy-Item -LiteralPath $sources[$index] -Destination $evidenceSource -Force
    $evidenceSources += $evidenceSource
}
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0018-tier4-' + [guid]::NewGuid().ToString('N'))
$packageDirectory = Join-Path $scratch 'package'
$toolDirectory = Join-Path $scratch 'tool'
$workspace = Join-Path $scratch 'consumer'
New-Item -ItemType Directory -Force -Path $packageDirectory, $toolDirectory, $workspace | Out-Null
$globalConfig = [ordered]@{ sdk = [ordered]@{ version = $sdkVersion; rollForward = 'disable'; allowPrerelease = $false } }
$globalConfig | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $scratch 'global.json') -Encoding utf8
Push-Location -LiteralPath $scratch
try {
    & dotnet pack (Join-Path $repo 'src/Yadg.Cli/Yadg.Cli.csproj') --configuration Release --output $packageDirectory
    if ($LASTEXITCODE -ne 0) { throw 'M0018 current-tool package build failed.' }
    $package = Join-Path $packageDirectory "Yadg.$version.nupkg"
    if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw "Expected current package not found: $package" }
    & dotnet tool install --tool-path $toolDirectory --add-source $packageDirectory --ignore-failed-sources Yadg --version $version
    if ($LASTEXITCODE -ne 0) { throw 'M0018 isolated current-tool installation failed.' }
    $shim = Join-Path $toolDirectory 'yadg.exe'
    $checkArgs = @('presentation','check','--template',$evidenceTemplate)
    foreach ($source in $evidenceSources) { $checkArgs += @('--source',$source) }
    & $shim @checkArgs
    if ($LASTEXITCODE -ne 0) { throw 'M0018 isolated packaged presentation check failed.' }
    $candidate = Join-Path $workspace 'consumer-sample.pptx'
    $buildArgs = @('presentation','build','--template',$evidenceTemplate)
    foreach ($source in $evidenceSources) { $buildArgs += @('--source',$source) }
    $buildArgs += @('--output',$candidate)
    & $shim @buildArgs
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw 'M0018 isolated packaged presentation build failed.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($candidate)
    try {
        $slideCount = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' }).Count
        if ($slideCount -lt 1) { throw 'Packaged consumer output contains no PPTX slide parts.' }
    } finally { $zip.Dispose() }
    $artifact = Join-Path $evidenceDirectory 'consumer-sample.pptx'
    Copy-Item -LiteralPath $candidate -Destination $artifact -Force
    $workingTreeStatus = (& git -C $repo status --porcelain --untracked-files=all) -join "`n"
    $workingTreeStatusSha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($workingTreeStatus)))
    $record = [ordered]@{
        milestone = 'M0018'
        repositoryRevision = (& git -C $repo rev-parse HEAD).Trim()
        workingTreeStatusSha256 = $workingTreeStatusSha256
        packageId = 'Yadg'
        packageVersion = $version
        sdkVersion = (& dotnet --version).Trim()
        packageSha256 = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToUpperInvariant()
        isolatedToolPath = 'temporary dotnet tool path'
        installedCommand = $shim
        templatePath = [IO.Path]::GetRelativePath($repo, $evidenceTemplate)
        templateSha256 = (Get-FileHash -LiteralPath $evidenceTemplate -Algorithm SHA256).Hash.ToUpperInvariant()
        sourcePaths = @($evidenceSources | ForEach-Object { [IO.Path]::GetRelativePath($repo, $_) })
        sourceSha256 = @($evidenceSources | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToUpperInvariant() })
        outputPath = [IO.Path]::GetRelativePath($repo, $artifact)
        outputSha256 = (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToUpperInvariant()
        outputSlideCount = $slideCount
        result = 'passed'
    }
    $record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidenceDirectory 'tier4-evidence.json') -Encoding utf8
    Write-Output "M0018 isolated packaged consumer passed for Yadg $version."
} finally {
    Pop-Location
    Set-Location -LiteralPath $repo
    [Environment]::CurrentDirectory = $repo
    if (Test-Path -LiteralPath $scratch) {
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
        $resolvedScratch = [IO.Path]::GetFullPath($scratch)
        if (-not $resolvedScratch.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedScratch) -notmatch '^yadg-m0018-tier4-[0-9a-f]{32}$') { throw 'Refusing to remove scratch files outside the expected M0018 temporary path.' }
        try { Remove-Item -LiteralPath $resolvedScratch -Recurse -Force }
        catch { Write-Warning "Unable to remove temporary Tier 4 workspace '$resolvedScratch': $($_.Exception.Message)" }
    }
}
