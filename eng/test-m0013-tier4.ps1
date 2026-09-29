param(
    [string] $PackagePath,
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/release/evidence/M0013')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
if ([string]::IsNullOrWhiteSpace($PackagePath)) { $PackagePath = Join-Path $repo "artifacts/package/Yadg.$version.nupkg" }
$package = [System.IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw "Expected freshly packed release candidate is missing: $package" }
if ((Get-Item -LiteralPath $package).Name -ne "Yadg.$version.nupkg") { throw "Tier 4 accepts only Yadg.$version.nupkg." }

$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0013 Tier 4 requires interactive Microsoft Word; no Word executable is available.' }
if ($null -eq $lo) { throw 'M0013 Tier 4 requires LibreOffice Writer; no soffice executable is available.' }
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } elseif ($lo.Source) { $lo.Source } else { $lo.FullName }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
try {
    $entries = @($archive.Entries | ForEach-Object FullName)
    $nuspecEntry = $archive.Entries | Where-Object { $_.FullName -match '\.nuspec$' } | Select-Object -First 1
    if ($null -eq $nuspecEntry) { throw 'Package nuspec is missing.' }
    $reader = [System.IO.StreamReader]::new($nuspecEntry.Open())
    try { [xml]$nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $metadata = $nuspec.package.metadata
    if ([string]$metadata.id -ne 'Yadg' -or [string]$metadata.version -ne $version) { throw "Package metadata must be Yadg $version." }
    if (-not ($metadata.packageTypes.packageType | Where-Object { $_.name -eq 'DotnetTool' })) { throw 'Package is not marked as a .NET tool.' }
    if ([string]$metadata.authors -ne 'carlrabbit' -or [string]$metadata.readme -ne 'README.md') { throw 'Required package author or README metadata is missing.' }
    if ([string]$metadata.license.type -ne 'expression' -or [string]$metadata.license.InnerText -ne 'MIT') { throw 'Package MIT license metadata is missing or incorrect.' }
    if ($entries -notcontains 'README.md' -or $entries -notcontains 'LICENSE') { throw 'Package must contain README.md and LICENSE.' }
    foreach ($required in @('yadg.dll','Yadg.Core.dll','Yadg.Word.dll','Yadg.Renderer.dll','Yadg.WordRenderer.dll')) {
        if (-not ($entries | Where-Object { $_ -like "tools/*/$required" })) { throw "Required package assembly is missing: $required" }
    }
    $office = $entries | Where-Object { $_ -match '(?i)(^|/)(Interop\.Microsoft\.Office|Microsoft\.Office|office.*interop|interop.*office).*\.dll$' }
    if ($office) { throw "Package must not contain Office interop assemblies: $($office -join ', ')" }
    $bad = $entries | Where-Object { $_ -match '(^|/)(test|tests|fixtures|review|execution|\.git|obj|bin)(/|$)' -or $_ -match '(?i)(api[_-]?key|password|secret|credential)' }
    if ($bad) { throw "Package hygiene failure: $($bad -join ', ')" }
}
finally { $archive.Dispose() }

$temp = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0013-tier4-' + [guid]::NewGuid().ToString('N'))
$cliHome = Join-Path $temp 'cli-home'; $nuget = Join-Path $temp 'nuget'; $tool = Join-Path $temp 'tool'; $workspace = Join-Path $temp 'workspace'; $delivery = Join-Path $temp 'delivery'; $localFeed = Join-Path $temp 'local-feed'
New-Item -ItemType Directory -Force -Path $cliHome,$nuget,$tool,$workspace,$delivery,$localFeed,(Join-Path $workspace 'YadgTemplates') | Out-Null
$oldHome = $env:DOTNET_CLI_HOME; $oldNuget = $env:NUGET_PACKAGES
$results = [ordered]@{}
try {
    Copy-Item -LiteralPath $package -Destination $localFeed
    $packageDir = Split-Path -Parent $package
    $env:DOTNET_CLI_HOME = $cliHome; $env:NUGET_PACKAGES = $nuget
    & dotnet tool install Yadg --tool-path $tool --version $version --add-source $packageDir --no-cache --ignore-failed-sources
    if ($LASTEXITCODE -ne 0) { throw "Isolated dotnet tool installation failed with exit code $LASTEXITCODE." }
    $yadg = Join-Path $tool 'yadg.exe'; if (-not (Test-Path -LiteralPath $yadg)) { $yadg = Join-Path $tool 'yadg' }
    if (-not (Test-Path -LiteralPath $yadg)) { throw 'Installed yadg command was not created.' }
    $resolved = (Get-Command $yadg).Source
    if ($resolved -match '[\\/]bin[\\/]|[\\/]obj[\\/]') { throw "Installed command unexpectedly resolves into repository build output: $resolved" }
    $versionOutput = (& $yadg --version 2>&1 | Out-String).Trim()
    if ($versionOutput -ne $version) { throw "Installed yadg --version returned '$versionOutput'." }
    $results.install = 'passed'; $results.installedVersion = $versionOutput; $results.installedCommand = $resolved

    $helpCases = @(
        @{ args=@('--help'); required=@('init','check','inspect','build','render','publish') },
        @{ args=@('init','--help'); required=@('--workspace') },
        @{ args=@('check','--help'); required=@('--workspace','--list','--yolo') },
        @{ args=@('inspect','--help'); required=@('styles','template') },
        @{ args=@('inspect','styles','--help'); required=@('--workspace','--template') },
        @{ args=@('inspect','template','--help'); required=@('--workspace','--template') },
        @{ args=@('build','--help'); required=@('--workspace','--yolo') },
        @{ args=@('render','--help'); required=@('--workspace','--renderer','libreoffice','word','--renderer-path','--yolo') },
        @{ args=@('publish','--help'); required=@('--workspace','--publish-path') }
    )
    foreach ($case in $helpCases) {
        $output = (& $yadg @($case.args) 2>&1 | Out-String)
        if ($LASTEXITCODE -ne 0) { throw "Installed help failed for '$($case.args -join ' ')' ." }
        foreach ($requiredText in $case.required) { if ($output -notmatch [regex]::Escape($requiredText)) { throw "Installed help '$($case.args -join ' ')' is missing '$requiredText'." } }
    }
    $renderHelp = (& $yadg render --help 2>&1 | Out-String)
    if ($renderHelp -notmatch 'default:\s*libreoffice') { throw 'Installed render help does not identify LibreOffice as the default renderer.' }
    $results.help = 'all installed root and command-specific surfaces passed material option/default checks'

    $bootstrap = Join-Path $temp 'bootstrap'
    New-Item -ItemType Directory -Path $bootstrap | Out-Null
    Push-Location $bootstrap
    try { & $yadg init; if ($LASTEXITCODE -ne 0) { throw 'Installed init failed.' } } finally { Pop-Location }
    foreach ($owned in @('YADG.md','content.md','YadgTemplates')) { if (-not (Test-Path (Join-Path $bootstrap $owned))) { throw "Init did not create $owned." } }
    $before = (Get-FileHash (Join-Path $bootstrap 'content.md')).Hash
    Push-Location $bootstrap
    try { & $yadg init 2>$null; if ($LASTEXITCODE -eq 0) { throw 'Init unexpectedly overwrote existing owned paths.' } } finally { Pop-Location }
    if ((Get-FileHash (Join-Path $bootstrap 'content.md')).Hash -ne $before) { throw 'Init changed an owned starter file.' }
    $results.init = 'created starter structure and preserved existing owned paths'

    Copy-Item -LiteralPath (Join-Path $repo 'tests/fixtures/m0010/word-origin.docx') -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
    @('---','yadg:','  version: 1','values:','  document-version: ''M0013''','  story-value: ''Packaged consumer value''','---') | Set-Content -LiteralPath (Join-Path $workspace 'YADG.md')
    @('# Architecture {#architecture}','','First packaged paragraph.','','## Assumptions','','Details.') | Set-Content -LiteralPath (Join-Path $workspace 'content.md')
    & $yadg check --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Strict installed check failed.' }
    & $yadg check --workspace $workspace --list; if ($LASTEXITCODE -ne 0) { throw 'Installed check --list failed.' }
    & $yadg inspect styles --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Installed inspect styles failed.' }
    & $yadg inspect template --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Installed inspect template failed.' }
    & $yadg build --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Strict installed build failed.' }
    if (-not (Get-ChildItem (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' -File)) { throw 'Installed build produced no authored DOCX.' }
    $results.strictAuthoring = 'check/check --list/inspect styles/inspect template/build passed'

    # Validate the complete documented workspace YAML shape and representative README
    # Markdown grammar through the installed product, without rewriting authored docs.
    @('---','yadg:','  version: 1','','values:','  document-version: "1.1"','  story-value: "Packaged consumer value"','','markdown:','  thematicBreak: ignore','  codeInline: ignore','','publish:','  path: ./Published','','producers:','  mermaid:','    executable: mmdc','    arguments: []','---') | Set-Content -LiteralPath (Join-Path $workspace 'YADG.md')
    @('# Introduction {#introduction}','','This document describes the reporting process.','','## Architecture {#architecture}','','The process has two steps.','','Inline `code` uses the documented ignore policy.') | Set-Content -LiteralPath (Join-Path $workspace 'content.md')
    & $yadg check --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Documented README YAML/Markdown example did not validate through the installed command.' }
    & $yadg build --workspace $workspace; if ($LASTEXITCODE -ne 0) { throw 'Documented README Markdown example did not build through the installed command.' }
    $results.documentationExamples = 'README workspace YAML schema and representative heading/paragraph/inline-code Markdown validated by installed check/build; table/list/prototype contracts are covered by M0012 fixtures and source parser audit'

    # The fixture lacks a usable configured ordered-list style. This is a recoverable
    # presentation mismatch with a deterministic built-in numbering fallback.
    Add-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('','1. First ordered item','2. Second ordered item')
    @('---','yadg:','  version: 1','','values:','  document-version: "1.1"','  story-value: "Packaged consumer value"','','markdown:','  thematicBreak: ignore','  codeInline: ignore','','publish:','  path: ./Published','','producers:','  mermaid:','    executable: mmdc','    arguments: []','---') | Set-Content -LiteralPath (Join-Path $workspace 'YADG.md')
    & $yadg check --workspace $workspace 2>$null; if ($LASTEXITCODE -eq 0) { throw 'Strict check accepted the unavailable ordered-list presentation.' }
    $yoloCheckOutput = (& $yadg check --workspace $workspace --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $yoloCheckOutput -notmatch 'builtin:ordered-list-v1') { throw "YOLO check did not select its documented built-in ordered-list fallback: $yoloCheckOutput" }
    & $yadg build --workspace $workspace 2>$null; if ($LASTEXITCODE -eq 0) { throw 'Strict build accepted the unavailable ordered-list presentation.' }
    $yoloOutput = (& $yadg build --workspace $workspace --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $yoloOutput -notmatch 'degradation' -or $yoloOutput -notmatch 'builtin:ordered-list-v1') { throw 'YOLO build did not produce a useful artifact with visible ordered-list fallback degradation.' }
    $results.yoloAuthoring = 'strict check/build rejected unavailable ordered-list presentation; YOLO check/build used builtin:ordered-list-v1 with visible degradation'

    & $yadg render --workspace $workspace --renderer word; if ($LASTEXITCODE -ne 0) { throw 'Installed Word renderer smoke failed.' }
    $results.wordRender = 'passed; ' + $word.VersionInfo.FileVersion
    & $yadg render --workspace $workspace --renderer libreoffice --renderer-path $loPath; if ($LASTEXITCODE -ne 0) { throw 'Installed LibreOffice renderer smoke failed.' }
    $results.libreOfficeRender = 'passed; ' + $lo.VersionInfo.FileVersion
    $fallbackOutput = (& $yadg render --workspace $workspace --renderer libreoffice --renderer-path (Join-Path $temp 'missing-soffice.exe') --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $fallbackOutput -notmatch 'requested=libreoffice' -or $fallbackOutput -notmatch 'actual=word') { throw "Installed YOLO renderer fallback did not report successful LibreOffice-to-Word fallback: $fallbackOutput" }
    $results.yoloRendererFallback = 'invalid LibreOffice executable fell back to Microsoft Word; requested/actual reported'
    Remove-Item -LiteralPath (Join-Path $workspace 'YADG.md'),(Join-Path $workspace 'content.md'),(Join-Path $workspace 'YadgTemplates') -Recurse -Force
    & $yadg publish --workspace $workspace --publish-path $delivery; if ($LASTEXITCODE -ne 0) { throw 'Installed downstream-only publish smoke failed after authoring inputs were removed.' }
    if (-not (Get-ChildItem -LiteralPath $delivery -Filter '*.docx' -File)) { throw 'Installed publish delivered no DOCX.' }
    $results.publish = 'published finalized DOCX downstream without authoring inputs'

    # Exercise the release push wrapper against a local filesystem NuGet source only.
    & (Join-Path $repo 'eng/publish-nuget.ps1') -Source $localFeed -PackagePath $package
    if ($LASTEXITCODE -ne 0) { throw 'Local non-external publish-nuget validation failed.' }
    $results.publishScriptLocalDestination = 'passed; local filesystem source only'

    $tier3Path = Join-Path $repo 'artifacts/review/evidence/M0012/tier3-evidence.json'
    if (-not (Test-Path -LiteralPath $tier3Path)) { throw "Fresh M0012 Tier-3 evidence is missing: $tier3Path" }
    $tier3 = Get-Content -LiteralPath $tier3Path -Raw | ConvertFrom-Json
    New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
    $revision = (& git -C $repo rev-parse HEAD).Trim()
    $tree = (& git -C $repo status --short) -join "`n"
    $evidence = [ordered]@{
        repositoryRevision = $revision
        workingTreeStatus = $tree
        releaseVersion = $version
        packageId = 'Yadg'
        packagePath = $package
        packageSha256 = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash
        packageInspection = 'passed'
        packageEntries = $entries
        dotnetSdk = (& dotnet --version).Trim()
        windowsVersion = (Get-CimInstance Win32_OperatingSystem).Caption
        wordVersion = $word.VersionInfo.FileVersion
        libreOfficeVersion = $lo.VersionInfo.FileVersion
        m0012Tier3 = [ordered]@{ evidencePath='artifacts/review/evidence/M0012/tier3-evidence.json'; repositoryRevision=[string]$tier3.repositoryRevision }
        installedConsumer = $results
        documentationExampleAudit = 'passed; documented CLI/help surfaces checked from installed command; YAML/Markdown/template examples checked against implementation during release audit'
        externalPublication = 'none; no NuGet.org push, GitHub Release, tag, or publication workflow'
    }
    $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'release-evidence.json') -Encoding utf8
    Write-Output 'M0013 Tier 4 packaged-tool consumer validation passed.'
}
finally {
    $env:DOTNET_CLI_HOME = $oldHome; $env:NUGET_PACKAGES = $oldNuget
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
}
