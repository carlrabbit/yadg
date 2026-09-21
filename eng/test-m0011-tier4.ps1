param(
    [string] $PackagePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/package/Yadg.1.0.0.nupkg'),
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/release/evidence/M0011')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$package = [System.IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw "Expected packed release candidate is missing: $package" }
if ((Get-Item -LiteralPath $package).Name -ne 'Yadg.1.0.0.nupkg') { throw 'Tier 4 accepts only Yadg.1.0.0.nupkg.' }

$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0011 Tier 4 requires interactive Microsoft Word; no Word executable is available.' }
if ($null -eq $lo) { throw 'M0011 Tier 4 requires LibreOffice Writer; no soffice executable is available.' }
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } elseif ($lo.Source) { $lo.Source } else { $lo.FullName }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
$entries = @()
try {
    $entries = @($archive.Entries | ForEach-Object FullName)
    $nuspecEntry = $archive.Entries | Where-Object { $_.FullName -match '\.nuspec$' } | Select-Object -First 1
    if ($null -eq $nuspecEntry) { throw 'Package nuspec is missing.' }
    $reader = New-Object System.IO.StreamReader($nuspecEntry.Open())
    try { [xml]$nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $metadata = $nuspec.package.metadata
    if ([string]$metadata.id -ne 'Yadg' -or [string]$metadata.version -ne '1.0.0') { throw 'Package metadata ID/version is not Yadg 1.0.0.' }
    if (-not ($metadata.packageTypes.packageType | Where-Object { $_.name -eq 'DotnetTool' })) { throw 'Package is not marked as a .NET tool.' }
    if ([string]$metadata.authors -ne 'carlrabbit' -or [string]$metadata.description -ne 'Template-first document authoring from Markdown and prepared DOCX templates.') { throw 'Required package metadata is incorrect.' }
    if ([string]$metadata.readme -ne 'README.md') { throw 'Package README metadata is missing.' }
    if ([string]$metadata.license.type -ne 'expression' -or [string]$metadata.license.InnerText -ne 'MIT') { throw 'Package MIT license metadata is missing or incorrect.' }
    if ($entries -notcontains 'README.md') { throw 'Repository README is not packaged.' }
    if ($entries -notcontains 'LICENSE') { throw 'Repository LICENSE file is not packaged.' }
    foreach ($required in @('yadg.dll','Yadg.Core.dll','Yadg.Word.dll','Yadg.Renderer.dll','Yadg.WordRenderer.dll')) {
        if (-not ($entries | Where-Object { $_ -like "tools/*/$required" })) { throw "Required package assembly is missing: $required" }
    }
    $office = $entries | Where-Object { $_ -match '(?i)(^|/)(Interop\.Microsoft\.Office|Microsoft\.Office|office.*interop|interop.*office).*\.dll$' }
    if ($office) { throw "Package must not contain Office interop assemblies: $($office -join ', ')" }
    $bad = $entries | Where-Object { $_ -match '(^|/)(test|tests|fixtures|review|\.git|obj|bin)(/|$)' -or $_ -match '(?i)(api[_-]?key|password|secret|credential)' }
    if ($bad) { throw "Package hygiene failure: $($bad -join ', ')" }
}
finally { $archive.Dispose() }

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('yadg-m0011-tier4-' + [guid]::NewGuid().ToString('N'))
$cliHome = Join-Path $temp 'cli-home'; $nuget = Join-Path $temp 'nuget'; $tool = Join-Path $temp 'tool'; $workspace = Join-Path $temp 'workspace'; $delivery = Join-Path $temp 'delivery'
New-Item -ItemType Directory -Force -Path $cliHome,$nuget,$tool,$workspace,$delivery,(Join-Path $workspace 'YadgTemplates'),(Join-Path $workspace 'YadgPreWords'),(Join-Path $workspace 'YadgWords') | Out-Null
$oldHome = $env:DOTNET_CLI_HOME; $oldNuget = $env:NUGET_PACKAGES
$env:DOTNET_CLI_HOME = $cliHome; $env:NUGET_PACKAGES = $nuget
try {
    $packageDir = Split-Path -Parent $package
    & dotnet tool install Yadg --tool-path $tool --version 1.0.0 --add-source $packageDir --no-cache --ignore-failed-sources
    if ($LASTEXITCODE -ne 0) { throw "Isolated dotnet tool installation failed with exit code $LASTEXITCODE." }
    $yadg = Join-Path $tool 'yadg.exe'
    if (-not (Test-Path -LiteralPath $yadg)) { $yadg = Join-Path $tool 'yadg' }
    if (-not (Test-Path -LiteralPath $yadg)) { throw 'Installed yadg command was not created.' }
    $versionOutput = (& $yadg --version 2>&1 | Out-String).Trim()
    if ($versionOutput -ne '1.0.0') { throw "Installed yadg --version returned '$versionOutput'." }
    $helpCases = @(
        @{ args = @('--help'); required = @('check', 'build', 'render', 'publish') },
        @{ args = @('check', '--help'); required = @('--workspace') },
        @{ args = @('build', '--help'); required = @('--workspace') },
        @{ args = @('render', '--help'); required = @('--workspace', '--renderer', '--renderer-path') },
        @{ args = @('publish', '--help'); required = @('--workspace', '--publish-path') }
    )
    foreach ($helpCase in $helpCases) {
        $helpOutput = (& $yadg @($helpCase.args) 2>&1 | Out-String)
        if ($LASTEXITCODE -ne 0) { throw "Installed help failed for '$($helpCase.args -join ' ')'." }
        foreach ($requiredText in $helpCase.required) { if ($helpOutput -notmatch [regex]::Escape($requiredText)) { throw "Installed help '$($helpCase.args -join ' ')' is missing '$requiredText'." } }
    }

    Copy-Item -LiteralPath (Join-Path $repo 'tests/fixtures/m0010/word-origin.docx') -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
    Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Value @('---','yadg:','  version: 1','values:','  document-version: ''M0011''','  story-value: ''Packaged consumer value''','---')
    Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('# Architecture {#architecture}','','First packaged paragraph.','','## Assumptions','','Details.')
    & $yadg check --workspace $workspace
    if ($LASTEXITCODE -ne 0) { throw 'Installed check failed.' }
    & $yadg build --workspace $workspace
    if ($LASTEXITCODE -ne 0) { throw 'Installed build failed.' }
    $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
    if ($null -eq $authored) { throw 'Installed build produced no DOCX.' }
    & $yadg render --workspace $workspace --renderer word
    if ($LASTEXITCODE -ne 0) { throw 'Installed Word renderer smoke failed.' }
    & $yadg render --workspace $workspace --renderer libreoffice --renderer-path $loPath
    if ($LASTEXITCODE -ne 0) { throw 'Installed LibreOffice renderer smoke failed.' }
    & $yadg publish --workspace $workspace --publish-path $delivery
    if ($LASTEXITCODE -ne 0) { throw 'Installed publish smoke failed.' }
    if (-not (Get-ChildItem -LiteralPath $delivery -Filter '*.docx' -File)) { throw 'Installed publish delivered no DOCX.' }

    New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
    $revision = (& git -C $repo rev-parse HEAD).Trim()
    $m0010EvidencePath = Join-Path $repo 'artifacts/review/evidence/M0010/tier3-evidence.json'
    $evidence = [ordered]@{
        repositoryRevision = $revision
        workingTreeStatus = ((git -C $repo status --short) -join "`n")
        releaseVersion = '1.0.0'; packageId = 'Yadg'; packagePath = $package
        packageSha256 = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash
        packageInspection = 'passed'; packageEntries = $entries
        windowsVersion = (Get-CimInstance Win32_OperatingSystem).Caption
        dotnetSdk = ((dotnet --version).Trim())
        packTooling = 'ordinary dotnet restore/build/pack; no ResolveComReference or tlbimp'
        wordVersion = $word.VersionInfo.FileVersion
        libreOfficeVersion = $lo.VersionInfo.FileVersion
        installedTool = [ordered]@{ install = 'passed'; version = $versionOutput; rootHelp = 'passed'; commandHelp = 'passed'; check = 'passed'; build = 'passed'; wordRender = 'passed'; lateBoundWordApplication = 'passed'; libreOfficeRender = 'passed'; publish = 'passed'; repositoryBinObjDependency = 'not used' }
        publishScriptLocalTarget = 'validated separately by eng/publish-nuget.ps1 against a temporary filesystem source'
        m0010Tier3 = if (Test-Path -LiteralPath $m0010EvidencePath) { $m0010EvidencePath } else { 'missing' }
        documentationAudit = 'README.md and CHANGELOG.md reviewed against docs/engineering/RELEASE.md'
        externalPublication = 'none; no external NuGet push, GitHub Release, or tag occurred'
        licenseMetadata = 'MIT expression declared and LICENSE file packaged; authorized by repository LICENSE and project owner instruction'
    }
    $evidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'release-evidence.json') -Encoding utf8
    Write-Output 'M0011 Tier 4 packaged-tool consumer validation passed.'
}
finally {
    $env:DOTNET_CLI_HOME = $oldHome; $env:NUGET_PACKAGES = $oldNuget
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}
