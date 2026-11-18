param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0020'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = [IO.Path]::GetFullPath($EvidenceRoot)
$packageDirectory = Join-Path $root 'consumer-package'
$consumerRoot = Join-Path $root ('consumer-workspace-' + [guid]::NewGuid().ToString('N'))
$toolPath = Join-Path $root ('consumer-tool-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $root,$packageDirectory | Out-Null
& (Join-Path $PSScriptRoot 'pack.ps1') -OutputPath $packageDirectory
if ($LASTEXITCODE -ne 0) { throw 'M0020 current tool pack failed.' }
$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
$package = Join-Path $packageDirectory "Yadg.$version.nupkg"
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw 'M0020 packed Yadg package is missing.' }
& dotnet tool install Yadg --version $version --tool-path $toolPath --add-source $packageDirectory --ignore-failed-sources
if ($LASTEXITCODE -ne 0) { throw 'M0020 isolated installed-tool installation failed.' }
$yadg = Join-Path $toolPath 'yadg.exe'
if (-not (Test-Path -LiteralPath $yadg -PathType Leaf)) { throw 'M0020 isolated tool shim is missing.' }

if (Test-Path -LiteralPath $consumerRoot) { throw "Refusing to reuse an existing M0020 consumer workspace: $consumerRoot" }
& $yadg init --workspace $consumerRoot
if ($LASTEXITCODE -ne 0) { throw 'Installed yadg init failed.' }
$templates = Join-Path $consumerRoot 'YadgTemplates'; New-Item -ItemType Directory -Force -Path $templates | Out-Null
$slides = Join-Path $consumerRoot 'slides'; New-Item -ItemType Directory -Force -Path $slides | Out-Null
$wordTemplate = Join-Path $templates 'report.docx'
Copy-Item (Join-Path $repo 'tests/fixtures/m0010/word-origin.docx') $wordTemplate
Set-Content -LiteralPath (Join-Path $consumerRoot 'content.md') -Encoding utf8 -Value @('# Architecture {#architecture}','','Consumer workspace DOCX compatibility sidecheck.')
$repoSource = Join-Path $root 'box-diagram-source.md'
if (-not (Test-Path -LiteralPath $repoSource -PathType Leaf)) { throw 'M0020 Tier 3 source is needed to exercise the same accepted sample.' }
$workspaceSource = Get-Content -LiteralPath $repoSource -Raw
$workspaceSource = $workspaceSource.Replace('```diagram {#m0020-box caption="Nested service diagram"}', '```text {#m0020-box caption="Nested service diagram"}')
Set-Content -LiteralPath (Join-Path $slides 'diagram.md') -Encoding utf8 -Value $workspaceSource
$powerPointTemplate = Join-Path $repo 'tests/fixtures/m0018/powerpoint-origin.pptx'
$impressTemplate = Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx'
Copy-Item $powerPointTemplate (Join-Path $templates 'powerpoint-origin.pptx')
Copy-Item $impressTemplate (Join-Path $templates 'impress-origin.pptx')
$sourceHash = (Get-FileHash (Join-Path $slides 'diagram.md') -Algorithm SHA256).Hash.ToUpperInvariant()
Set-Content -LiteralPath (Join-Path $consumerRoot 'YADG.md') -Encoding utf8 -Value @(
    '---','yadg:','  version: 1','values:','  story-value: "Installed M0020 consumer"','diagrams:','  autoDetect: true','presentations:',
    '  - id: powerpoint-svg','    template: YadgTemplates/powerpoint-origin.pptx','    sources:','      - slides/diagram.md','    output: powerpoint-svg.pptx',
    '  - id: impress-svg','    template: YadgTemplates/impress-origin.pptx','    sources:','      - slides/diagram.md','    output: impress-svg.pptx','---'
)
$initOutput = & $yadg init --workspace $consumerRoot 2>&1
if ($LASTEXITCODE -eq 0) { throw 'Second init unexpectedly overwrote the bootstrapped workspace.' }
$workspaceCheck = & $yadg check --workspace $consumerRoot 2>&1
if ($LASTEXITCODE -ne 0) { throw "Installed workspace check failed: $workspaceCheck" }
& $yadg presentation check --id powerpoint-svg --workspace $consumerRoot
if ($LASTEXITCODE -ne 0) { throw 'Installed selected-deck check failed.' }
& $yadg presentation build --id powerpoint-svg --workspace $consumerRoot
if ($LASTEXITCODE -ne 0) { throw 'Installed selected-deck build failed.' }
& $yadg build --workspace $consumerRoot --yolo
if ($LASTEXITCODE -ne 0) { throw 'Installed whole-workspace YOLO build failed.' }

$directConfig = Join-Path $root 'consumer-direct-config.yaml'
Set-Content -LiteralPath $directConfig -Encoding utf8 -Value @('yadg:','  version: 1','diagrams:','  autoDetect: true')
$directOutput = Join-Path $root 'direct-config-output.pptx'
$ambient = Join-Path $consumerRoot 'YADG.md'
$saved = Get-Content -LiteralPath $ambient -Raw
Push-Location $consumerRoot
try {
    $direct = & $yadg presentation build --template $powerPointTemplate --source (Join-Path $slides 'diagram.md') --output $directOutput --config $directConfig 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Installed direct/config presentation build failed: $direct" }
    Set-Content -LiteralPath $ambient -Encoding utf8 -Value 'not: [valid'
    $isolatedOutput = Join-Path $root 'direct-isolation-output.pptx'
    $isolated = & $yadg presentation build --template $powerPointTemplate --source (Join-Path $slides 'diagram.md') --output $isolatedOutput 2>&1
    $isolatedText = $isolated -join "`n"
    if ($LASTEXITCODE -eq 0 -or $isolatedText -notmatch 'YADG-PPT-MD-' -or $isolatedText -match 'YADG-CONFIG-') { throw "Direct mode unexpectedly read ambient workspace config or enabled diagram detection by default: $isolatedText" }
} finally { Set-Content -LiteralPath $ambient -Encoding utf8 -Value $saved; Pop-Location }

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Check-SvgPackage([string] $path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $svg = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/media/.*\.svg$' })
        if ($svg.Count -eq 0) { throw "Installed output lacks SVG vector media: $path" }
        $slideXml = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | ForEach-Object { $reader=[IO.StreamReader]::new($_.Open()); try {$reader.ReadToEnd()} finally {$reader.Dispose()} }) -join "`n"
        if ($slideXml -notmatch 'asvg:svgBlip' -or $slideXml -notmatch '<p:pic') { throw "Installed output does not reference SVG through a native picture: $path" }
        foreach ($entry in $svg) { $reader=[IO.StreamReader]::new($entry.Open()); try {$text=$reader.ReadToEnd()} finally {$reader.Dispose()}; foreach ($label in @('Outer','Inner')) { if ($text -notmatch [regex]::Escape($label)) { throw "SVG omitted documented diagram label '$label'." } } }
        return [ordered]@{ path=[IO.Path]::GetRelativePath($repo,$path).Replace('\','/'); sha256=(Get-FileHash $path -Algorithm SHA256).Hash.ToUpperInvariant(); svgParts=$svg.Count }
    } finally { $zip.Dispose() }
}
$workspaceOutputRoot = Join-Path $consumerRoot 'YadgPresentations'
$powerPointDeck = Check-SvgPackage (Join-Path $workspaceOutputRoot 'powerpoint-svg.pptx')
$impressDeck = Check-SvgPackage (Join-Path $workspaceOutputRoot 'impress-svg.pptx')
$directDeck = Check-SvgPackage $directOutput
$nativePowerPointReceipt = Join-Path $root 'tier4-powerpoint-native.json'
$nativeImpressReceipt = Join-Path $root 'tier4-impress-native.json'
$source = Join-Path $slides 'diagram.md'
Copy-Item (Join-Path $workspaceOutputRoot 'powerpoint-svg.pptx') (Join-Path $root 'tier4-powerpoint-installed.pptx') -Force
& (Join-Path $PSScriptRoot 'test-m0018-powerpoint-tier3.ps1') -TemplatePath $powerPointTemplate -SourcePath $source -OutputPath (Join-Path $root 'tier4-powerpoint-installed.pptx') -EvidencePath $nativePowerPointReceipt -SkipBuild -Milestone M0020 -Gate VAL-04
if (-not $?) { throw 'Installed-tool PowerPoint candidate failed native validation.' }
Copy-Item (Join-Path $workspaceOutputRoot 'impress-svg.pptx') (Join-Path $root 'tier4-impress-installed.pptx') -Force
& (Join-Path $PSScriptRoot 'test-m0018-tier3.ps1') -TemplatePath $impressTemplate -SourcePath $source -CandidatePath (Join-Path $root 'tier4-impress-installed.pptx') -EvidencePath $nativeImpressReceipt -ConfigPath $directConfig -SkipBuild -AutomatedReviewOnly -Milestone M0020 -Gate VAL-04
if (-not $?) { throw 'Installed-tool Impress candidate failed native validation.' }
$packageSha = (Get-FileHash $package -Algorithm SHA256).Hash.ToUpperInvariant()
$evidence = [ordered]@{
    milestone='M0020'; gate='VAL-04'; result='passed'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim()
    packageId='Yadg'; packageVersion=$version; packagePath=[IO.Path]::GetRelativePath($repo,$package).Replace('\','/'); packageSha256=$packageSha
    installPath=[IO.Path]::GetRelativePath($repo,$toolPath).Replace('\','/'); installMode='isolated dotnet tool --tool-path'
    init='passed; second init preserved workspace'; workspaceCheck='passed'; selectedDeck='passed'; workspaceBuild='passed'; directWithConfig='passed'; directAmbientIsolation='passed'
    sourceSha256=$sourceHash; directConfigSha256=(Get-FileHash $directConfig -Algorithm SHA256).Hash.ToUpperInvariant()
    installedOutputs=@($powerPointDeck,$impressDeck,$directDeck)
    nativeReceipts=@(
        [ordered]@{path='artifacts/review/evidence/M0020/tier4-powerpoint-native.json';sha256=(Get-FileHash $nativePowerPointReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); candidate='artifacts/review/evidence/M0020/tier4-powerpoint-installed.pptx';candidateSha256=(Get-FileHash (Join-Path $root 'tier4-powerpoint-installed.pptx') -Algorithm SHA256).Hash.ToUpperInvariant()},
        [ordered]@{path='artifacts/review/evidence/M0020/tier4-impress-native.json';sha256=(Get-FileHash $nativeImpressReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); candidate='artifacts/review/evidence/M0020/tier4-impress-installed.pptx';candidateSha256=(Get-FileHash (Join-Path $root 'tier4-impress-installed.pptx') -Algorithm SHA256).Hash.ToUpperInvariant()}
    )
}
$evidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'tier4-evidence.json') -Encoding utf8
Write-Output 'M0020 isolated installed-tool consumer checks passed with native PowerPoint and Impress acceptance.'
