param(
    [string] $PackagePath,
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/release/evidence/M0015'),
    [string] $MermaidExecutable
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
if ($version -ne '1.2.0') { throw "M0015 Tier 4 requires central release version 1.2.0; found '$version'." }
if ([string]::IsNullOrWhiteSpace($PackagePath)) { $PackagePath = Join-Path $repo "artifacts/package/Yadg.$version.nupkg" }
$package = [IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf) -or (Split-Path -Leaf $package) -ne "Yadg.$version.nupkg") { throw "Expected freshly packed Yadg.$version.nupkg: $package" }

$revision = (& git -C $repo rev-parse HEAD).Trim()
$tree = (& git -C $repo status --porcelain --untracked-files=all) -join "`n"
if (-not [string]::IsNullOrWhiteSpace($tree)) { throw "Tier 4 requires a clean exact release-candidate revision:`n$tree" }
$tier3Path = Join-Path $repo 'artifacts/review/evidence/M0014/tier3-evidence.json'
if (-not (Test-Path -LiteralPath $tier3Path)) { throw 'Fresh M0014 Tier-3 evidence is missing; run ./eng/test-m0014-tier3.ps1 at this revision.' }
$tier3 = Get-Content -LiteralPath $tier3Path -Raw | ConvertFrom-Json
if ([string]$tier3.repositoryRevision -ne $revision) { throw "M0014 Tier-3 evidence revision '$($tier3.repositoryRevision)' does not match candidate '$revision'." }
foreach ($renderer in @('word','libreoffice')) {
    $run = @($tier3.runs | Where-Object { $_.renderer -eq $renderer -and $_.result -eq 'passed' })
    if ($run.Count -ne 1) { throw "M0014 Tier-3 evidence must contain exactly one successful $renderer run." }
    $artifact = Join-Path $repo ([string]$run[0].finalized)
    if (-not (Test-Path -LiteralPath $artifact) -or (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash -ne [string]$run[0].finalizedSha256) { throw "M0014 $renderer finalized artifact does not match its evidence hash." }
}

$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0015 Tier 4 is externally BLOCKED: Microsoft Word is unavailable in this interactive session.' }
if ($null -eq $lo) { throw 'M0015 Tier 4 is externally BLOCKED: LibreOffice Writer is unavailable.' }
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } elseif ($lo.Source) { $lo.Source } else { $lo.FullName }

$tools = Get-Content -LiteralPath (Join-Path $repo 'eng/test-tools.json') -Raw | ConvertFrom-Json
$mermaidPackage = [string]$tools.mermaidCli.package
$mermaidVersion = [string]$tools.mermaidCli.version
if ([string]::IsNullOrWhiteSpace($MermaidExecutable)) {
    $bun = Get-Command bun -ErrorAction SilentlyContinue
    $mmdc = Get-Command mmdc -ErrorAction SilentlyContinue
    if ($null -ne $bun) {
        $MermaidExecutable = $bun.Source
        $mermaidArguments = @('x','--bun','--package',"$mermaidPackage@$mermaidVersion",'mmdc')
        $producerKind = 'Bun x using pinned repository test-tools.json package'
    } elseif ($null -ne $mmdc) {
        $MermaidExecutable = $mmdc.Source
        $mermaidArguments = @()
        $producerKind = 'installed mmdc'
    } else {
        throw "M0015 Tier 4 is externally BLOCKED: install the repository-pinned real Mermaid producer ($mermaidPackage@$mermaidVersion) through an authorized package-manager locus, or pass -MermaidExecutable with a real compliant command."
    }
} else {
    $MermaidExecutable = [IO.Path]::GetFullPath($MermaidExecutable)
    if (-not (Test-Path -LiteralPath $MermaidExecutable -PathType Leaf)) { throw "Mermaid executable does not exist: $MermaidExecutable" }
    if ((Split-Path -Leaf $MermaidExecutable) -match '^bun(?:\.exe)?$') {
        $mermaidArguments = @('x','--bun','--package',"$mermaidPackage@$mermaidVersion",'mmdc')
        $producerKind = 'Bun x using pinned repository test-tools.json package'
    } else {
        $mermaidArguments = @()
        $producerKind = 'explicit installed real Mermaid executable'
    }
}
$producerExecutable = [IO.Path]::GetFileNameWithoutExtension($MermaidExecutable)
if ([string]::IsNullOrWhiteSpace($producerExecutable)) { throw 'The configured Mermaid producer executable name is empty.' }
$producerDirectory = Split-Path -Parent $MermaidExecutable
$oldPath = $env:Path
$env:Path = "$producerDirectory;$env:Path"

Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($package)
try {
    $entries = @($archive.Entries | ForEach-Object FullName)
    $nuspecEntry = $archive.Entries | Where-Object { $_.FullName -match '\.nuspec$' } | Select-Object -First 1
    if ($null -eq $nuspecEntry) { throw 'Package nuspec is missing.' }
    $reader = [IO.StreamReader]::new($nuspecEntry.Open())
    try { [xml]$nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $metadata = $nuspec.package.metadata
    if ([string]$metadata.id -ne 'Yadg' -or [string]$metadata.version -ne $version) { throw "Package metadata must be Yadg $version." }
    if (-not ($metadata.packageTypes.packageType | Where-Object { $_.name -eq 'DotnetTool' })) { throw 'Package is not marked as a .NET tool.' }
    if ([string]$metadata.authors -ne 'carlrabbit' -or [string]$metadata.readme -ne 'README.md') { throw 'Required package author or README metadata is missing.' }
    if ([string]$metadata.license.type -ne 'expression' -or [string]$metadata.license.InnerText -ne 'MIT') { throw 'Package MIT license metadata is missing or incorrect.' }
    foreach ($required in @('README.md','LICENSE')) { if ($entries -notcontains $required) { throw "Package is missing $required." } }
    foreach ($required in @('yadg.dll','Yadg.Core.dll','Yadg.Word.dll','Yadg.Renderer.dll','Yadg.WordRenderer.dll')) {
        if (-not ($entries | Where-Object { $_ -like "tools/*/$required" })) { throw "Required package assembly is missing: $required" }
    }
    $frameworks = @($entries | ForEach-Object { if ($_ -match '^tools/([^/]+)/') { $Matches[1] } } | Sort-Object -Unique)
    if ($frameworks.Count -ne 1 -or $frameworks[0] -ne 'net10.0') { throw "Expected only net10.0 tool payload; found $($frameworks -join ', ')." }
    $cliAssemblyEntry = $archive.GetEntry(($entries | Where-Object { $_ -match '^tools/net10\.0/.*/yadg\.dll$' } | Select-Object -First 1))
    if ($null -eq $cliAssemblyEntry) { throw 'Packed x64 CLI assembly is missing.' }
    $assemblyStream = $cliAssemblyEntry.Open(); $assemblyMemory = [IO.MemoryStream]::new()
    try { $assemblyStream.CopyTo($assemblyMemory); $assemblyMemory.Position = 0; $peReader = [System.Reflection.PortableExecutable.PEReader]::new($assemblyMemory); try { if ($peReader.PEHeaders.CoffHeader.Machine -ne [System.Reflection.PortableExecutable.Machine]::Amd64) { throw 'Packed CLI assembly is not x64.' } } finally { $peReader.Dispose() } } finally { $assemblyStream.Dispose(); $assemblyMemory.Dispose() }
    $bad = @($entries | Where-Object { $_ -match '(?i)(^|/)(Interop\.Microsoft\.Office|Microsoft\.Office|office.*interop|interop.*office).*\.dll$|(^|/)(test|tests|fixtures|review|execution|\.git|obj|bin)(/|$)|(?i)(api[_-]?key|password|secret|credential)' })
    if ($bad.Count) { throw "Package content/hygiene check failed: $($bad -join ', ')" }
    $credentialPattern = '(?i)\b(api[_-]?key|password|secret|credential)\b["''\s]*[:=]\s*["'']?[A-Za-z0-9+/=_-]{12,}'
    foreach ($entry in $archive.Entries | Where-Object { $_.Length -le 2000000 -and $_.FullName -match '\.(json|config|xml|props|targets|md|nuspec|txt)$' }) {
        $entryReader = [IO.StreamReader]::new($entry.Open())
        try { $entryText = $entryReader.ReadToEnd() } finally { $entryReader.Dispose() }
        if ($entryText -match $credentialPattern) { throw "Package contains a credential-like value in $($entry.FullName)." }
    }
    $readmeEntry = $archive.GetEntry('README.md')
    $reader = [IO.StreamReader]::new($readmeEntry.Open())
    try { $packagedReadme = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($packagedReadme -notmatch 'YADG 1\.2\.0' -or $packagedReadme -notmatch 'MARKDOWN-AUTHORING' -or $packagedReadme -notmatch 'DOCX-TEMPLATE-AUTHORING') { throw 'Packaged README does not identify the candidate and canonical guides.' }
} finally { $archive.Dispose() }

$wordVersion = $word.VersionInfo.FileVersion
$loVersion = $lo.VersionInfo.FileVersion
$mermaidVersionOutput = 'package ' + $mermaidPackage + '@' + $mermaidVersion
if ($producerKind.StartsWith('Bun')) {
    $bunVersion = (& $MermaidExecutable --version 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) { throw "Mermaid package-manager runtime failed: $bunVersion" }
    $mermaidVersionOutput = "Bun $bunVersion; $mermaidPackage@$mermaidVersion"
}

$temporary = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0015-tier4-' + [guid]::NewGuid().ToString('N'))
$cliHome = Join-Path $temporary 'cli-home'; $nuget = Join-Path $temporary 'nuget'; $tool = Join-Path $temporary 'tool'; $consumer = Join-Path $temporary 'consumer'; $yolo = Join-Path $temporary 'yolo'; $weak = Join-Path $temporary 'weak'; $autoOff = Join-Path $temporary 'auto-off'; $delivery = Join-Path $temporary 'delivery'; $localFeed = Join-Path $temporary 'local-feed'
New-Item -ItemType Directory -Force -Path $cliHome,$nuget,$tool,$consumer,$yolo,$weak,$delivery,$localFeed | Out-Null
$oldHome = $env:DOTNET_CLI_HOME; $oldNuget = $env:NUGET_PACKAGES
$results = [ordered]@{}
try {
    $packageDirectory = Split-Path -Parent $package
    $env:DOTNET_CLI_HOME = $cliHome; $env:NUGET_PACKAGES = $nuget
    & dotnet tool install Yadg --tool-path $tool --version $version --add-source $packageDirectory --no-cache --ignore-failed-sources
    if ($LASTEXITCODE -ne 0) { throw 'Isolated installed-package setup failed.' }
    $yadg = Join-Path $tool 'yadg.exe'
    if (-not (Test-Path -LiteralPath $yadg)) { throw 'Installed yadg.exe command is missing.' }
    $resolved = (Get-Command $yadg).Source
    if ($resolved -match '[\\/]bin[\\/]|[\\/]obj[\\/]') { throw "Installed command resolves into repository build output: $resolved" }
    $installedVersion = (& $yadg --version 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -or $installedVersion -ne $version) { throw "Installed yadg --version mismatch: '$installedVersion'." }
    $results.install = [ordered]@{ result='passed'; command=$resolved; version=$installedVersion }

    $helpCases = @(
        @{args=@('--help'); required=@('init','check','inspect','build','render','publish')},
        @{args=@('init','--help'); required=@('--workspace')},
        @{args=@('check','--help'); required=@('--workspace','--list','--yolo')},
        @{args=@('inspect','--help'); required=@('styles','template')},
        @{args=@('inspect','styles','--help'); required=@('--workspace','--template')},
        @{args=@('inspect','template','--help'); required=@('--workspace','--template')},
        @{args=@('build','--help'); required=@('--workspace','--yolo')},
        @{args=@('render','--help'); required=@('--workspace','--renderer','libreoffice','word','--renderer-path','--yolo')},
        @{args=@('publish','--help'); required=@('--workspace','--publish-path')}
    )
    foreach ($case in $helpCases) {
        $output = (& $yadg @($case.args) 2>&1 | Out-String)
        if ($LASTEXITCODE -ne 0) { throw "Installed help failed: $($case.args -join ' ')." }
        foreach ($expected in $case.required) { if ($output -notmatch [regex]::Escape($expected)) { throw "Installed help '$($case.args -join ' ')' lacks '$expected'." } }
    }
    $renderHelp = (& $yadg render --help 2>&1 | Out-String)
    if ($renderHelp -notmatch 'default:\s*libreoffice') { throw 'Installed render help does not name libreoffice as the default.' }
    $results.help = 'root and command-specific help passed required command, option, and renderer-default assertions'

    $bootstrap = Join-Path $temporary 'bootstrap'; New-Item -ItemType Directory -Path $bootstrap | Out-Null
    Push-Location $bootstrap
    try { & $yadg init; if ($LASTEXITCODE -ne 0) { throw 'Installed init failed.' } } finally { Pop-Location }
    foreach ($owned in @('YADG.md','content.md','YadgTemplates')) { if (-not (Test-Path (Join-Path $bootstrap $owned))) { throw "Installed init omitted $owned." } }
    $before = (Get-FileHash (Join-Path $bootstrap 'content.md')).Hash
    Push-Location $bootstrap
    try { & $yadg init 2>$null; if ($LASTEXITCODE -eq 0) { throw 'Installed init overwrote existing owned paths.' } } finally { Pop-Location }
    if ((Get-FileHash (Join-Path $bootstrap 'content.md')).Hash -ne $before) { throw 'Installed init mutated an existing starter path.' }
    $results.bootstrap = 'passed; created starter workspace and preserved existing owned paths'

    $fixture = Join-Path $repo 'tests/fixtures/m0010/word-origin.docx'
    if (-not (Test-Path -LiteralPath $fixture)) { throw "Application-produced Word template fixture is missing: $fixture" }
    foreach ($workspace in @($consumer,$yolo,$weak,$autoOff)) { New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null; Copy-Item -LiteralPath $fixture -Destination (Join-Path $workspace 'YadgTemplates/template.docx') }
    $yaml = @('---','yadg:','  version: 1','values:','  document-version: "1.2"','  story-value: "M0015 installed consumer value"','markdown:','  thematicBreak: ignore','  codeInline: ignore','diagrams:','  autoDetect: true','publish:','  path: ./Published','producers:','  mermaid:','    executable: ' + $producerExecutable,'    arguments:') + @($mermaidArguments | ForEach-Object { '      - "' + $_.Replace('"','\"') + '"' }) + @('---')
    Set-Content -LiteralPath (Join-Path $consumer 'YADG.md') -Value $yaml -Encoding utf8
    $source = @'
# Architecture {#architecture}

An installed-tool release sample.

![](images/pixel.png){#sample-image}

```diagram {#overlap}
┌────────┐
│lower   │
│        │
│      ┌─┼───┐
└──────│     │
       │     │
       └─────┘
```

```diagram {#vertical-flow}
┌───┐
│ A │
└───┘
  ↓
  ↓
```

```
┌────────────┐
│ Auto       │
│ ┌──────┐   │
│ │Inner │   │
│ └──────┘   │
└────────────┘
```

```text {#detected-flow}
┌──────┐
│  A   │
└──┬───┘
   ▼
```

```mermaid {#mermaid-flow}
flowchart LR
    A[Source] --> B[DOCX]
```
'@
    Set-Content -LiteralPath (Join-Path $consumer 'content.md') -Value $source -Encoding utf8
    $imageDirectory = Join-Path $consumer 'images'; New-Item -ItemType Directory -Force -Path $imageDirectory | Out-Null
    $pixelPng = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aMn8AAAAASUVORK5CYII=')
    [IO.File]::WriteAllBytes((Join-Path $imageDirectory 'pixel.png'),$pixelPng)
    $check = (& $yadg check --workspace $consumer 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "Installed strict check failed for diagram examples: $check" }
    $list = (& $yadg check --workspace $consumer --list 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $list -notmatch 'detected-flow\s+figure') { throw "Installed check --list did not list identified text diagram as a figure: $list" }
    & $yadg inspect styles --workspace $consumer
    if ($LASTEXITCODE -ne 0) { throw 'Installed inspect styles failed.' }
    & $yadg inspect template --workspace $consumer
    if ($LASTEXITCODE -ne 0) { throw 'Installed inspect template failed.' }
    & $yadg build --workspace $consumer
    if ($LASTEXITCODE -ne 0) { throw 'Installed strict build failed.' }
    $authored = Get-ChildItem -LiteralPath (Join-Path $consumer 'YadgPreWords') -Filter '*.docx' -File | Select-Object -First 1
    if ($null -eq $authored) { throw 'Installed strict build produced no authored DOCX.' }
    $zip = [IO.Compression.ZipFile]::OpenRead($authored.FullName)
    try {
        $documentEntry = $zip.GetEntry('word/document.xml')
        if ($null -eq $documentEntry) { throw 'Authored DOCX has no word/document.xml.' }
        $reader = [IO.StreamReader]::new($documentEntry.Open())
        try { $documentXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $svgEntries = @($zip.Entries | Where-Object { $_.FullName -match '\.svg$' })
        $pngEntries = @($zip.Entries | Where-Object { $_.FullName -match '\.png$' })
        if ($svgEntries.Count -lt 4 -or ([regex]::Matches($documentXml,'svgBlip')).Count -lt 4) { throw "Expected four installed text-diagram SVG parts and figure references; found $($svgEntries.Count)." }
        if ($pngEntries.Count -lt 2) { throw 'Expected both a source PNG figure and a real Mermaid-produced PNG in the authored DOCX.' }
        $pngHashes = foreach ($entry in $pngEntries) { $stream = $entry.Open(); try { $sha = [Security.Cryptography.SHA256]::Create(); try { [Convert]::ToHexString($sha.ComputeHash($stream)) } finally { $sha.Dispose() } } finally { $stream.Dispose() } }
        $sourcePngHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($pixelPng))
        if (-not ($pngHashes | Where-Object { $_ -ne $sourcePngHash })) { throw 'A PNG distinct from the source-file figure was not embedded; Mermaid output was not established.' }
        $svgText = foreach ($entry in $svgEntries) { $reader = [IO.StreamReader]::new($entry.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }
        $svgAll = $svgText -join "`n"
        foreach ($term in @('lower','Auto','Inner','A')) { if ($svgAll -notmatch [regex]::Escape($term)) { throw "Generated SVG content omitted expected grid text '$term'." } }
        if ($svgAll -notmatch '<path' -or $svgAll -match '<script|foreignObject|(?:href|src)\s*=\s*["''](?:https?:|//|data:)') { throw 'Generated text-diagram SVG lacks arrow geometry or includes active/external content.' }
    } finally { $zip.Dispose() }
    $results.strictAuthoring = 'installed strict check/build, check --list, inspect styles/template passed'
    $results.textDiagrams = [ordered]@{ result='passed'; explicitIdentified='overlap corner and standalone vertical arrow'; autoDetect='anonymous unlabeled and identified lowercase text fence'; svgParts=$svgEntries.Count; svgReferences=([regex]::Matches($documentXml,'svgBlip')).Count }
    $results.mermaid = [ordered]@{ result='passed'; producer=$producerKind; executable=$MermaidExecutable; arguments=$mermaidArguments; pinnedPackage="$mermaidPackage@$mermaidVersion"; embeddedPngParts=$pngEntries.Count }

    @('---','yadg:','  version: 1','values:','  story-value: "M0015 installed consumer value"','diagrams:','  autoDetect: true','---') | Set-Content -LiteralPath (Join-Path $weak 'YADG.md') -Encoding utf8
    @('# Weak candidate {#weak}','','```text','ordinary label: value','```') | Set-Content -LiteralPath (Join-Path $weak 'content.md') -Encoding utf8
    $weakResult = (& $yadg check --workspace $weak 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0 -or $weakResult -notmatch 'YADG-MD-UNSUPPORTED') { throw "Auto-detection converted a weak text fence or returned the wrong diagnostic: $weakResult" }
    $results.autoDetectionNegative = [ordered]@{ result='passed'; expected='YADG-MD-UNSUPPORTED'; observed='weak ordinary text fence remains unsupported' }
    @('---','yadg:','  version: 1','values:','  story-value: "M0015 installed consumer value"','---') | Set-Content -LiteralPath (Join-Path $autoOff 'YADG.md') -Encoding utf8
    @('# Auto-detection disabled {#disabled}','','```text','┌────────────┐','│ Auto       │','│ ┌──────┐   │','│ │Inner │   │','│ └──────┘   │','└────────────┘','```') | Set-Content -LiteralPath (Join-Path $autoOff 'content.md') -Encoding utf8
    $autoOffResult = (& $yadg check --workspace $autoOff 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0 -or $autoOffResult -notmatch 'YADG-MD-UNSUPPORTED') { throw 'Strong text-fence diagram was accepted while diagrams.autoDetect was absent (default off).' }
    $results.autoDetectionOptIn = 'passed; default-off strong text fence remained unsupported'

    @('---','yadg:','  version: 1','values:','  story-value: "M0015 installed consumer value"','markdown:','  thematicBreak: ignore','  codeInline: ignore','---') | Set-Content -LiteralPath (Join-Path $yolo 'YADG.md') -Encoding utf8
    @('# Architecture {#architecture}','','1. First item','2. Second item','','| Component | Purpose |','|---|---|','| YADG | Authors DOCX |','{#yolo-table}','','```diagram {#captioned caption="Release sample"}','┌────┐','│ Y  │','└────┘','```') | Set-Content -LiteralPath (Join-Path $yolo 'content.md') -Encoding utf8
    $strictYoloCheck = (& $yadg check --workspace $yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0) { throw 'Strict packaged authoring unexpectedly accepted the unavailable fixture list presentation.' }
    $yoloCheck = (& $yadg check --workspace $yolo --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $yoloCheck -notmatch 'builtin:ordered-list-v1') { throw "Installed YOLO check did not use expected ordered-list recovery: $yoloCheck" }
    $strictBuild = (& $yadg build --workspace $yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0) { throw 'Strict packaged build unexpectedly accepted the unavailable fixture list presentation.' }
    $yoloBuild = (& $yadg build --workspace $yolo --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $yoloBuild -notmatch 'degradation' -or $yoloBuild -notmatch 'builtin:ordered-list-v1') { throw "Installed YOLO build did not produce truthful degraded output: $yoloBuild" }
    $results.yoloAuthoring = 'strict check/build rejected unavailable ordered-list presentation; installed YOLO check/build used builtin:ordered-list-v1 with visible degradation'

    & $yadg render --workspace $consumer --renderer word
    if ($LASTEXITCODE -ne 0) { throw 'Installed Microsoft Word renderer failed.' }
    $results.wordRender = [ordered]@{ result='passed'; version=$wordVersion }
    & $yadg render --workspace $consumer --renderer libreoffice --renderer-path $loPath
    if ($LASTEXITCODE -ne 0) { throw 'Installed LibreOffice renderer failed.' }
    $results.libreOfficeRender = [ordered]@{ result='passed'; version=$loVersion }
    $fallback = (& $yadg render --workspace $consumer --renderer libreoffice --renderer-path (Join-Path $temporary 'missing-soffice.exe') --yolo 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $fallback -notmatch 'requested=libreoffice' -or $fallback -notmatch 'actual=word') { throw "Installed renderer YOLO fallback failed or hid actual renderer: $fallback" }
    $results.rendererFallback = 'passed; requested libreoffice and actual word were reported'

    $links = @('README.md','docs/user/MARKDOWN-AUTHORING.md','docs/user/DOCX-TEMPLATE-AUTHORING.md','docs/user/AUTHORING.md','docs/user/TEMPLATES.md','docs/user/GETTING-STARTED.md','docs/user/TROUBLESHOOTING.md')
    $linkFailures = @()
    foreach ($relative in $links) {
        $file = Join-Path $repo $relative
        $markdownText = (Get-Content -LiteralPath $file -Raw) -replace '(?s)```.*?```',''
        foreach ($match in [regex]::Matches($markdownText,'(?<!!)\[[^\]]+\]\(([^)]+)\)')) {
            $target = [string]$match.Groups[1].Value
            if ($target -match '^(?:https?:|mailto:|#)') { continue }
            $pathPart = ($target -split '#',2)[0]
            if (-not [string]::IsNullOrWhiteSpace($pathPart) -and -not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $file) $pathPart))) { $linkFailures += "$relative -> $target" }
        }
    }
    if ($linkFailures.Count) { throw "Documentation links do not resolve: $($linkFailures -join '; ')" }
    foreach ($relative in @('docs/user/AUTHORING.md','docs/user/TEMPLATES.md')) {
        $content = Get-Content -LiteralPath (Join-Path $repo $relative) -Raw
        if ($content -notmatch 'MARKDOWN-AUTHORING|DOCX-TEMPLATE-AUTHORING') { throw "Compatibility landing page is not useful: $relative" }
    }
    foreach ($required in @('YADG 1\.2\.0','check --list','yadg build','yadg render','yadg publish','diagram','Mermaid','DOCX-TEMPLATE-AUTHORING')) { if ((Get-Content -LiteralPath (Join-Path $repo 'README.md') -Raw) -notmatch $required) { throw "README is missing required release surface '$required'." } }
    $markdownGuide = Get-Content -LiteralPath (Join-Path $repo 'docs/user/MARKDOWN-AUTHORING.md') -Raw
    foreach ($required in @('stable IDs','Sections and content','Paragraphs and inline content','Flat lists','Pipe tables','PNG','caption','thematicBreak','codeInline','```diagram','character grid','Standalone arrows','Nested boxes','corner','auto-detection','anonymous','```mermaid','Text diagram or Mermaid?','YOLO')) { if ($markdownGuide -notmatch [regex]::Escape($required)) { throw "Markdown authoring guide is missing required source-authoring topic '$required'." } }
    $templateGuide = Get-Content -LiteralPath (Join-Path $repo 'docs/user/DOCX-TEMPLATE-AUTHORING.md') -Raw
    foreach ($required in @('ordinary visible text tags','{{section:architecture}}','{{content:architecture}}','{{table:interfaces}}','{{figure:system-context}}','{{value:document-version}}','front matter','styles:','prototypes:','inspect styles','inspect template','Headings','List-item prototypes','exactly one `SEQ` field','bookmark/reference mechanics','{{table-rows:interfaces}}','{{cell}}','codeInline','TOC','YOLO','near:','both reach template authoring as figure-like content')) { if ($templateGuide -notmatch [regex]::Escape($required)) { throw "DOCX template guide is missing required template-authoring topic '$required'." } }
    $results.documentation = 'passed; links/compatibility pages and required README/template-guide surface checked; installed source examples parsed and built'

    Remove-Item -LiteralPath (Join-Path $consumer 'YADG.md'),(Join-Path $consumer 'content.md'),(Join-Path $consumer 'YadgTemplates') -Force -Recurse
    $published = Join-Path $delivery 'consumer'
    & $yadg publish --workspace $consumer --publish-path $published
    if ($LASTEXITCODE -ne 0 -or -not (Get-ChildItem -LiteralPath $published -Filter '*.docx' -File)) { throw 'Installed downstream publish failed.' }
    $results.publish = 'passed; installed tool published finalized DOCX to local filesystem destination'
    Copy-Item -LiteralPath $package -Destination $localFeed
    & (Join-Path $repo 'eng/publish-nuget.ps1') -Source $localFeed -PackagePath $package
    if ($LASTEXITCODE -ne 0) { throw 'Local-only publish-nuget validation failed.' }
    $results.publishHelperLocal = 'passed; local filesystem NuGet source only'

    $statusHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($tree)))
    New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
    $evidence = [ordered]@{
        milestone='M0015'; repositoryRevision=$revision; workingTreeStatus=$tree; workingTreeStatusSha256=$statusHash
        releaseVersion=$version; packageId='Yadg'; packagePath=$package; packageSha256=(Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash
        packageInspection=[ordered]@{result='passed'; targetFramework=$frameworks[0]; contentSummary=$entries; officeInteropAssemblies='none'; forbiddenContent='none'}
        dotnetSdk=(& dotnet --version).Trim(); windowsVersion=(Get-CimInstance Win32_OperatingSystem).Caption
        wordVersion=$wordVersion; libreOfficeVersion=$loVersion
        mermaidProducer=[ordered]@{kind=$producerKind; executable=$MermaidExecutable; version=$mermaidVersionOutput; invocationArguments=$mermaidArguments; package="$mermaidPackage@$mermaidVersion"}
        m0014Tier3=[ordered]@{evidencePath='artifacts/review/evidence/M0014/tier3-evidence.json'; repositoryRevision=[string]$tier3.repositoryRevision}
        installedConsumer=$results; documentationExampleAudit='passed; installed candidate parsed/built explicit occluded and standalone-arrow diagrams, auto-detection positive and negative cases, and real Mermaid; links and guide contracts passed'
        externalPublication='none'
    }
    $evidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'release-evidence.json') -Encoding utf8
    Write-Output 'M0015 Tier 4 installed-package consumer validation passed.'
}
finally {
    $env:DOTNET_CLI_HOME = $oldHome; $env:NUGET_PACKAGES = $oldNuget; $env:Path = $oldPath
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
}
