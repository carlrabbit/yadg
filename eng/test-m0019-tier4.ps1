param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0019'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
$sdkVersion = (& dotnet --list-sdks | ForEach-Object { if ($_ -match '^\s*(10\.0\.\d+)\s') { $Matches[1] } } | Sort-Object { [version]$_ } -Descending | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace([string]$sdkVersion)) { throw 'M0019 Tier 4 requires an installed .NET 10 SDK.' }
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0019-tier4-' + [guid]::NewGuid().ToString('N'))
$packageDir = Join-Path $scratch 'package'
$toolDir = Join-Path $scratch 'tool'
$workspace = Join-Path $scratch 'workspace'
New-Item -ItemType Directory -Force -Path $packageDir,$toolDir,$workspace,$EvidenceRoot,(Join-Path $workspace 'YadgTemplates'),(Join-Path $workspace 'images') | Out-Null
$globalPath = Join-Path $scratch 'global.json'
Set-Content -LiteralPath $globalPath -Value ('{"sdk":{"version":"' + $sdkVersion + '","rollForward":"disable","allowPrerelease":false}}') -Encoding ascii

function Add-ZipText($zip, [string] $name, [string] $content) {
    $entry = $zip.CreateEntry($name)
    $stream = $entry.Open()
    try { $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false)); try { $writer.Write($content) } finally { $writer.Dispose() } } finally { $stream.Dispose() }
}

try {
    $template = Join-Path $workspace 'template.docx'
    $archive = [IO.Compression.ZipFile]::Open($template,[IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipText $archive '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        Add-ZipText $archive '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        Add-ZipText $archive 'word/document.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{{content:report}}</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
    } finally { $archive.Dispose() }

    $source = Join-Path $workspace 'consumer.md'
    Set-Content -LiteralPath $source -Encoding utf8 -Value @('# Consumer report {#report}','','Hello {{value:name}}.','')
    $config = Join-Path $workspace 'authoring.yaml'
    Set-Content -LiteralPath $config -Encoding utf8 -Value @('yadg:','  version: 1','values:','  name: "Installed tool"','markdown:','  codeInline: ignore')

    $presentationTemplate = Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx'
    $presentationSource = Join-Path $repo 'tests/fixtures/m0018/review-source.md'
    Copy-Item -LiteralPath $presentationTemplate -Destination (Join-Path $workspace 'YadgTemplates/impress.pptx')
    Copy-Item -LiteralPath $presentationSource -Destination (Join-Path $workspace 'slides.md')
    Copy-Item -LiteralPath (Join-Path $repo 'tests/fixtures/m0018/images/native-office-sample.png') -Destination (Join-Path $workspace 'images/native-office-sample.png')
    Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Encoding utf8 -Value @('---','yadg:','  version: 1','presentations:','  - id: overview','    template: YadgTemplates/impress.pptx','    sources:','      - slides.md','    output: overview.pptx','---')

    Push-Location -LiteralPath $scratch
    try {
        & dotnet pack (Join-Path $repo 'src/Yadg.Cli/Yadg.Cli.csproj') --configuration Release --output $packageDir
        if ($LASTEXITCODE -ne 0) { throw 'M0019 current Yadg tool package build failed.' }
        $package = Join-Path $packageDir "Yadg.$version.nupkg"
        & dotnet tool install --tool-path $toolDir --add-source $packageDir --ignore-failed-sources Yadg --version $version
        if ($LASTEXITCODE -ne 0) { throw 'M0019 isolated current-tool install failed.' }
        $shim = Join-Path $toolDir 'yadg.exe'
        & $shim document check --template $template --source $source --config $config
        if ($LASTEXITCODE -ne 0) { throw 'Installed tool direct DOCX check failed.' }
        $docxOutput = Join-Path $workspace 'authored.docx'
        & $shim document build --template $template --source $source --output $docxOutput --config $config
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxOutput)) { throw 'Installed tool direct DOCX build failed.' }
        $yoloSource = Join-Path $workspace 'consumer-yolo.md'
        Set-Content -LiteralPath $yoloSource -Encoding utf8 -Value @('# YOLO report {#report}','','Unresolved token: {{value:missing}}.')
        $yoloDocx = Join-Path $workspace 'authored-yolo.docx'
        $yoloDocxSummary = (& $shim document build --template $template --source $yoloSource --output $yoloDocx --config $config --yolo 2>&1) -join "`n"
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $yoloDocx) -or $yoloDocxSummary -notmatch 'degradations=1') { throw 'Installed tool direct DOCX YOLO value recovery failed or reported an inaccurate count.' }
        & $shim presentation check --template (Join-Path $workspace 'YadgTemplates/impress.pptx') --source (Join-Path $workspace 'slides.md') --config $config
        if ($LASTEXITCODE -ne 0) { throw 'Installed tool direct PPTX check with explicit config failed.' }
        $pptxOutput = Join-Path $workspace 'direct.pptx'
        & $shim presentation build --template (Join-Path $workspace 'YadgTemplates/impress.pptx') --source (Join-Path $workspace 'slides.md') --output $pptxOutput --config $config
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $pptxOutput)) { throw 'Installed tool direct PPTX build failed.' }
        $yoloPresentationSource = Join-Path $workspace 'slides-yolo.md'
        $reviewMarkdown = Get-Content -LiteralPath (Join-Path $workspace 'slides.md') -Raw
        $reviewMarkdown = [regex]::Replace($reviewMarkdown,'!\[[^\]]*\]\([^\)]+\)','![missing](images/not-present.png)')
        Set-Content -LiteralPath $yoloPresentationSource -Encoding utf8 -Value $reviewMarkdown
        $yoloPptx = Join-Path $workspace 'direct-yolo.pptx'
        $yoloPptxSummary = (& $shim presentation build --template (Join-Path $workspace 'YadgTemplates/impress.pptx') --source $yoloPresentationSource --output $yoloPptx --config $config --yolo 2>&1) -join "`n"
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $yoloPptx) -or $yoloPptxSummary -notmatch 'degradations=1') { throw 'Installed tool direct PPTX YOLO image recovery failed or reported an inaccurate count.' }
        & $shim presentation check --id overview --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw 'Installed tool selected-deck check failed.' }
        & $shim presentation build --id overview --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw 'Installed tool selected-deck build failed.' }
        $inventory = (& $shim check --list --workspace $workspace 2>&1) -join "`n"
        if ($inventory -notmatch 'overview' -or $inventory -notmatch 'slides.md' -or $inventory -notmatch 'overview.pptx') { throw "Installed tool presentation inventory did not list the configured deck declaration. Output: $inventory" }

        $docHash = (Get-FileHash -LiteralPath $docxOutput -Algorithm SHA256).Hash.ToUpperInvariant()
        $pptHash = (Get-FileHash -LiteralPath $pptxOutput -Algorithm SHA256).Hash.ToUpperInvariant()
        Copy-Item -LiteralPath $package -Destination (Join-Path $EvidenceRoot "Yadg.$version.nupkg") -Force
        $record = [ordered]@{ milestone='M0019'; gate='VAL-06'; result='passed'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim(); packageId='Yadg'; packageVersion=$version; sdkVersion=$sdkVersion; packagePath="artifacts/review/evidence/M0019/Yadg.$version.nupkg"; packageSha256=(Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToUpperInvariant(); isolatedToolPath='temporary dotnet tool path'; commands=@('document check/build --config','document build --config --yolo','presentation check/build --config','presentation build --config --yolo','presentation check/build --id','presentation build --id','check --list'); authoredDocxSha256=$docHash; yoloDocxSha256=(Get-FileHash -LiteralPath $yoloDocx -Algorithm SHA256).Hash.ToUpperInvariant(); directPptxSha256=$pptHash; yoloPptxSha256=(Get-FileHash -LiteralPath $yoloPptx -Algorithm SHA256).Hash.ToUpperInvariant(); selectedPptxSha256=(Get-FileHash -LiteralPath (Join-Path $workspace 'YadgPresentations/overview.pptx') -Algorithm SHA256).Hash.ToUpperInvariant() }
        $record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier4-evidence.json') -Encoding utf8
        Write-Output "M0019 isolated Yadg $version Tier 4 commands passed."
    } finally { Pop-Location }
} finally {
    if (Test-Path -LiteralPath $scratch) {
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
        $resolved = [IO.Path]::GetFullPath($scratch)
        if (-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notmatch '^yadg-m0019-tier4-[0-9a-f]{32}$') { throw 'Refusing to delete a Tier 4 scratch directory outside the expected temporary path.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}
