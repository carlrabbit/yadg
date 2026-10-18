param(
    [string] $PackagePath,
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/release/evidence/M0015'),
    [string] $MermaidExecutable
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

function Get-M0015DocumentedExample([string]$Path, [string]$Id) {
    $lines = [IO.File]::ReadAllLines($Path)
    $marker = "<!-- M0015-example: $Id -->"
    $matches = @(); for ($i = 0; $i -lt $lines.Length; $i++) { if ($lines[$i].Trim() -ceq $marker) { $matches += $i } }
    if ($matches.Count -ne 1) { throw "Expected exactly one documented example marker '$Id' in $Path; found $($matches.Count)." }
    $start = $matches[0] + 1
    if ($start -ge $lines.Length -or $lines[$start] -notmatch '^(?<fence>`{3,})(?<language>[A-Za-z0-9_-]+)\s*$') { throw "Documented example '$Id' must be followed immediately by a fenced code block with a language." }
    $fence = $Matches.fence; $language = $Matches.language
    $end = -1
    for ($i = $start + 1; $i -lt $lines.Length; $i++) { if ($lines[$i] -ceq $fence) { $end = $i; break } }
    if ($end -lt 0) { throw "Documented example '$Id' has no closing $fence fence." }
    return [pscustomobject]@{ Id=$Id; Path=$Path; Language=$language; Body=($lines[($start + 1)..($end - 1)] -join "`n") }
}

function Get-M0015FrontMatterLines($Example) {
    if ($Example.Language -ne 'yaml') { throw "Example '$($Example.Id)' must be a yaml fence." }
    $lines = @($Example.Body -split '\r?\n')
    if ($lines.Count -lt 3 -or $lines[0] -cne '---' -or $lines[-1] -cne '---') { throw "Example '$($Example.Id)' must contain a complete front-matter block." }
    return ,$lines
}

function New-M0015Paragraph([Xml.XmlDocument]$Document, [string]$Text) {
    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $paragraph = $Document.CreateElement('w','p',$wordNs)
    $run = $Document.CreateElement('w','r',$wordNs)
    $textNode = $Document.CreateElement('w','t',$wordNs); $textNode.InnerText = $Text
    if ($Text.StartsWith(' ') -or $Text.EndsWith(' ')) { [void]$textNode.SetAttribute('xml:space','preserve') }
    [void]$run.AppendChild($textNode); [void]$paragraph.AppendChild($run)
    return $paragraph
}

function Add-M0015RunText([Xml.XmlDocument]$Document, [Xml.XmlElement]$Paragraph, [string]$Text) {
    if ($Text.Length -eq 0) { return }
    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $run = $Document.CreateElement('w','r',$wordNs)
    $textNode = $Document.CreateElement('w','t',$wordNs); $textNode.InnerText = $Text
    if ($Text.StartsWith(' ') -or $Text.EndsWith(' ')) { [void]$textNode.SetAttribute('xml:space','preserve') }
    [void]$run.AppendChild($textNode); [void]$Paragraph.AppendChild($run)
}

function New-M0015CaptionParagraph([Xml.XmlDocument]$Document, [string]$CaptionText) {
    $match = [regex]::Match($CaptionText,'^(?<prefix>.*?)\[SEQ (?<sequence>Figure|Table) field\](?<suffix>.*)$')
    if (-not $match.Success -or ([regex]::Matches($CaptionText,'\{\{caption\}\}')).Count -ne 1) { throw "Documented caption-prototype example is malformed: '$CaptionText'." }
    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $paragraph = $Document.CreateElement('w','p',$wordNs)
    Add-M0015RunText $Document $paragraph $match.Groups['prefix'].Value
    $field = $Document.CreateElement('w','fldSimple',$wordNs)
    [void]$field.SetAttribute('instr',' SEQ ' + $match.Groups['sequence'].Value + ' \* ARABIC ')
    $run = $Document.CreateElement('w','r',$wordNs); $textNode = $Document.CreateElement('w','t',$wordNs); $textNode.InnerText = '1'
    [void]$run.AppendChild($textNode); [void]$field.AppendChild($run); [void]$paragraph.AppendChild($field)
    Add-M0015RunText $Document $paragraph $match.Groups['suffix'].Value
    return $paragraph
}

function New-M0015TemplateFromExamples([string]$Fixture, [string]$Destination, [string[]]$FrontMatter, [string[]]$DocumentedTags, [object[]]$PrototypeExamples) {
    $zipIn = [IO.Compression.ZipFile]::OpenRead($Fixture)
    $zipOut = [IO.Compression.ZipFile]::Open($Destination,[IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $zipIn.Entries) {
            $newEntry = $zipOut.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::Optimal)
            $outStream = $newEntry.Open()
            try {
                $inStream = $entry.Open()
                try {
                    if ($entry.FullName -ceq 'word/document.xml') {
                        $sourceReader = [IO.StreamReader]::new($inStream)
                        try { $xml = [Xml.XmlDocument]::new(); $xml.PreserveWhitespace = $true; $xml.LoadXml($sourceReader.ReadToEnd()) } finally { $sourceReader.Dispose() }
                        $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w',$wordNs)
                        $body = $xml.SelectSingleNode('/w:document/w:body',$ns)
                        $paragraphs = @($body.SelectNodes('./w:p',$ns))
                        $startIndex = -1; $endIndex = -1
                        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
                            $paragraphText = [string]::Join('',@($paragraphs[$i].SelectNodes('./w:r/w:t',$ns) | ForEach-Object InnerText))
                            if ($paragraphText -ceq '{{yadg:frontmatter}}') { $startIndex = $i }
                            if ($paragraphText -ceq '{{/yadg:frontmatter}}') { $endIndex = $i; break }
                        }
                        if ($startIndex -lt 0 -or $endIndex -le $startIndex) { throw 'Base template has no complete front-matter marker region.' }
                        for ($i = $endIndex - 1; $i -gt $startIndex; $i--) { [void]$body.RemoveChild($paragraphs[$i]) }
                        $closing = $paragraphs[$endIndex]; $anchor = $closing
                        foreach ($line in $FrontMatter[1..($FrontMatter.Count - 2)]) {
                            $paragraph = New-M0015Paragraph $xml $line
                            [void]$body.InsertBefore($paragraph,$closing); $anchor = $paragraph
                        }
                        $anchor = $closing
                        $insertPrototypes = @()
                        foreach ($prototype in $PrototypeExamples) {
                            if ($prototype.Id -match 'list-prototype$') {
                                $prototypeLines = @($prototype.Body -split '\r?\n' | Where-Object { $_ -ne '' })
                                $prototypeId = [regex]::Match($prototypeLines[0],'\{\{yadg:prototype:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}').Groups['id'].Value
                                $startText = '{{yadg:prototype:' + $prototypeId + '}}'
                                $endText = '{{/yadg:prototype:' + $prototypeId + '}}'
                                $bodyParagraphs = @($body.SelectNodes('./w:p',$ns)); $found = $false
                                for ($p = 0; $p -lt $bodyParagraphs.Count - 2; $p++) {
                                    $texts = @($bodyParagraphs[$p..($p + 2)] | ForEach-Object { [string]::Join('',@($_.SelectNodes('./w:r/w:t',$ns) | ForEach-Object InnerText)) })
                                    if ($texts[0] -ceq $startText -and $texts[1] -ceq '{{item}}' -and $texts[2] -ceq $endText) {
                                        if ($null -eq $bodyParagraphs[$p + 1].SelectSingleNode('./w:pPr/w:numPr',$ns)) { throw "Existing documented list prototype '$prototypeId' has no concrete Word numbering." }
                                        $found = $true; break
                                    }
                                }
                                if (-not $found) { throw "Application-produced fixture does not contain documented list prototype '$prototypeId' with its item placeholder and Word numbering." }
                            } else { $insertPrototypes += $prototype }
                        }
                        foreach ($prototype in $insertPrototypes) {
                            $bodyLines = @($prototype.Body -split '\r?\n' | Where-Object { $_ -ne '' })
                            if ($bodyLines.Count -ne 3) { throw "Documented template prototype '$($prototype.Id)' must contain start, body, and end lines." }
                            $startMarker = $bodyLines[0].Trim(); $endMarker = $bodyLines[2].Trim()
                            if ($startMarker -notmatch '^\{\{yadg:prototype:(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$') { throw "Documented template prototype '$($prototype.Id)' has an invalid start marker." }
                            $prototypeId = $Matches.id
                            if ($endMarker -cne ('{{/yadg:prototype:' + $prototypeId + '}}')) { throw "Documented template prototype '$($prototype.Id)' has mismatched marker lines." }
                            $paragraph = New-M0015Paragraph $xml $startMarker; [void]$body.InsertAfter($paragraph,$anchor); $anchor = $paragraph
                            if ($prototype.Id -match 'caption-prototype$') { $paragraph = New-M0015CaptionParagraph $xml $bodyLines[1] }
                            else {
                                if ($bodyLines[1].Trim() -cne '{{item}}') { throw "Documented list prototype '$($prototype.Id)' must contain exactly {{item}}." }
                                $paragraph = New-M0015Paragraph $xml '{{item}}'
                            }
                            [void]$body.InsertAfter($paragraph,$anchor); $anchor = $paragraph
                            $paragraph = New-M0015Paragraph $xml $endMarker; [void]$body.InsertAfter($paragraph,$anchor); $anchor = $paragraph
                        }
                        $tagMap = @{}
                        foreach ($tagLine in $DocumentedTags) {
                            $tag = $tagLine.Trim()
                            if ($tag -notmatch '^\{\{(?<kind>section|content|table|figure|value):(?<id>[A-Za-z][A-Za-z0-9_-]*)\}\}$') { throw "Unrecognized documented template tag '$tag'." }
                            $kind = $Matches.kind
                            if ($tagMap.ContainsKey($kind)) { throw "Documented tag example repeats kind '$kind'." }
                            $tagMap[$kind] = $tag
                        }
                        foreach ($requiredKind in @('section','content','table','figure','value')) { if (-not $tagMap.ContainsKey($requiredKind)) { throw "Documented template tag example is missing '$requiredKind'." } }
                        $replacement = @{'{{section:architecture}}'=$tagMap.section; '{{content:architecture}}'=$tagMap.content; '{{content:introduction}}'=$tagMap.content; '{{value:story-value}}'=$tagMap.value}
                        foreach ($paragraph in @($body.SelectNodes('./w:p',$ns))) {
                            $paragraphText = [string]::Join('',@($paragraph.SelectNodes('./w:r/w:t',$ns) | ForEach-Object InnerText))
                            if ($replacement.ContainsKey($paragraphText)) {
                                $textNodes = @($paragraph.SelectNodes('./w:r/w:t',$ns))
                                if ($textNodes.Count -ne 1) { throw "Expected one text run in template marker paragraph '$paragraphText'." }
                                $textNodes[0].InnerText = $replacement[$paragraphText]
                            }
                        }
                        foreach ($kind in @('section','table','figure','value')) { $paragraph = New-M0015Paragraph $xml $tagMap[$kind]; [void]$body.InsertAfter($paragraph,$anchor); $anchor=$paragraph }
                        $memory = [IO.MemoryStream]::new()
                        try { $xml.Save($memory); $memory.Position=0; $memory.CopyTo($outStream) } finally { $memory.Dispose() }
                    } else { $inStream.CopyTo($outStream) }
                } finally { $inStream.Dispose() }
            } finally { $outStream.Dispose() }
        }
    } finally { $zipOut.Dispose(); $zipIn.Dispose() }
}

$props = [xml](Get-Content -LiteralPath (Join-Path $repo 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.VersionPrefix
if ($version -ne '1.2.0') { throw "M0015 Tier 4 requires central release version 1.2.0; found '$version'." }
if ([string]::IsNullOrWhiteSpace($PackagePath)) { $PackagePath = Join-Path $repo "artifacts/package/Yadg.$version.nupkg" }
$package = [IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf) -or (Split-Path -Leaf $package) -ne "Yadg.$version.nupkg") { throw "Expected freshly packed Yadg.$version.nupkg: $package" }

$markdownGuidePath = Join-Path $repo 'docs/user/MARKDOWN-AUTHORING.md'
$templateGuidePath = Join-Path $repo 'docs/user/DOCX-TEMPLATE-AUTHORING.md'
$documentedExampleSpecs = @(
    @{id='workspace-schema-v1'; path=$markdownGuidePath; language='yaml'},
    @{id='markdown-unordered-list'; path=$markdownGuidePath; language='markdown'},
    @{id='markdown-ordered-list'; path=$markdownGuidePath; language='markdown'},
    @{id='markdown-table'; path=$markdownGuidePath; language='markdown'},
    @{id='markdown-figure'; path=$markdownGuidePath; language='markdown'},
    @{id='explicit-text-diagram'; path=$markdownGuidePath; language='markdown'},
    @{id='standalone-arrow-diagram'; path=$markdownGuidePath; language='markdown'},
    @{id='nested-box-diagram'; path=$markdownGuidePath; language='markdown'},
    @{id='corner-occlusion-diagram'; path=$markdownGuidePath; language='markdown'},
    @{id='autodetect-schema-v1'; path=$markdownGuidePath; language='yaml'},
    @{id='autodetect-weak-negative'; path=$markdownGuidePath; language='markdown'},
    @{id='autodetect-identified-positive'; path=$markdownGuidePath; language='markdown'},
    @{id='mermaid-producer-config'; path=$markdownGuidePath; language='yaml'},
    @{id='mermaid-source-fence'; path=$markdownGuidePath; language='markdown'},
    @{id='template-visible-tags'; path=$templateGuidePath; language='text'},
    @{id='template-frontmatter-schema'; path=$templateGuidePath; language='yaml'},
    @{id='template-list-prototype-binding'; path=$templateGuidePath; language='yaml'},
    @{id='template-list-prototype'; path=$templateGuidePath; language='text'},
    @{id='template-ordered-list-prototype'; path=$templateGuidePath; language='text'},
    @{id='template-caption-prototype'; path=$templateGuidePath; language='text'},
    @{id='template-table-caption-prototype'; path=$templateGuidePath; language='text'}
)
$documentedExamples = @{}
foreach ($spec in $documentedExampleSpecs) {
    $example = Get-M0015DocumentedExample $spec.path $spec.id
    if ($example.Language -cne $spec.language -or [string]::IsNullOrWhiteSpace($example.Body)) { throw "Documented example '$($spec.id)' must be a non-empty $($spec.language) fence." }
    $documentedExamples[$spec.id] = $example
}

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
    if ($packagedReadme -notmatch 'YADG 1\.2\.0') { throw 'Packaged README does not identify the candidate.' }
    foreach ($guidePath in @('docs/user/MARKDOWN-AUTHORING.md','docs/user/DOCX-TEMPLATE-AUTHORING.md')) {
        $stableUrl = 'https://github.com/carlrabbit/yadg/blob/main/' + $guidePath
        if ($packagedReadme -notmatch [regex]::Escape($stableUrl)) { throw "Packaged README lacks usable stable guide link '$stableUrl'." }
    }
    if ($packagedReadme -match '\]\(docs/user/') { throw 'Packaged README contains repository-relative guide links that cannot resolve from the tool package.' }
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
$cliHome = Join-Path $temporary 'cli-home'; $nuget = Join-Path $temporary 'nuget'; $tool = Join-Path $temporary 'tool'; $consumer = Join-Path $temporary 'consumer'; $templateExample = Join-Path $temporary 'template-example'; $yolo = Join-Path $temporary 'yolo'; $weak = Join-Path $temporary 'weak'; $autoOff = Join-Path $temporary 'auto-off'; $delivery = Join-Path $temporary 'delivery'; $localFeed = Join-Path $temporary 'local-feed'
New-Item -ItemType Directory -Force -Path $cliHome,$nuget,$tool,$consumer,$yolo,$weak,$delivery,$localFeed | Out-Null
$oldHome = $env:DOTNET_CLI_HOME; $oldNuget = $env:NUGET_PACKAGES
$results = [ordered]@{}
$results.packagedReadme = 'passed; exact packaged README contains absolute GitHub links to both canonical authoring guides and no repository-relative docs/user links'
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
$templateFixture = Join-Path $repo 'tests/fixtures/m0012/list-prototypes.docx'
if (-not (Test-Path -LiteralPath $templateFixture)) { throw "Application-produced Word list-prototype fixture is missing: $templateFixture" }
    foreach ($workspace in @($consumer,$templateExample,$yolo,$weak,$autoOff)) { New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null; Copy-Item -LiteralPath $fixture -Destination (Join-Path $workspace 'YadgTemplates/template.docx') }
    $schemaFrontMatter = Get-M0015FrontMatterLines $documentedExamples['workspace-schema-v1']
    $autoDetectFrontMatter = Get-M0015FrontMatterLines $documentedExamples['autodetect-schema-v1']
    $autoDetectLines = @($autoDetectFrontMatter[1..($autoDetectFrontMatter.Count - 2)])
    $diagramRootIndex = [Array]::IndexOf($autoDetectLines,'diagrams:')
    if ($diagramRootIndex -lt 0 -or $autoDetectLines[-1] -cne '  autoDetect: true') { throw 'Documented auto-detection YAML example is not the expected schema-v1 mapping.' }
    $producerLines = @($documentedExamples['mermaid-producer-config'].Body -split '\r?\n')
    if ([string]::Join("`n",$producerLines) -notmatch '(?m)^producers:\n  mermaid:\n    executable: mmdc\n    arguments: \[\]$') { throw 'Documented Mermaid producer example no longer matches the supported mmdc process contract.' }
    $producerLines = @($producerLines | ForEach-Object {
        if ($_ -ceq '    executable: mmdc') { '    executable: ' + $producerExecutable }
        elseif ($_ -ceq '    arguments: []') { '    arguments:' }
        else { $_ }
    })
    if ($mermaidArguments.Count) { $producerLines += @($mermaidArguments | ForEach-Object { '      - "' + $_.Replace('"','\"') + '"' }) }
    $mainYaml = @('---') + @($schemaFrontMatter[1..($schemaFrontMatter.Count - 2)]) + @('values:','  document-version: "1.2"','  story-value: "M0015 installed consumer value"') + @($autoDetectLines[$diagramRootIndex..($autoDetectLines.Count - 1)]) + @('publish:','  path: ./Published') + $producerLines + @('---')
    $mainYaml | Set-Content -LiteralPath (Join-Path $consumer 'YADG.md') -Encoding utf8
    $sourcePieces = @('# Architecture {#architecture}','An installed-tool release sample.',$documentedExamples['markdown-unordered-list'].Body,$documentedExamples['markdown-ordered-list'].Body)
    foreach ($exampleId in @('explicit-text-diagram','standalone-arrow-diagram','nested-box-diagram','corner-occlusion-diagram','autodetect-identified-positive','mermaid-source-fence')) { $sourcePieces += $documentedExamples[$exampleId].Body }
    $source = $sourcePieces -join "`n`n"
    Set-Content -LiteralPath (Join-Path $consumer 'content.md') -Value $source -Encoding utf8
    $templateSourcePieces = @($sourcePieces[0],$sourcePieces[1],$documentedExamples['markdown-table'].Body,$documentedExamples['markdown-figure'].Body) + @($sourcePieces[2..($sourcePieces.Count - 1)])
    $templateSource = $templateSourcePieces -join "`n`n"
    $imageDirectory = Join-Path $consumer 'images'; New-Item -ItemType Directory -Force -Path $imageDirectory | Out-Null
    $pixelPng = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aMn8AAAAASUVORK5CYII=')
    [IO.File]::WriteAllBytes((Join-Path $imageDirectory 'context.png'),$pixelPng)
    $templateYamlBody = @($documentedExamples['template-frontmatter-schema'].Body -split '\r?\n')
    $templateFrontMatter = @('---') + $templateYamlBody + @('---')
    $bindingLines = @($documentedExamples['template-list-prototype-binding'].Body -split '\r?\n' | Where-Object { $_ -ne '' })
    foreach ($bindingLine in $bindingLines) { if (-not ($templateYamlBody -ccontains $bindingLine)) { throw "Documented template prototype binding '$bindingLine' is not present in the executed front-matter example." } }
    $templateTags = @($documentedExamples['template-visible-tags'].Body -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $templatePrototypes = @($documentedExamples['template-list-prototype'],$documentedExamples['template-ordered-list-prototype'],$documentedExamples['template-caption-prototype'],$documentedExamples['template-table-caption-prototype'])
    $consumerTemplate = Join-Path $consumer 'YadgTemplates/template.docx'
    Remove-Item -LiteralPath $consumerTemplate -Force
    New-M0015TemplateFromExamples $templateFixture $consumerTemplate $templateFrontMatter $templateTags $templatePrototypes
    $documentedPrefix = @('# Architecture {#architecture}','An installed-tool release sample.',$documentedExamples['markdown-unordered-list'].Body,$documentedExamples['markdown-ordered-list'].Body,$documentedExamples['markdown-table'].Body,$documentedExamples['markdown-figure'].Body)
    $templateDocument = Join-Path $templateExample 'YadgTemplates/template.docx'
    Remove-Item -LiteralPath $templateDocument -Force
    New-M0015TemplateFromExamples $templateFixture $templateDocument $templateFrontMatter $templateTags $templatePrototypes
    Copy-Item -LiteralPath (Join-Path $consumer 'YADG.md') -Destination (Join-Path $templateExample 'YADG.md')
    Set-Content -LiteralPath (Join-Path $templateExample 'content.md') -Value $templateSource -Encoding utf8
    Copy-Item -LiteralPath $imageDirectory -Destination (Join-Path $templateExample 'images') -Recurse
    foreach ($exampleId in @('explicit-text-diagram','standalone-arrow-diagram','nested-box-diagram','corner-occlusion-diagram','autodetect-identified-positive','mermaid-source-fence')) {
        Set-Content -LiteralPath (Join-Path $consumer 'content.md') -Value (($documentedPrefix + $documentedExamples[$exampleId].Body) -join "`n`n") -Encoding utf8
        $exampleCheck = (& $yadg check --workspace $consumer 2>&1 | Out-String)
        if ($LASTEXITCODE -ne 0) { throw "Installed strict check rejected canonical Markdown example '$exampleId': $exampleCheck" }
        if (Test-Path (Join-Path $consumer 'YadgPreWords')) { Remove-Item -LiteralPath (Join-Path $consumer 'YadgPreWords') -Force -Recurse }
        & $yadg build --workspace $consumer
        if ($LASTEXITCODE -ne 0) { throw "Installed strict build rejected canonical Markdown example '$exampleId'." }
    }
    Set-Content -LiteralPath (Join-Path $consumer 'content.md') -Value $templateSource -Encoding utf8
    $check = (& $yadg check --workspace $consumer 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "Installed strict check failed for diagram examples: $check" }
    $list = (& $yadg check --workspace $consumer --list 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $list -notmatch 'detected-flow\s+figure') { throw "Installed check --list did not list identified text diagram as a figure: $list" }
    & $yadg inspect styles --workspace $consumer
    if ($LASTEXITCODE -ne 0) { throw 'Installed inspect styles failed.' }
    & $yadg inspect template --workspace $consumer
    if ($LASTEXITCODE -ne 0) { throw 'Installed inspect template failed.' }
    if (Test-Path (Join-Path $consumer 'YadgPreWords')) { Remove-Item -LiteralPath (Join-Path $consumer 'YadgPreWords') -Force -Recurse }
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
        foreach ($term in @('lower','Gateway','Input','Backend','Inner','A')) { if ($svgAll -notmatch [regex]::Escape($term)) { throw "Generated SVG content omitted expected documented grid text '$term'." } }
        if ($svgAll -notmatch '<path' -or $svgAll -match '<script|foreignObject|(?:href|src)\s*=\s*["''](?:https?:|//|data:)') { throw 'Generated text-diagram SVG lacks arrow geometry or includes active/external content.' }
    } finally { $zip.Dispose() }
    $results.strictAuthoring = 'installed strict check/build, check --list, inspect styles/template passed'
    $results.textDiagrams = [ordered]@{ result='passed'; extractedExamples=@('explicit-text-diagram','standalone-arrow-diagram','nested-box-diagram','corner-occlusion-diagram'); autoDetect='documented identified lowercase text fence'; svgParts=$svgEntries.Count; svgReferences=([regex]::Matches($documentXml,'svgBlip')).Count }
    $results.mermaid = [ordered]@{ result='passed'; producer=$producerKind; executable=$MermaidExecutable; arguments=$mermaidArguments; pinnedPackage="$mermaidPackage@$mermaidVersion"; embeddedPngParts=$pngEntries.Count }

    $templateExampleCheck = (& $yadg check --workspace $templateExample 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "Installed tool rejected the canonical DOCX-template tag/front-matter/prototype examples: $templateExampleCheck" }
    $templateExampleInspection = (& $yadg inspect template --workspace $templateExample 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) { throw "Installed inspect template rejected canonical DOCX-template examples: $templateExampleInspection" }
    foreach ($expected in @('figure-caption','table-caption','bullet-item','number-item','table:interfaces','figure:system-context','value:document-version')) {
        if (($templateExampleCheck + $templateExampleInspection) -notmatch [regex]::Escape($expected)) { throw "Installed DOCX-template example audit did not resolve documented tag/prototype '$expected'." }
    }
    & $yadg build --workspace $templateExample
    if ($LASTEXITCODE -ne 0) { throw 'Installed tool could not build the canonical DOCX-template examples.' }
    $templateOutput = Get-ChildItem -LiteralPath (Join-Path $templateExample 'YadgPreWords') -Filter '*.docx' -File | Select-Object -First 1
    if ($null -eq $templateOutput) { throw 'Canonical DOCX-template example build produced no document.' }
    $templateZip = [IO.Compression.ZipFile]::OpenRead($templateOutput.FullName)
    try {
        $templateDocumentEntry = $templateZip.GetEntry('word/document.xml'); $templateReader = [IO.StreamReader]::new($templateDocumentEntry.Open())
        try { $templateDocumentXml = $templateReader.ReadToEnd() } finally { $templateReader.Dispose() }
        foreach ($field in @('SEQ Figure','SEQ Table')) { if ($templateDocumentXml -notmatch [regex]::Escape($field)) { throw "Documented caption prototype did not preserve a real $field field in the authored DOCX." } }
        if ($templateDocumentXml -match '\{\{(?:section|content|table|figure|value):[^{}]+\}\}') { throw 'A documented template placement/value tag remained unresolved in the authored DOCX.' }
    } finally { $templateZip.Dispose() }
    $results.documentedTemplateExamples = 'passed; extracted guide tag, front-matter, list/caption prototype examples were applied to a real template, checked, and built; figure/table SEQ fields verified'

    $weakYaml = @($autoDetectFrontMatter[0..($autoDetectFrontMatter.Count - 2)]) + @('values:','  story-value: "M0015 installed consumer value"','---')
    $weakYaml | Set-Content -LiteralPath (Join-Path $weak 'YADG.md') -Encoding utf8
    @('# Architecture {#architecture}',$documentedExamples['autodetect-weak-negative'].Body) -join "`n`n" | Set-Content -LiteralPath (Join-Path $weak 'content.md') -Encoding utf8
    $weakResult = (& $yadg check --workspace $weak 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0 -or $weakResult -notmatch 'YADG-MD-UNSUPPORTED') { throw "Auto-detection converted a weak text fence or returned the wrong diagnostic: $weakResult" }
    $results.autoDetectionNegative = [ordered]@{ result='passed'; expected='YADG-MD-UNSUPPORTED'; observed='weak ordinary text fence remains unsupported' }
    $autoOffYaml = @($schemaFrontMatter[0..($schemaFrontMatter.Count - 2)]) + @('values:','  story-value: "M0015 installed consumer value"','---')
    $autoOffYaml | Set-Content -LiteralPath (Join-Path $autoOff 'YADG.md') -Encoding utf8
    @('# Architecture {#architecture}',$documentedExamples['autodetect-identified-positive'].Body) -join "`n`n" | Set-Content -LiteralPath (Join-Path $autoOff 'content.md') -Encoding utf8
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
    $results.documentation = [ordered]@{ result='passed'; canonicalExampleIds=@($documentedExampleSpecs | ForEach-Object id); navigation='repository links, compatibility links, and packaged README guide URLs checked'; exampleExecution='schema-v1/autodetect/Mermaid YAML plus Markdown table, figure, explicit diagram, standalone arrow, nested/corner geometry, positive/negative auto-detection, Mermaid source, and DOCX template tags/frontmatter/prototypes extracted from canonical guides and accepted by installed 1.2 tool' }

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
        installedConsumer=$results; documentationExampleAudit='passed; bounded canonical guide blocks were extracted by stable markers, validated with the installed candidate, and included schema-v1 YAML, Markdown table/figure, explicit text diagrams, standalone arrows, nested/corner geometry, positive/negative auto-detection, real Mermaid config/source, README package navigation, and DOCX tags/frontmatter/prototype authoring.'
        externalPublication='none'
    }
    $evidence | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'release-evidence.json') -Encoding utf8
    Write-Output 'M0015 Tier 4 installed-package consumer validation passed.'
}
finally {
    $env:DOTNET_CLI_HOME = $oldHome; $env:NUGET_PACKAGES = $oldNuget; $env:Path = $oldPath
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
}
