param(
    [Parameter(Mandatory=$true)][string] $RepositoryRoot,
    [Parameter(Mandatory=$true)][string] $Cli,
    [Parameter(Mandatory=$true)][string] $EvidenceRoot,
    [Parameter(Mandatory=$true)][string] $TemporaryRoot,
    [Parameter(Mandatory=$true)][object[]] $Fixtures,
    [Parameter(Mandatory=$true)][string] $WordVersion,
    [Parameter(Mandatory=$true)][string] $LibreOfficeVersion
)
$ErrorActionPreference = 'Stop'
$runs = @()
$revision = (& git -C $RepositoryRoot rev-parse HEAD).Trim()
$worktreeStatus = (& git -C $RepositoryRoot status --short --untracked-files=all) -join "`n"
$statusBytes = [Text.Encoding]::UTF8.GetBytes($worktreeStatus)
$statusHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($statusBytes))
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } else { (Get-Command soffice.exe).Source }
foreach ($fixture in $Fixtures) {
    if ([string]$fixture.application -ne 'Microsoft Word') { throw "Unsupported M0012 fixture origin '$($fixture.application)'." }
    $templatePath = Join-Path $RepositoryRoot ([string]$fixture.path)
    $inspectRoot = Join-Path $TemporaryRoot ("inspect-" + [string]$fixture.id)
    New-Item -ItemType Directory -Force -Path (Join-Path $inspectRoot 'YadgTemplates') | Out-Null
    Copy-Item -LiteralPath $templatePath -Destination (Join-Path $inspectRoot 'YadgTemplates/template.docx')
    $inspection = & $Cli inspect styles --workspace $inspectRoot
    if ($LASTEXITCODE -ne 0) { throw "inspect styles failed for $($fixture.id): $inspection" }
    if ([string]$fixture.id -eq 'style-resolution') {
        if (($inspection -join "`n") -notmatch 'paragraph "User Bullets" id=UserBullets .*numbering=style') { throw 'Word style-numbering relationship was not reported for the custom inherited list style.' }
        if (($inspection -join "`n") -notmatch 'paragraph "Heading 1" id=Heading1') { throw 'The Word-visible heading name/internal ID distinction was not reported.' }
        $templateInspection = & $Cli inspect template --workspace $inspectRoot
        if ($LASTEXITCODE -ne 0 -or ($templateInspection -join "`n") -notmatch 'placement content:introduction') { throw 'inspect template did not expose Word-origin placement structure.' }
        if (($templateInspection -join "`n") -notmatch 'role unordered:') { throw 'inspect template did not explain list presentation resolution.' }
    }
    foreach ($renderer in @('word','libreoffice')) {
        $workspace = Join-Path $TemporaryRoot ("$($fixture.id)-$renderer")
        New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null
        Copy-Item -LiteralPath $templatePath -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
        if ([string]$fixture.id -eq 'style-resolution') {
            $yadg = @('---','yadg:','  version: 1','markdown:','  codeInline: style','---')
        } else {
            $yadg = @('---','yadg:','  version: 1','markdown:','  codeInline: style','---')
        }
        Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Value $yadg
        Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('# Introduction {#introduction}','','A styled `code sample`.','','- Bullet **strong** item','- Second bullet','','1. Ordered first','2. Ordered second')
        & $Cli check --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0012 check failed for $($fixture.id) -> $renderer" }
        & $Cli build --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0012 build failed for $($fixture.id) -> $renderer" }
        $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
        if (-not $authored) { throw "M0012 authored DOCX missing for $($fixture.id) -> $renderer" }
        $renderArgs = @('render','--workspace',$workspace,'--renderer',$renderer)
        if ($renderer -eq 'libreoffice') { $renderArgs += @('--renderer-path',$loPath) }
        & $Cli @renderArgs
        if ($LASTEXITCODE -ne 0) { throw "M0012 finalization failed for $($fixture.id) -> $renderer" }
        $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
        if (-not $finalized) { throw "M0012 finalized DOCX missing for $($fixture.id) -> $renderer" }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($finalized.FullName)
        try {
            $xml = ($zip.Entries | Where-Object { $_.FullName -like '*.xml' } | ForEach-Object { $reader = [IO.StreamReader]::new($_.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }) -join "`n"
            $documentReader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
            try { $documentText = $documentReader.ReadToEnd() } finally { $documentReader.Dispose() }
        } finally { $zip.Dispose() }
        if ($xml -match '\{\{(item|content|yadg:prototype)') { throw "A YADG prototype/control placeholder remains for $($fixture.id) -> $renderer" }
        [xml]$documentXml = $documentText
        $ns = [Xml.XmlNamespaceManager]::new($documentXml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        if (-not $documentXml.SelectSingleNode("//w:r[w:rPr/w:rStyle[@w:val='CodeCharacter']][w:t[contains(.,'code sample')]]", $ns)) { throw "Inline-code character style did not survive for $($fixture.id) -> $renderer" }
        foreach ($expectedText in @('Bullet', 'Ordered first')) {
            $paragraph = $documentXml.SelectNodes('//w:p', $ns) | Where-Object { (($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '') -like "*$expectedText*" } | Select-Object -First 1
            if (-not $paragraph) { throw "List content did not survive for $($fixture.id) -> $renderer ($expectedText)" }
            if ([string]$fixture.id -eq 'list-prototypes' -and -not $paragraph.SelectSingleNode('w:pPr/w:numPr', $ns)) { throw "Prototype list numbering did not survive for $renderer ($expectedText)" }
            if ([string]$fixture.id -eq 'style-resolution') {
                $expectedStyle = if ($expectedText -eq 'Bullet') { 'UserBullets' } else { 'ListNumber' }
                if ($paragraph.SelectSingleNode("w:pPr/w:pStyle[@w:val='$expectedStyle']", $ns) -eq $null) { throw "Style-based list paragraph lost its selected style for $renderer ($expectedText)" }
            }
        }
        New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
        $name = "$($fixture.id)-$renderer.docx"
        Copy-Item -LiteralPath $finalized.FullName -Destination (Join-Path $EvidenceRoot $name) -Force
        $runs += [ordered]@{ fixture = $fixture.id; fixtureSha256 = $fixture.sha256; renderer = $renderer; authoredSha256 = (Get-FileHash $authored.FullName -Algorithm SHA256).Hash; finalizedSha256 = (Get-FileHash $finalized.FullName -Algorithm SHA256).Hash; result = 'passed' }
    }
}

function New-RecoveryTemplate([string] $Source, [string] $Destination, [bool] $RemoveListResources, [bool] $RemoveDefaultBulletStyle) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $sourceArchive = [IO.Compression.ZipFile]::OpenRead($Source)
    $targetArchive = [IO.Compression.ZipFile]::Open($Destination, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $sourceArchive.Entries) {
            $target = $targetArchive.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
            $sourceStream = $entry.Open(); $targetStream = $target.Open()
            try {
                if ($entry.FullName -eq 'word/document.xml' -or (($RemoveListResources -or $RemoveDefaultBulletStyle) -and $entry.FullName -eq 'word/styles.xml')) {
                    $reader = [IO.StreamReader]::new($sourceStream); try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    if ($entry.FullName -eq 'word/document.xml') {
                        foreach ($node in $xml.SelectNodes('//w:t', $ns)) { if ($node.InnerText.Contains('User Bullets')) { $node.InnerText = $node.InnerText.Replace('User Bullets', 'Missing Bullet') } }
                        if ($RemoveListResources) {
                            foreach ($node in @($xml.SelectNodes('//w:numPr', $ns))) { $node.ParentNode.RemoveChild($node) | Out-Null }
                            foreach ($node in @($xml.SelectNodes('//w:pStyle[@w:val="ListBullet" or @w:val="ListNumber" or @w:val="UserBullets"]', $ns))) { $node.ParentNode.RemoveChild($node) | Out-Null }
                        }
                    } else {
                        $styleQuery = if ($RemoveListResources) { '//w:style[@w:styleId="ListBullet" or @w:styleId="ListNumber" or @w:styleId="UserBullets"]' } else { '//w:style[@w:styleId="UserBullets"]' }
                        foreach ($node in @($xml.SelectNodes($styleQuery, $ns))) { $node.ParentNode.RemoveChild($node) | Out-Null }
                    }
                    if ($entry.FullName -eq 'word/numbering.xml' -and $RemoveListResources) {
                        $reader = [IO.StreamReader]::new($sourceStream); try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                        foreach ($node in @($xml.SelectNodes('//w:num | //w:abstractNum', $ns))) { $node.ParentNode.RemoveChild($node) | Out-Null }
                    }
                    $xml.Save($targetStream)
                } elseif ($entry.FullName -eq 'word/numbering.xml' -and ($RemoveListResources -or $RemoveDefaultBulletStyle)) {
                    $reader = [IO.StreamReader]::new($sourceStream); try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    if ($RemoveListResources) { foreach ($node in @($xml.SelectNodes('//w:num | //w:abstractNum', $ns))) { $node.ParentNode.RemoveChild($node) | Out-Null } }
                    elseif ($RemoveDefaultBulletStyle) { foreach ($node in @($xml.SelectNodes('//w:pStyle[@w:val="ListBullet"]', $ns))) { $node.SetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'UserBullets') } }
                    $xml.Save($targetStream)
                } else { $sourceStream.CopyTo($targetStream) }
            } finally { $sourceStream.Dispose(); $targetStream.Dispose() }
        }
    } finally { $targetArchive.Dispose(); $sourceArchive.Dispose() }
}

$baseFixture = $Fixtures | Where-Object { [string]$_.id -eq 'style-resolution' } | Select-Object -First 1
$basePath = Join-Path $RepositoryRoot ([string]$baseFixture.path)
$recoveryRuns = @()
foreach ($scenario in @('related-list','builtin-lists')) {
    $removeResources = $scenario -eq 'builtin-lists'
    $derivedTemplate = Join-Path $TemporaryRoot "$scenario-template.docx"
    New-RecoveryTemplate $basePath $derivedTemplate $removeResources ($scenario -eq 'related-list')
    $workspace = Join-Path $TemporaryRoot "$scenario-workspace"
    New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null
    Copy-Item $derivedTemplate (Join-Path $workspace 'YadgTemplates/template.docx')
    Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Value @('---','yadg:','  version: 1','markdown:','  codeInline: style','---')
    Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('# Introduction {#introduction}','','- Related bullet one','- Related bullet two','','1. Ordered one','2. Ordered two')
    $strictOutput = & $Cli check --workspace $workspace 2>&1
    if ($LASTEXITCODE -eq 0) { throw "$scenario strict check unexpectedly passed; correction requires strict/YOLO paired evidence." }
    if (($strictOutput -join "`n") -notmatch 'near:') { throw "$scenario strict DOCX diagnostic omitted searchable nearby text/context." }
    $yoloOutput = & $Cli build --yolo --workspace $workspace 2>&1
    if ($LASTEXITCODE -ne 0) { throw "$scenario YOLO build failed: $yoloOutput" }
    if ($scenario -eq 'builtin-lists' -and ($yoloOutput -join "`n") -notmatch 'builtin:unordered-list-v1') { throw 'YOLO did not report its built-in unordered-list example.' }
    $inspection = & $Cli inspect template --workspace $workspace
    if ($LASTEXITCODE -ne 0 -or ($inspection -join "`n") -notmatch 'strict=unresolved') { throw "$scenario inspect template did not show strict failure/fallback preview." }
    if ($scenario -eq 'related-list' -and ($inspection -join "`n") -notmatch 'yolo candidate: List Bullet') { throw 'YOLO did not select the related Word bullet presentation deterministically.' }
    $renderers = if ($scenario -eq 'builtin-lists') { @('word','libreoffice') } else { @('word') }
    foreach ($renderer in $renderers) {
        if ($scenario -eq 'related-list') {
            $render = & $Cli render --workspace $workspace --renderer libreoffice --renderer-path (Join-Path $TemporaryRoot 'missing-soffice.exe') --yolo 2>&1
            if ($LASTEXITCODE -ne 0 -or ($render -join "`n") -notmatch 'requested=libreoffice; actual=word' -or ($render -join "`n") -notmatch 'degradations=1') { throw "YOLO renderer substitution did not truthfully report requested LibreOffice/actual Word: $render" }
        } else {
            $rendererArgs = @('render','--workspace',$workspace,'--renderer',$renderer)
            if ($renderer -eq 'libreoffice') { $rendererArgs += @('--renderer-path',$loPath) }
            $render = & $Cli @rendererArgs 2>&1
            if ($LASTEXITCODE -ne 0) { throw "$scenario finalization failed through $renderer`: $render" }
        }
        $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($finalized.FullName)
        try {
            $read = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open()); try { [xml]$docXml = $read.ReadToEnd() } finally { $read.Dispose() }
            $ns = [Xml.XmlNamespaceManager]::new($docXml.NameTable); $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            if ($scenario -eq 'builtin-lists' -and $docXml.SelectNodes('//w:pPr/w:numPr', $ns).Count -lt 4) { throw "$scenario list numbering did not survive $renderer finalization." }
            if ($scenario -eq 'related-list' -and $docXml.SelectNodes('//w:pPr/w:pStyle[@w:val="ListBullet"]', $ns).Count -lt 2) { throw "$scenario did not preserve the selected related Word list style through $renderer finalization." }
            if ($removeResources -or $scenario -eq 'related-list') {
                $numberingEntry = $zip.GetEntry('word/numbering.xml'); $reader = [IO.StreamReader]::new($numberingEntry.Open()); try { [xml]$numXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
                $numNs = [Xml.XmlNamespaceManager]::new($numXml.NameTable); $numNs.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                if ($removeResources -and (-not $numXml.SelectSingleNode('//w:abstractNum/w:lvl/w:numFmt[@w:val="bullet"]', $numNs) -or -not $numXml.SelectSingleNode('//w:abstractNum/w:lvl/w:numFmt[@w:val="decimal"]', $numNs))) { throw "Built-in ordered/unordered real numbering did not survive $renderer finalization." }
                if ($scenario -eq 'related-list' -and -not $numXml.SelectSingleNode('//w:abstractNum/w:lvl[w:pStyle[@w:val="ListBullet"]]/w:numFmt[@w:val="bullet"]', $numNs)) { throw "Related Word style's inherited real bullet numbering did not survive $renderer finalization." }
            }
        } finally { $zip.Dispose() }
        $artifactName = "yolo-$scenario-$renderer.docx"
        Copy-Item $finalized.FullName (Join-Path $EvidenceRoot $artifactName) -Force
        $recoveryRuns += [ordered]@{ scenario = $scenario; sourceFixture = $baseFixture.id; sourceFixtureSha256 = $baseFixture.sha256; derivedTemplateSha256 = (Get-FileHash $derivedTemplate -Algorithm SHA256).Hash; renderer = $renderer; artifact = $artifactName; result = 'passed' }
    }
}
[ordered]@{ repositoryRevision = $revision; worktreeDirty = [bool]$worktreeStatus; worktreeStatusSha256 = $statusHash; os = $os; wordVersion = $WordVersion; libreOfficeVersion = $LibreOfficeVersion; fixtures = $Fixtures; runs = $runs; recoveryRuns = $recoveryRuns } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json')
Write-Output 'M0012 Tier 3 real Word/LibreOffice authoring matrix passed.'
