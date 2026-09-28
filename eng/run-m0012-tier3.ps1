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
[ordered]@{ repositoryRevision = $revision; worktreeDirty = [bool]$worktreeStatus; worktreeStatusSha256 = $statusHash; os = $os; wordVersion = $WordVersion; libreOfficeVersion = $LibreOfficeVersion; fixtures = $Fixtures; runs = $runs } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json')
Write-Output 'M0012 Tier 3 real Word/LibreOffice authoring matrix passed.'
