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
foreach ($fixture in ($Fixtures | Where-Object { [string]$_.id -notmatch 'recovery' })) {
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

$recoveryRuns = @()
foreach ($scenario in @('related-list-recovery','builtin-list-recovery')) {
    $fixture = $Fixtures | Where-Object { [string]$_.id -eq $scenario } | Select-Object -First 1
    if ($null -eq $fixture) { throw "Missing dedicated Word-origin recovery fixture '$scenario'." }
    $templatePath = Join-Path $RepositoryRoot ([string]$fixture.path)
    $workspace = Join-Path $TemporaryRoot "$scenario-workspace"
    New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null
    Copy-Item -LiteralPath $templatePath -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
    Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Value @('---','yadg:','  version: 1','---')
    Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('# Introduction {#introduction}','','- Related bullet one','- Related bullet two','','1. Ordered one','2. Ordered two')
    $strictOutput = & $Cli check --workspace $workspace 2>&1
    if ($LASTEXITCODE -eq 0) { throw "$scenario strict check unexpectedly passed." }
    if (($strictOutput -join "`n") -notmatch 'near:') { throw "$scenario strict DOCX diagnostic omitted searchable text." }
    $inspection = & $Cli inspect template --workspace $workspace
    if ($LASTEXITCODE -ne 0) { throw "$scenario inspect template failed: $inspection" }
    if ($scenario -eq 'related-list-recovery' -and ($inspection -join "`n") -notmatch 'yolo candidate: List Bullet') { throw 'Inspection did not select the related Word bullet presentation.' }
    if ($scenario -eq 'builtin-list-recovery' -and ($inspection -join "`n") -notmatch 'builtin:unordered-list-v1') { throw 'Inspection did not preview built-in unordered numbering.' }
    $buildOutput = & $Cli build --workspace $workspace --yolo 2>&1
    if ($LASTEXITCODE -ne 0) { throw "$scenario YOLO build failed: $buildOutput" }
    if ($scenario -eq 'builtin-list-recovery' -and ($buildOutput -join "`n") -notmatch 'builtin:unordered-list-v1') { throw 'YOLO build did not report built-in list degradation.' }
    $renderers = if ($scenario -eq 'builtin-list-recovery') { @('word','libreoffice') } else { @('word') }
    foreach ($renderer in $renderers) {
        if ($scenario -eq 'related-list-recovery') {
            $prime = & $Cli render --workspace $workspace --renderer libreoffice --renderer-path $loPath 2>&1
            if ($LASTEXITCODE -ne 0) { throw "Initial LibreOffice result-set render failed before fallback check: $prime" }
            if (-not (Test-Path -LiteralPath (Join-Path $workspace 'YadgPdfs/template.pdf'))) { throw 'LibreOffice did not create the PDF needed for Word fallback freshness proof.' }
            $render = & $Cli render --workspace $workspace --renderer libreoffice --renderer-path (Join-Path $TemporaryRoot 'missing-soffice.exe') --yolo 2>&1
            if ($LASTEXITCODE -ne 0 -or ($render -join "`n") -notmatch 'requested=libreoffice; actual=word' -or ($render -join "`n") -notmatch 'degradations=1') { throw "Renderer fallback provenance/accounting failed: $render" }
            if (@(Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPdfs') -Filter '*.pdf').Count -ne 0) { throw 'Word fallback retained a stale LibreOffice PDF.' }
        } else {
            $rendererArgs = @('render','--workspace',$workspace,'--renderer',$renderer)
            if ($renderer -eq 'libreoffice') { $rendererArgs += @('--renderer-path',$loPath) }
            $render = & $Cli @rendererArgs 2>&1
            if ($LASTEXITCODE -ne 0) { throw "$scenario finalization through $renderer failed: $render" }
        }
        $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
        if (-not $finalized) { throw "$scenario finalized DOCX missing through $renderer." }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($finalized.FullName)
        try {
            $reader = [IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open()); try { [xml]$docXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $ns = [Xml.XmlNamespaceManager]::new($docXml.NameTable); $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            if ($scenario -eq 'builtin-list-recovery' -and $docXml.SelectNodes('//w:pPr/w:numPr', $ns).Count -lt 4) { throw "Built-in real numbering failed through $renderer." }
            if ($scenario -eq 'related-list-recovery' -and $docXml.SelectNodes('//w:pPr/w:pStyle[@w:val="ListBullet"]', $ns).Count -lt 2) { throw 'Related Word list style was not used.' }
        } finally { $zip.Dispose() }
        $artifactName = "yolo-$scenario-$renderer.docx"
        Copy-Item $finalized.FullName (Join-Path $EvidenceRoot $artifactName) -Force
        $recoveryRuns += [ordered]@{ scenario = $scenario; sourceFixture = $fixture.id; sourceFixtureSha256 = $fixture.sha256; renderer = $renderer; artifact = $artifactName; result = 'passed' }
    }
}

$staleFixture = $Fixtures | Where-Object { [string]$_.id -eq 'style-resolution' } | Select-Object -First 1
$staleWorkspace = Join-Path $TemporaryRoot 'stale-preword-replacement'
$staleTemplateRoot = Join-Path $staleWorkspace 'YadgTemplates'
New-Item -ItemType Directory -Force -Path $staleTemplateRoot | Out-Null
$staleSource = Join-Path $RepositoryRoot ([string]$staleFixture.path)
Copy-Item -LiteralPath $staleSource -Destination (Join-Path $staleTemplateRoot 'a.docx')
Copy-Item -LiteralPath $staleSource -Destination (Join-Path $staleTemplateRoot 'b.docx')
Set-Content -LiteralPath (Join-Path $staleWorkspace 'YADG.md') -Value @('---','yadg:','  version: 1','markdown:','  codeInline: style','---')
Set-Content -LiteralPath (Join-Path $staleWorkspace 'content.md') -Value @('# Introduction {#introduction}','','A styled `code sample`.','','- Bullet item','','1. Ordered item')
& $Cli build --workspace $staleWorkspace
if ($LASTEXITCODE -ne 0) { throw 'Initial A+B strict build failed in stale-output Tier 3 scenario.' }
if ((Get-ChildItem -LiteralPath (Join-Path $staleWorkspace 'YadgPreWords') -Filter '*.docx').Count -ne 2) { throw 'Initial stale-output scenario did not produce A+B.' }
$bPath = Join-Path $staleTemplateRoot 'b.docx'
$lockedTemplate = [IO.File]::Open($bPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try {
    & $Cli build --workspace $staleWorkspace --yolo
    if ($LASTEXITCODE -ne 0) { throw 'YOLO replacement build failed while template B was recoverably unavailable.' }
} finally { $lockedTemplate.Dispose() }
$preWordNames = @(Get-ChildItem -LiteralPath (Join-Path $staleWorkspace 'YadgPreWords') -Filter '*.docx' | Select-Object -ExpandProperty Name)
if ($preWordNames.Count -ne 1 -or $preWordNames[0] -ne 'a.docx') { throw "Successful YOLO replacement retained a stale/skipped PreWord: $($preWordNames -join ', ')." }
& $Cli render --workspace $staleWorkspace --renderer word
if ($LASTEXITCODE -ne 0) { throw 'Word renderer failed after successful YOLO output-set replacement.' }
$finalNames = @(Get-ChildItem -LiteralPath (Join-Path $staleWorkspace 'YadgWords') -Filter '*.docx' | Select-Object -ExpandProperty Name)
if ($finalNames.Count -ne 1 -or $finalNames[0] -ne 'a.docx') { throw "Renderer processed a stale/skipped PreWord: $($finalNames -join ', ')." }
$recoveryRuns += [ordered]@{ scenario = 'stale-preword-replacement-render'; sourceFixture = $staleFixture.id; sourceFixtureSha256 = $staleFixture.sha256; initialTemplateSet = @('a.docx','b.docx'); skippedTemplate = 'b.docx'; finalPreWords = $preWordNames; finalized = $finalNames; renderer = 'word'; result = 'passed' }

$renderSetWorkspace = Join-Path $TemporaryRoot 'render-result-set-ownership'
$renderTemplateRoot = Join-Path $renderSetWorkspace 'YadgTemplates'
New-Item -ItemType Directory -Force -Path $renderTemplateRoot | Out-Null
Copy-Item -LiteralPath $staleSource -Destination (Join-Path $renderTemplateRoot 'a.docx')
Copy-Item -LiteralPath $staleSource -Destination (Join-Path $renderTemplateRoot 'b.docx')
Set-Content -LiteralPath (Join-Path $renderSetWorkspace 'YADG.md') -Value @('---','yadg:','  version: 1','markdown:','  codeInline: style','---')
Set-Content -LiteralPath (Join-Path $renderSetWorkspace 'content.md') -Value @('# Introduction {#introduction}','','A styled `code sample`.','','- Bullet item','','1. Ordered item')
& $Cli build --workspace $renderSetWorkspace
if ($LASTEXITCODE -ne 0) { throw 'A+B render freshness build failed.' }
& $Cli render --workspace $renderSetWorkspace --renderer libreoffice --renderer-path $loPath
if ($LASTEXITCODE -ne 0) { throw 'Initial A+B LibreOffice render failed.' }
Remove-Item -LiteralPath (Join-Path $renderSetWorkspace 'YadgPreWords/b.docx')
& $Cli render --workspace $renderSetWorkspace --renderer libreoffice --renderer-path $loPath
if ($LASTEXITCODE -ne 0) { throw 'A-only LibreOffice rerender failed.' }
$loWordNames = @(Get-ChildItem -LiteralPath (Join-Path $renderSetWorkspace 'YadgWords') -Filter '*.docx' | Select-Object -ExpandProperty Name)
$loPdfNames = @(Get-ChildItem -LiteralPath (Join-Path $renderSetWorkspace 'YadgPdfs') -Filter '*.pdf' | Select-Object -ExpandProperty Name)
if ($loWordNames.Count -ne 1 -or $loWordNames[0] -ne 'a.docx' -or $loPdfNames.Count -ne 1 -or $loPdfNames[0] -ne 'a.pdf') { throw 'LibreOffice rerender retained stale finalized artifacts.' }
& $Cli render --workspace $renderSetWorkspace --renderer word
if ($LASTEXITCODE -ne 0) { throw 'Word render following LibreOffice failed.' }
$wordPdfNames = @(Get-ChildItem -LiteralPath (Join-Path $renderSetWorkspace 'YadgPdfs') -Filter '*.pdf' | Select-Object -ExpandProperty Name)
if ($wordPdfNames.Count -ne 0) { throw 'Word success retained stale LibreOffice PDFs.' }
$committedHash = (Get-FileHash -LiteralPath (Join-Path $renderSetWorkspace 'YadgWords/a.docx') -Algorithm SHA256).Hash
Set-Content -LiteralPath (Join-Path $renderSetWorkspace 'YadgPreWords/a.docx') -Value 'not a DOCX'
$failedRender = & $Cli render --workspace $renderSetWorkspace --renderer word 2>&1
if ($LASTEXITCODE -eq 0) { throw 'Corrupt PreWord unexpectedly rendered successfully.' }
$afterFailureHash = (Get-FileHash -LiteralPath (Join-Path $renderSetWorkspace 'YadgWords/a.docx') -Algorithm SHA256).Hash
if ($afterFailureHash -ne $committedHash) { throw 'Failed renderer attempt changed the prior finalized result set.' }
$recoveryRuns += [ordered]@{ scenario = 'finalized-result-set-ownership'; initial = @('a.docx','b.docx'); afterLibreOfficeRerender = $loWordNames; pdfAfterLibreOffice = $loPdfNames; pdfAfterWord = $wordPdfNames; failedAttemptPreservedPrior = $afterFailureHash -eq $committedHash; result = 'passed' }
[ordered]@{ repositoryRevision = $revision; worktreeDirty = [bool]$worktreeStatus; worktreeStatusSha256 = $statusHash; os = $os; wordVersion = $WordVersion; libreOfficeVersion = $LibreOfficeVersion; fixtures = $Fixtures; runs = $runs; recoveryRuns = $recoveryRuns } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json')
$global:LASTEXITCODE = 0
Write-Output 'M0012 Tier 3 real Word/LibreOffice authoring matrix passed.'
