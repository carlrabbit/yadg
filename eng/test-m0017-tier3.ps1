param(
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0017')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$wordExe = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
$loExe = 'C:\Program Files\LibreOffice\program\soffice.exe'
if (-not (Test-Path -LiteralPath $wordExe)) { throw 'M0017 Tier 3 is blocked: desktop Microsoft Word is unavailable in this Windows session.' }
if (-not (Test-Path -LiteralPath $loExe)) { throw 'M0017 Tier 3 is blocked: LibreOffice Writer is unavailable in this Windows session.' }
$fixture = Join-Path $repo 'tests/fixtures/m0010/word-origin.docx'
$provenance = Get-Content (Join-Path $repo 'tests/fixtures/m0010/provenance.json') -Raw | ConvertFrom-Json
$fixtureRecord = $provenance.fixtures | Where-Object { $_.id -eq 'word-origin' } | Select-Object -First 1
if ($null -eq $fixtureRecord) { throw 'M0017 Tier 3 is blocked: Word-origin M0010 fixture provenance is missing.' }
$fixtureHash = (Get-FileHash -LiteralPath $fixture -Algorithm SHA256).Hash.ToUpperInvariant()
if ($fixtureHash -ne ([string]$fixtureRecord.sha256).ToUpperInvariant()) { throw 'M0017 Tier 3 is blocked: Word-origin fixture hash does not match its provenance.' }
$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.exe'
if (-not (Test-Path -LiteralPath $cli)) { throw "Build the Windows CLI before Tier 3: $cli" }
$wordVersion = (Get-Item -LiteralPath $wordExe).VersionInfo.FileVersion
$loVersion = (Get-Item -LiteralPath $loExe).VersionInfo.FileVersion
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } else { $loExe }
$source = @'
# Architecture {#architecture}

Read [the architecture explanation](#architecture), [the system flow](#system-flow), [the interface matrix](#interface-matrix), and [the external guide](https://example.test/yadg).

See [@system-flow] and [@interface-matrix].

![System flow](images/flow.png){#system-flow}

| Component | Protocol |
|---|---|
| Gateway | HTTPS |
| Worker | HTTPS |
{#interface-matrix caption="Interface matrix"}

- unordered parent
  1. ordered child
     - unordered grandchild
1. ordered parent
   - unordered child

> A **quoted** paragraph with [a link](https://example.test/quote).
>
> > Nested quote paragraph.

Fenced literal code:

```text
  keep leading spaces
{#literal-code-line}

code line
```

Indented literal code:

    keep four spaces

Task status [x] ordinary prose stays literal.

- [x] completed task
- [ ] pending task

~~crossed out~~ and `inline-code`.

+---+---+---+
| AAAAA | B |
+---+---+ B +
| D | E | B |
+ D +---+---+
| D | CCCCC |
+---+---+---+
{#merged-matrix}
'@

function Find-DocumentTextRange($Document, [string] $Text, [int] $Start = 0, [int] $End = -1) {
    if ($End -lt 0) { $End = $Document.Content.End }
    $range = $Document.Range($Start, $End)
    $find = $range.Find
    $find.ClearFormatting()
    $find.Text = $Text
    $find.Forward = $true
    $find.Wrap = 0
    if (-not $find.Execute()) { throw "Word fixture edit could not find '$Text'." }
    return $range
}

$temporary = Join-Path ([System.IO.Path]::GetTempPath()) ('yadg-m0017-tier3-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $temporary, $EvidenceRoot | Out-Null
$diagramPath = Join-Path $EvidenceRoot 'system-flow.png'
Add-Type -AssemblyName System.Drawing
$diagram = [System.Drawing.Bitmap]::new(720, 220)
$graphics = [System.Drawing.Graphics]::FromImage($diagram)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::White)
$boxBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(42, 84, 112))
$textBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
$arrowPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(42, 84, 112), 7)
$arrowPen.CustomEndCap = [System.Drawing.Drawing2D.AdjustableArrowCap]::new(8, 8, $true)
$labelFont = [System.Drawing.Font]::new('Arial', 18, [System.Drawing.FontStyle]::Bold)
$protocolFont = [System.Drawing.Font]::new('Arial', 14, [System.Drawing.FontStyle]::Regular)
try {
    $graphics.FillRectangle($boxBrush, 32, 58, 190, 104)
    $graphics.FillRectangle($boxBrush, 498, 58, 190, 104)
    $graphics.DrawString('Gateway', $labelFont, $textBrush, 69, 94)
    $graphics.DrawString('Worker', $labelFont, $textBrush, 548, 94)
    $graphics.DrawLine($arrowPen, 240, 110, 480, 110)
    $graphics.DrawString('HTTPS', $protocolFont, [System.Drawing.Brushes]::Black, 321, 66)
    $diagram.Save($diagramPath, [System.Drawing.Imaging.ImageFormat]::Png)
}
finally {
    $protocolFont.Dispose(); $labelFont.Dispose(); $arrowPen.Dispose(); $textBrush.Dispose(); $boxBrush.Dispose(); $graphics.Dispose(); $diagram.Dispose()
}
$templatePath = Join-Path $EvidenceRoot 'word-authored-reference-template.docx'
Copy-Item -LiteralPath $fixture -Destination $templatePath -Force
$word = $null
$document = $null
$templateCreation = 'failed'
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($templatePath, $false, $false)

    $frontMatterEnd = Find-DocumentTextRange $document '{{/yadg:frontmatter}}'
    $frontMatterEnd.InsertBefore("styles:`r  generatedTable: TableGrid`r  caption: Normal`r  codeInline: Code`rprototypes:`r  figureCaption: figure-caption`r  tableCaption: table-caption`r  unorderedListItem: list-bullet`r  orderedListItem: list-number`r")
    try { $null = $document.Styles.Item('TableGrid') } catch { $null = $document.Styles.Add('TableGrid', 3) }
    try { $codeStyle = $document.Styles.Item('Code') } catch { $codeStyle = $document.Styles.Add('Code', 2) }
    $codeStyle.Font.Name = 'Consolas'

    $firstBody = Find-DocumentTextRange $document 'Word-origin static cover content.'
    $insertAt = [int]$firstBody.Start
    $prototypeText = "{{yadg:prototype:figure-caption}}`rFigure §: {{caption}}`r{{/yadg:prototype:figure-caption}}`r{{yadg:prototype:table-caption}}`rTable §: {{caption}}`r{{/yadg:prototype:table-caption}}`r{{yadg:prototype:list-bullet}}`r{{item}}`r{{/yadg:prototype:list-bullet}}`r{{yadg:prototype:list-number}}`r{{item}}`r{{/yadg:prototype:list-number}}`r"
    $firstBody.Collapse(1)
    $firstBody.InsertBefore($prototypeText)
    $bulletTemplate = $document.ListTemplates.Add($true)
    $numberTemplate = $document.ListTemplates.Add($true)
    for ($level = 1; $level -le 9; $level++) {
        $bulletLevel = $bulletTemplate.ListLevels.Item($level)
        $bulletLevel.NumberFormat = [string][char]0x2022
        $bulletLevel.NumberStyle = 23
        $numberLevel = $numberTemplate.ListLevels.Item($level)
        $numberLevel.NumberFormat = "%$level."
        $numberLevel.NumberStyle = 0
    }
    $bulletItem = Find-DocumentTextRange $document '{{item}}'
    $bulletItem.Paragraphs.Item(1).Range.ListFormat.ApplyListTemplateWithLevel($bulletTemplate, $false, 0, 0, 1)
    $numberItem = Find-DocumentTextRange $document '{{item}}' ([int]$bulletItem.End)
    $numberItem.Paragraphs.Item(1).Range.ListFormat.ApplyListTemplateWithLevel($numberTemplate, $false, 0, 0, 1)
    foreach ($fieldName in @('Figure', 'Table')) {
        $marker = Find-DocumentTextRange $document '§'
        $position = [int]$marker.Start
        $marker.Text = ''
        $fieldRange = $document.Range($position, $position)
        $fieldType = 12 # Word wdFieldSequence
        $fieldCode = "$fieldName \* ARABIC"
        $null = $document.Fields.Add($fieldRange, $fieldType, $fieldCode, $true)
    }
    $duplicateBodyPlacement = Find-DocumentTextRange $document '{{content:architecture}}'
    $duplicateBodyPlacement.Paragraphs.Item(1).Range.Delete()
    $document.Save()
    $document.Close(0)
    $document = $null
    $word.Quit()
    $word = $null
    $templateCreation = 'saved through Microsoft Word document model'
}
finally {
    if ($null -ne $document) { try { $document.Close(0) } catch {} }
    if ($null -ne $word) { try { $word.Quit() } catch {} }
}

$templateSha256 = (Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash.ToUpperInvariant()
$revision = (& git -C $repo rev-parse HEAD).Trim()
$revisionSuffix = $revision.Substring(0, 7)
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$worktree = ((& git -C $repo status --short --untracked-files=all) | Where-Object { $_ -notmatch '^\?\? (\.execution|\.review|artifacts)/|^[ MARCUD?!]{2} (\.execution|\.review|artifacts)/|docs/milestones/M0016\.md' }) -join "`n"
$worktreeHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($worktree)))
$runs = @()
try {
    foreach ($renderer in @('word', 'libreoffice')) {
        $workspace = Join-Path $temporary $renderer
        New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null
        Copy-Item -LiteralPath $templatePath -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
        Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Encoding utf8 -Value @('---','yadg:','  version: 1','values:','  story-value: "M0017 Tier 3"','references:','  headings: number','  figures: label-number','  tables: label-number','---')
        New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'images') | Out-Null
        Copy-Item -LiteralPath $diagramPath -Destination (Join-Path $workspace 'images/flow.png')
        Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Encoding utf8 -Value $source

        & $cli check --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0017 check failed for real $renderer rendering." }
        & $cli build --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0017 build failed for real $renderer rendering." }
        $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
        if ($null -eq $authored) { throw "M0016 authored DOCX is missing for $renderer." }
        $renderArgs = @('render', '--workspace', $workspace, '--renderer', $renderer)
        if ($renderer -eq 'libreoffice') { $renderArgs += @('--renderer-path', $loPath) }
        & $cli @renderArgs
        if ($LASTEXITCODE -ne 0) { throw "M0017 real $renderer finalization failed." }
        $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
        if ($null -eq $finalized) { throw "M0016 finalized DOCX is missing for $renderer." }

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($finalized.FullName)
        try {
            $entry = $zip.GetEntry('word/document.xml')
            if ($null -eq $entry) { throw "Finalized $renderer DOCX has no main document part." }
            $reader = [IO.StreamReader]::new($entry.Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $manager = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $manager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
            $manager.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $allText = ($xml.SelectNodes('//w:t', $manager) | ForEach-Object { $_.InnerText }) -join ''
            foreach ($expected in @('the architecture explanation','the system flow','the interface matrix','the external guide')) {
                if (-not $allText.Contains($expected, [StringComparison]::Ordinal)) { throw "Finalized $renderer output lost authored link wording '$expected'." }
            }
            if (-not $allText.Contains('See Figure 1 and Table 1.', [StringComparison]::Ordinal)) { throw "Finalized $renderer output did not update the configured figure/table references to visible label-number results." }
            $instructions = @($xml.SelectNodes('//w:instrText', $manager) | ForEach-Object { $_.InnerText }) + @($xml.SelectNodes('//w:fldSimple', $manager) | ForEach-Object { $_.GetAttribute('instr', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') })
            if (@($instructions | Where-Object { $_ -match '^\s*SEQ\s+Figure\b' }).Count -lt 1 -or @($instructions | Where-Object { $_ -match '^\s*SEQ\s+Table\b' }).Count -lt 1) { throw "Finalized $renderer output is missing template-owned figure/table SEQ fields. Instructions: $($instructions -join ' | ')" }
            if ($allText.Contains('Error! Bookmark not defined.', [StringComparison]::Ordinal) -or $allText.Contains('Error! Not a valid bookmark', [StringComparison]::Ordinal)) { throw "Finalized $renderer output contains a broken generated reference field." }
            foreach ($expected in @('unordered parent','ordered child','unordered grandchild','ordered parent','Nested quote paragraph.','keep leading spaces','{#literal-code-line}','keep four spaces','completed task','pending task','crossed out','inline-code','AAAAA','CCCCC')) {
                if (-not $allText.Contains($expected, [StringComparison]::Ordinal)) { throw "Finalized $renderer output lost M0017 content '$expected'." }
            }
            if (@($xml.SelectNodes('//w:strike', $manager)).Count -lt 1) { throw "Finalized $renderer output lost strike formatting." }
            if (@($xml.SelectNodes('//w:gridSpan[@w:val="2"]', $manager)).Count -lt 1 -or @($xml.SelectNodes('//w:vMerge', $manager)).Count -lt 1) { throw "Finalized $renderer output lost grid-table merge structures." }
            if (@($instructions | Where-Object { $_ -match 'REF\s+' }).Count -lt 2) { throw "Finalized $renderer output is missing generated figure/table REF fields." }
            $anchors = @($xml.SelectNodes('//w:hyperlink[@w:anchor]', $manager) | ForEach-Object { [string]$_.GetAttribute('anchor', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') })
            $bookmarkNames = @($xml.SelectNodes('//w:bookmarkStart', $manager) | ForEach-Object { [string]$_.GetAttribute('name', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') })
            foreach ($anchor in $anchors) { if (-not $bookmarkNames.Contains($anchor)) { throw "Finalized $renderer hyperlink points to missing bookmark '$anchor'." } }
            if ($anchors.Count -lt 3) { throw "Finalized $renderer output lost one or more semantic fragment hyperlinks." }
            $external = @($xml.SelectNodes('//w:hyperlink[@r:id]', $manager))
            if ($external.Count -lt 1) { throw "Finalized $renderer output lost the external hyperlink." }
        }
        finally { $zip.Dispose() }

        $name = "$renderer-finalized-$revisionSuffix.docx"
        $destination = Join-Path $EvidenceRoot $name
        Copy-Item -LiteralPath $finalized.FullName -Destination $destination -Force
        $relativeArtifact = [IO.Path]::GetRelativePath($repo, $destination).Replace([char]92, [char]47)
        $runs += [ordered]@{ renderer = $renderer; runtimeVersion = if ($renderer -eq 'word') { $wordVersion } else { $loVersion }; templateSha256 = $templateSha256; authoredSha256 = (Get-FileHash -LiteralPath $authored.FullName -Algorithm SHA256).Hash.ToUpperInvariant(); finalized = $relativeArtifact; finalizedSha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToUpperInvariant(); semanticHyperlinks = $anchors.Count; generatedReferenceFields = @($instructions | Where-Object { $_ -match 'REF\s+' }).Count; result = 'passed' }
    }
    $relativeTemplate = [IO.Path]::GetRelativePath($repo, $templatePath).Replace([char]92, [char]47)
    $evidence = [ordered]@{ milestone = 'M0017'; repositoryRevision = $revision; worktreeStatusSha256 = $worktreeHash; os = $os; fixture = 'tests/fixtures/m0010/word-origin.docx'; fixtureSha256 = $fixtureHash; templateCreation = $templateCreation; templatePath = $relativeTemplate; templateSha256 = $templateSha256; wordVersion = $wordVersion; libreOfficeVersion = $loVersion; sourceSha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($source))); runs = $runs }
    $evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json') -Encoding utf8
    Write-Output 'M0017 real Word and LibreOffice finalization passed.'
}
finally { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
