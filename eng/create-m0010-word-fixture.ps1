$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$fixtureDir = Join-Path $repo 'tests/fixtures/m0010'
New-Item -ItemType Directory -Force -Path $fixtureDir | Out-Null
$path = Join-Path $fixtureDir 'word-origin.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $null
try {
    $doc = $word.Documents.Add()
    $doc.PageSetup.TopMargin = 54
    $doc.PageSetup.BottomMargin = 54
    $doc.PageSetup.LeftMargin = 72
    $doc.PageSetup.RightMargin = 72
    $doc.Sections.Item(1).PageSetup.DifferentFirstPageHeaderFooter = $true
    $doc.Sections.Item(1).PageSetup.OddAndEvenPagesHeaderFooter = $true

    $addParagraph = { param([string]$value) $r = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $r.InsertAfter($value + [char]13) }
    foreach ($line in @('{{yadg:frontmatter}}','version: 1','styles:','  headings:','    1: Heading1','    2: Heading2','    3: Heading3','    4: Heading4','    5: Heading5','    6: Heading6','    7: Heading7','    8: Heading8','    9: Heading9','{{/yadg:frontmatter}}','Word-origin static cover content.','Template-owned section','{{section:architecture}}','Template-owned paragraph between placements.','Second template-owned section','{{content:architecture}}','Word-origin static closing content.')) { & $addParagraph $line }
    foreach ($paragraph in $doc.Paragraphs) {
        if ($paragraph.Range.Text.Trim() -in @('Template-owned section','Second template-owned section')) { $paragraph.Style = $doc.Styles.Item('Heading 4') }
    }

    $tableRange = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $table = $doc.Tables.Add($tableRange, 2, 2)
    $table.Cell(1,1).Range.Text = 'Template table'
    $table.Cell(1,2).Range.Text = '{{value:story-value}}'
    $table.Cell(2,1).Range.Text = 'Preserved'
    $table.Cell(2,2).Range.Text = 'Layout'

    $footRange = $doc.Paragraphs.Item(2).Range
    $footRange.Collapse(0)
    $footnote = $doc.Footnotes.Add($footRange); $footnote.Range.Text = 'Footnote {{value:story-value}}'
    $endnote = $doc.Endnotes.Add($footRange); $endnote.Range.Text = 'Endnote {{value:story-value}}'
    $doc.Comments.Add($footRange, 'Comment {{value:story-value}}') | Out-Null

    foreach ($section in $doc.Sections) {
        foreach ($hf in @($section.Headers.Item(1), $section.Footers.Item(1), $section.Headers.Item(2), $section.Footers.Item(2), $section.Headers.Item(3), $section.Footers.Item(3))) {
            $hf.Range.Text = "Static header/footer {{value:story-value}} | "
            $fieldRange = $hf.Range; $fieldRange.Collapse(0); $hf.Range.Fields.Add($fieldRange, 33) | Out-Null
        }
        $section.Headers.Item(2).Range.Text = 'First-page {{value:story-value}}'
        $section.Footers.Item(2).Range.Text = 'First-footer {{value:story-value}}'
        $section.Headers.Item(3).Range.Text = 'Even-page {{value:story-value}}'
        $section.Footers.Item(3).Range.Text = 'Even-footer {{value:story-value}}'
    }

    $shape = $doc.Shapes.AddTextbox(1, 72, 120, 240, 60, $doc.Range(0,0))
    $shape.TextFrame.TextRange.Text = 'Text box {{value:story-value}}'
    $shape.AlternativeText = 'Synthetic fixture text box'
    $doc.SaveAs2($path, 16)
    $doc.Close($false); $doc = $null
    $wordVersion = (Get-Item (Join-Path $word.Path 'WINWORD.EXE')).VersionInfo.FileVersion
    [pscustomobject]@{ path = 'tests/fixtures/m0010/word-origin.docx'; application = 'Microsoft Word'; version = $wordVersion; platform = [System.Environment]::OSVersion.VersionString; method = 'Created as a new document through Word COM document model and saved by Word as DOCX'; date = (Get-Date).ToString('o'); sha256 = (Get-FileHash $path -Algorithm SHA256).Hash; synthetic = $true; nonConfidential = $true } | ConvertTo-Json | Set-Content (Join-Path $fixtureDir 'word-origin.provenance.json')
}
finally {
    if ($doc) { $doc.Close($false) }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
