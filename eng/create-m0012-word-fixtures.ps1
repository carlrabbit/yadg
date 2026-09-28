$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$fixtureDir = Join-Path $repo 'tests/fixtures/m0012'
New-Item -ItemType Directory -Force -Path $fixtureDir | Out-Null
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$created = [System.Collections.Generic.List[object]]::new()
try {
    foreach ($scenario in @('style-resolution','list-prototypes')) {
        $path = Join-Path $fixtureDir "$scenario.docx"
        $doc = $word.Documents.Add()
        try {
            $code = $doc.Styles.Add('Code Character', 2)
            $code.Font.Name = 'Consolas'
            $code.Font.Color = 0x7030A0
            if ($scenario -eq 'style-resolution') {
                $custom = $doc.Styles.Add('User Bullets', 1)
                $custom.BaseStyle = $doc.Styles.Item('List Bullet')
                $custom.Font.Color = 0x008000
                $lines = @('{{yadg:frontmatter}}','version: 1','styles:','  headings:','    1: Heading 1','  lists:','    unordered: User Bullets','    ordered: List Number','  codeInline: Code Character','{{/yadg:frontmatter}}','Word-origin style discovery sample.','{{content:introduction}}','Template-owned list style sample.','Word list relationship sample.','Word ordered list relationship sample.')
            } else {
                $lines = @('{{yadg:frontmatter}}','version: 1','styles:','  headings:','    1: Heading 1','  codeInline: Code Character','prototypes:','  unorderedListItem: bullet-item','  orderedListItem: number-item','{{/yadg:frontmatter}}','{{yadg:prototype:bullet-item}}','{{item}}','{{/yadg:prototype:bullet-item}}','{{yadg:prototype:number-item}}','{{item}}','{{/yadg:prototype:number-item}}','Word-origin list prototypes.','{{content:introduction}}')
            }
            foreach ($line in $lines) {
                $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
                $range.InsertAfter($line + [char]13)
            }
            foreach ($paragraph in $doc.Paragraphs) {
                $text = $paragraph.Range.Text.Trim()
                if ($text -eq 'Word-origin style discovery sample.' -or $text -eq 'Word-origin list prototypes.') { $paragraph.Style = $doc.Styles.Item('Heading 1') }
                if ($scenario -eq 'style-resolution' -and $text -eq 'Word list relationship sample.') { $paragraph.Style = $doc.Styles.Item('User Bullets'); $paragraph.Range.ListFormat.ApplyBulletDefault() }
                if ($scenario -eq 'style-resolution' -and $text -eq 'Word ordered list relationship sample.') { $paragraph.Style = $doc.Styles.Item('List Number'); $paragraph.Range.ListFormat.ApplyNumberDefault() }
                if ($scenario -eq 'list-prototypes' -and $text -eq '{{item}}') {
                    $previous = $paragraph.Previous()
                    if ($previous.Range.Text.Trim() -eq '{{yadg:prototype:bullet-item}}') { $paragraph.Range.ListFormat.ApplyBulletDefault() }
                    else { $paragraph.Range.ListFormat.ApplyNumberDefault() }
                }
            }
            $doc.SaveAs2($path, 16)
            $doc.Close($false)
        } catch { if ($doc) { $doc.Close($false) }; throw }
        $wordVersion = (Get-Item (Join-Path $word.Path 'WINWORD.EXE')).VersionInfo.FileVersion
        $created.Add([ordered]@{ id = $scenario; path = "tests/fixtures/m0012/$scenario.docx"; application = 'Microsoft Word'; version = $wordVersion; platform = [Environment]::OSVersion.VersionString; method = 'Created as a new blank document and authored through the Microsoft Word COM document model, then saved as DOCX by Word'; date = (Get-Date).ToString('o'); sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; synthetic = $true; nonConfidential = $true })
    }
    [ordered]@{ fixtures = $created } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $fixtureDir 'provenance.json')
    Write-Output "Created $($created.Count) new-document Word fixtures and recorded hashes."
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
