param(
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0014')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0014 Tier 3 is externally blocked: Microsoft Word is not installed/available in this interactive session.' }
if ($null -eq $lo) { throw 'M0014 Tier 3 is externally blocked: LibreOffice Writer is not installed/available in this interactive session.' }

$fixture = Join-Path $repo 'tests/fixtures/m0010/word-origin.docx'
if (-not (Test-Path -LiteralPath $fixture)) { throw "Required Word-origin template fixture is missing: $fixture" }
$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.exe'
if (-not (Test-Path -LiteralPath $cli)) { throw "Build the Word-enabled CLI before Tier 3: $cli" }
$revision = (& git -C $repo rev-parse HEAD).Trim()
$worktree = (& git -C $repo status --short --untracked-files=all) -join "`n"
$worktreeHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($worktree)))
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$wordVersion = $word.VersionInfo.FileVersion
$loVersion = $lo.VersionInfo.FileVersion
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } else { $lo.Source }
$source = @'
# Architecture {#architecture}

Representative source for the M0014 text-diagram renderer.

```diagram {#explicit-flow}
┌────────────┐
│ Outer      │
│ ┌──────┐   │
│ │Inner │   │
│ └───┴──┘   │
└─────┼──────┘
      │
      ▼
```

```text
┌────────────┐
│ Outer      │
│ ┌──────┐   │
│ │Inner │   │
│ └──────┘   │
└────────────┘
      ↓
```

```text {#auto-flow}
      ┌───────┐
      │ upper │
┌─────│       │─────┐
│lower│       │     │
│     └───────┘     │
│                   │
│                   │
└───────────────────┘
```
'@

$temporary = Join-Path ([System.IO.Path]::GetTempPath()) ('yadg-m0014-tier3-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $temporary, $EvidenceRoot | Out-Null
$runs = @()
try {
    foreach ($renderer in @('word', 'libreoffice')) {
        $workspace = Join-Path $temporary $renderer
        New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates') | Out-Null
        Copy-Item -LiteralPath $fixture -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
        Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Encoding utf8 -Value @('---','yadg:','  version: 1','values:','  story-value: "M0014 Tier 3"','diagrams:','  autoDetect: true','---')
        Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Encoding utf8 -Value $source

        & $cli check --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0014 check failed for $renderer." }
        & $cli build --workspace $workspace
        if ($LASTEXITCODE -ne 0) { throw "M0014 build failed for $renderer." }
        $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
        if ($null -eq $authored) { throw "M0014 authored DOCX is missing for $renderer." }

        $renderArgs = @('render','--workspace',$workspace,'--renderer',$renderer)
        if ($renderer -eq 'libreoffice') { $renderArgs += @('--renderer-path',$loPath) }
        & $cli @renderArgs
        if ($LASTEXITCODE -ne 0) { throw "M0014 $renderer finalization failed." }
        $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
        if ($null -eq $finalized) { throw "M0014 finalized DOCX is missing for $renderer." }

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($finalized.FullName)
        try {
            $documentEntry = $zip.GetEntry('word/document.xml')
            if ($null -eq $documentEntry) { throw "Finalized $renderer DOCX has no main document part." }
            $documentReader = [IO.StreamReader]::new($documentEntry.Open())
            try { $documentXml = $documentReader.ReadToEnd() } finally { $documentReader.Dispose() }
            $svgEntries = @($zip.Entries | Where-Object { $_.FullName -match '\.svg$' })
            if ($svgEntries.Count -lt 3) { throw "Finalized $renderer DOCX retained only $($svgEntries.Count) SVG diagram assets; expected explicit, anonymous, and identified auto-detected diagrams." }
            if (($documentXml | Select-String -AllMatches -Pattern 'svgBlip').Matches.Count -lt 3) { throw "Finalized $renderer DOCX does not reference all generated SVG diagrams." }
            $allSvg = foreach ($entry in $svgEntries) {
                $reader = [IO.StreamReader]::new($entry.Open())
                try { $reader.ReadToEnd() } finally { $reader.Dispose() }
            }
            $svgText = $allSvg -join "`n"
            foreach ($expected in @('Outer','Inner','upper','lower')) { if ($svgText -notmatch [regex]::Escape($expected)) { throw "Expected diagram text '$expected' is absent from finalized $renderer SVG content." } }
            if ($svgText -notmatch '<path') { throw "Finalized $renderer SVG content is missing the standalone arrow primitive." }
            if ($svgText -match '<script|foreignObject|(?:href|src)\s*=\s*["''](?:https?:|//|data:)') { throw "Finalized $renderer SVG contains active or external content." }
        }
        finally { $zip.Dispose() }

        $name = if ($renderer -eq 'word') { 'word-diagram.docx' } else { 'libreoffice-diagram.docx' }
        Copy-Item -LiteralPath $finalized.FullName -Destination (Join-Path $EvidenceRoot $name) -Force
        $runs += [ordered]@{ renderer = $renderer; runtimeVersion = if ($renderer -eq 'word') { $wordVersion } else { $loVersion }; template = 'tests/fixtures/m0010/word-origin.docx'; authoredSha256 = (Get-FileHash $authored.FullName -Algorithm SHA256).Hash; finalized = "artifacts/review/evidence/M0014/$name"; finalizedSha256 = (Get-FileHash -LiteralPath (Join-Path $EvidenceRoot $name) -Algorithm SHA256).Hash; svgPartCount = $svgEntries.Count; result = 'passed' }
    }
    $negativeWorkspace = Join-Path $temporary 'negative-text-fence'
    New-Item -ItemType Directory -Force -Path (Join-Path $negativeWorkspace 'YadgTemplates') | Out-Null
    Copy-Item -LiteralPath $fixture -Destination (Join-Path $negativeWorkspace 'YadgTemplates/template.docx')
    Set-Content -LiteralPath (Join-Path $negativeWorkspace 'YADG.md') -Encoding utf8 -Value @('---','yadg:','  version: 1','diagrams:','  autoDetect: true','---')
    Set-Content -LiteralPath (Join-Path $negativeWorkspace 'content.md') -Encoding utf8 -Value @('```text','ordinary label: value','```')
    $negative = & $cli check --workspace $negativeWorkspace 2>&1
    if ($LASTEXITCODE -eq 0 -or (($negative -join "`n") -notmatch 'YADG-MD-UNSUPPORTED')) { throw 'Weak ordinary text fence was converted instead of retaining the unsupported-fence check diagnostic.' }
    $negativeCase = [ordered]@{ case = 'ordinary-text-fence'; command = 'check'; expectedDiagnostic = 'YADG-MD-UNSUPPORTED'; result = 'passed' }
    $evidence = [ordered]@{ milestone = 'M0014'; repositoryRevision = $revision; worktreeStatusSha256 = $worktreeHash; os = $os; wordVersion = $wordVersion; libreOfficeVersion = $loVersion; sourceSha256 = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($source))); negativeFenceCheck = $negativeCase; runs = $runs }
    $evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json') -Encoding utf8
    Write-Output 'M0014 Tier 3 Word and LibreOffice runs passed.'
}
finally { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
