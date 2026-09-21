param(
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0010')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path $repo 'tests/fixtures/m0010'
$provenancePath = Join-Path $fixtureRoot 'provenance.json'
if (-not (Test-Path -LiteralPath $provenancePath)) { throw "M0010 fixture provenance is missing: $provenancePath" }
$provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json
foreach ($fixture in $provenance.fixtures) {
    $path = Join-Path $repo ([string]$fixture.path)
    if (-not (Test-Path -LiteralPath $path)) { throw "Committed M0010 fixture is missing: $path. Do not synthesize a fallback template." }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne ([string]$fixture.sha256).ToUpperInvariant()) { throw "M0010 fixture hash mismatch: $path" }
}

$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0010 Tier 3 is externally blocked: Microsoft Word is not installed/available in this interactive session.' }
if ($null -eq $lo) { throw 'M0010 Tier 3 is externally blocked: LibreOffice Writer is not installed/available in this interactive session.' }

$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0-windows/yadg.exe'
if (-not (Test-Path -LiteralPath $cli)) { throw "Build the Word-enabled CLI before Tier 3: $cli" }
$revision = (& git -C $repo rev-parse HEAD).Trim()
$os = (Get-CimInstance Win32_OperatingSystem).Caption
$wordVersion = $word.VersionInfo.FileVersion
$loVersion = $lo.VersionInfo.FileVersion
$loPath = if (Test-Path 'C:\Program Files\LibreOffice\program\soffice.com') { 'C:\Program Files\LibreOffice\program\soffice.com' } elseif ($lo.Source) { $lo.Source } else { $lo.FullName }
$runs = @()
$temporary = Join-Path ([System.IO.Path]::GetTempPath()) ('yadg-m0010-tier3-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $temporary | Out-Null
try {
    foreach ($origin in $provenance.fixtures) {
        foreach ($renderer in @('word','libreoffice')) {
            $workspace = Join-Path $temporary ("$($origin.id)-$renderer")
            New-Item -ItemType Directory -Force -Path (Join-Path $workspace 'YadgTemplates'), (Join-Path $workspace 'YadgPreWords'), (Join-Path $workspace 'YadgWords') | Out-Null
            Copy-Item -LiteralPath (Join-Path $repo ([string]$origin.path)) -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
            Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Value @("---", "yadg:", "  version: 1", "values:", "  document-version: 'M0010'", "  story-value: 'Realistic story value'", "---")
            Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Value @('# Architecture {#architecture}', '', 'First real paragraph.', '', 'Second paragraph with **strong** and *emphasis*.', '', '## Assumptions', '', 'Details.')
            & $cli check --workspace $workspace
            if ($LASTEXITCODE -ne 0) { throw "M0010 check failed for $($origin.id) -> $renderer" }
            & $cli build --workspace $workspace
            if ($LASTEXITCODE -ne 0) { throw "M0010 build failed for $($origin.id) -> $renderer" }
            $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
            if ($null -eq $authored) { throw "M0010 authored DOCX missing for $($origin.id) -> $renderer" }
            $renderArgs = @('render','--workspace',$workspace,'--renderer',$renderer)
            if ($renderer -eq 'libreoffice') { $renderArgs += @('--renderer-path',$loPath) }
            & $cli @renderArgs
            if ($LASTEXITCODE -ne 0) { throw "M0010 renderer failed for $($origin.id) -> $renderer" }
            $finalized = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgWords') -Filter '*.docx' | Select-Object -First 1
            if ($null -eq $finalized) { throw "M0010 finalized DOCX missing for $($origin.id) -> $renderer" }
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::OpenRead($finalized.FullName)
            try {
                $xmlText = ($zip.Entries | Where-Object { $_.FullName -like '*.xml' } | ForEach-Object { $reader = New-Object IO.StreamReader($_.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }) -join "`n"
            }
            finally { $zip.Dispose() }
            if ($xmlText -notmatch 'Realistic story value') { throw "Expected substituted story value is absent for $($origin.id) -> $renderer" }
            if ($xmlText -match '\{\{value:') { throw "An expected value tag remains unresolved for $($origin.id) -> $renderer" }
            if ($xmlText -notmatch 'w:pStyle[^>]*w:val="Heading5"') { throw "Expected rebased Heading5 paragraph style is absent for $($origin.id) -> $renderer" }
            if ($xmlText -notmatch '(?i)template') { throw "Static template content is absent for $($origin.id) -> $renderer" }
            $evidenceName = "$($origin.id)-$renderer.docx"
            New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
            Copy-Item -LiteralPath $finalized.FullName -Destination (Join-Path $EvidenceRoot $evidenceName) -Force
            $runs += [ordered]@{ origin = $origin.id; originSha256 = $origin.sha256; renderer = $renderer; authoredSha256 = (Get-FileHash $authored.FullName -Algorithm SHA256).Hash; finalizedSha256 = (Get-FileHash $finalized.FullName -Algorithm SHA256).Hash; result = 'passed' }
        }
    }
    [ordered]@{ repositoryRevision = $revision; os = $os; wordVersion = $wordVersion; libreOfficeVersion = $loVersion; runs = $runs } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier3-evidence.json')
    Write-Output 'M0010 Tier 3 four-path matrix passed.'
}
finally { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
