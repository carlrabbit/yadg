param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0020'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = [IO.Path]::GetFullPath($EvidenceRoot)
$powerPointExe = Join-Path $env:ProgramFiles 'Microsoft Office/root/Office16/POWERPNT.EXE'
$soffice = Join-Path $env:ProgramFiles 'LibreOffice/program/soffice.exe'
$loPython = Join-Path $env:ProgramFiles 'LibreOffice/program/python.exe'
foreach ($required in @($powerPointExe,$soffice,$loPython)) { if (-not (Test-Path -LiteralPath $required -PathType Leaf)) { throw "M0020 VAL-03 is BLOCKED; required native application/runtime is missing: $required" } }
New-Item -ItemType Directory -Force -Path $root | Out-Null
$source = Join-Path $root 'box-diagram-source.md'
$config = Join-Path $root 'box-diagram-config.yaml'
$sourceText = Get-Content -LiteralPath (Join-Path $repo 'tests/fixtures/m0018/review-source.md') -Raw
$diagram = @'
```diagram {#m0020-box caption="Nested service diagram"}
┌────────────┐
│ Outer      │
│ ┌──────┐   │
│ │Inner │   │
│ └──────┘   │
└────────────┘
  ↓
  ↓
```
'@
$sourceText = [regex]::Replace($sourceText,'!\[[^\]]*\]\([^\)]*\)',$diagram.Trim())
Set-Content -LiteralPath $source -Encoding utf8 -Value $sourceText
Set-Content -LiteralPath $config -Encoding utf8 -Value @('yadg:','  version: 1','diagrams:','  autoDetect: false')
$powerPointFixture = Join-Path $repo 'tests/fixtures/m0018/powerpoint-origin.pptx'
$impressFixture = Join-Path $repo 'tests/fixtures/m0018/impress-origin.pptx'
$powerPointOutput = Join-Path $root 'deck-powerpoint-svg.pptx'
$impressOutput = Join-Path $root 'deck-impress-svg.pptx'
$powerPointReceipt = Join-Path $root 'powerpoint-tier3-evidence.json'
$impressReceipt = Join-Path $root 'impress-tier3-evidence.json'
& (Join-Path $PSScriptRoot 'verify-m0018-fixtures.ps1')
if (-not $?) { throw 'M0020 requires current verified M0018 PowerPoint and Impress origin fixtures.' }
& (Join-Path $PSScriptRoot 'test-m0018-powerpoint-tier3.ps1') -TemplatePath $powerPointFixture -SourcePath $source -OutputPath $powerPointOutput -EvidencePath $powerPointReceipt -ConfigPath $config -Milestone M0020 -Gate VAL-03
if (-not $?) { throw 'PowerPoint-origin -> YADG -> PowerPoint failed.' }
& (Join-Path $PSScriptRoot 'test-m0018-tier3.ps1') -TemplatePath $impressFixture -SourcePath $source -CandidatePath $impressOutput -EvidencePath $impressReceipt -ConfigPath $config -AutomatedReviewOnly -Milestone M0020 -Gate VAL-03
if (-not $?) { throw 'Impress-origin -> YADG -> Impress failed.' }

function Confirm-SvgPicture([string] $path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $svg = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/media/.*\.svg$' })
        if ($svg.Count -lt 1) { throw "Candidate is missing embedded SVG payload: $path" }
        $svgText = $svg | ForEach-Object { $reader = [IO.StreamReader]::new($_.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }
        $allSvg = $svgText -join "`n"
        foreach ($label in @('Outer','Inner')) { if ($allSvg -notmatch [regex]::Escape($label)) { throw "SVG picture omits diagram label '$label': $path" } }
        $slides = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' })
        $pictures = 0; $embeddedSvgRelationships = 0
        foreach ($slideEntry in $slides) {
            $reader = [IO.StreamReader]::new($slideEntry.Open()); try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $pictures += ([regex]::Matches($xml,'<p:pic(?:\s|>)')).Count
            $embeddedSvgRelationships += ([regex]::Matches($xml,'asvg:svgBlip')).Count
        }
        if ($embeddedSvgRelationships -lt 1 -or $pictures -lt 1) { throw "Candidate does not reference its SVG through an ordinary picture shape: $path" }
        return [ordered]@{ svgPartCount=$svg.Count; pictureCount=$pictures; svgPictureReferences=$embeddedSvgRelationships; labelSha256=(Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($allSvg))) -Algorithm SHA256).Hash.ToUpperInvariant() }
    } finally { $zip.Dispose() }
}
$powerPoint = Get-Content -LiteralPath $powerPointReceipt -Raw | ConvertFrom-Json -DateKind String
$impress = Get-Content -LiteralPath $impressReceipt -Raw | ConvertFrom-Json -DateKind String
if ($powerPoint.result -cne 'passed' -or $impress.result -cne 'passed') { throw 'An origin-specific native application receipt did not pass.' }
$powerPointSvg = Confirm-SvgPicture $powerPointOutput
$impressSvg = Confirm-SvgPicture $impressOutput
$revision = (& git -C $repo rev-parse HEAD).Trim()
$record = [ordered]@{
    milestone='M0020'; gate='VAL-03'; result='passed'; repositoryRevision=$revision
    sourcePath=[IO.Path]::GetRelativePath($repo,$source).Replace('\','/'); sourceSha256=(Get-FileHash $source -Algorithm SHA256).Hash.ToUpperInvariant()
    configPath=[IO.Path]::GetRelativePath($repo,$config).Replace('\','/'); configSha256=(Get-FileHash $config -Algorithm SHA256).Hash.ToUpperInvariant()
    paths=@(
        [ordered]@{origin='Microsoft PowerPoint'; fixture='tests/fixtures/m0018/powerpoint-origin.pptx'; fixtureSha256=(Get-FileHash $powerPointFixture -Algorithm SHA256).Hash.ToUpperInvariant(); candidate='artifacts/review/evidence/M0020/deck-powerpoint-svg.pptx'; candidateSha256=(Get-FileHash $powerPointOutput -Algorithm SHA256).Hash.ToUpperInvariant(); applicationVersion=$powerPoint.applicationVersion; receipt='artifacts/review/evidence/M0020/powerpoint-tier3-evidence.json'; receiptSha256=(Get-FileHash $powerPointReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); svg=$powerPointSvg; nativeRenderings=$powerPoint.inspection.nativeRenderings }
        [ordered]@{origin='LibreOffice Impress'; fixture='tests/fixtures/m0018/impress-origin.pptx'; fixtureSha256=(Get-FileHash $impressFixture -Algorithm SHA256).Hash.ToUpperInvariant(); candidate='artifacts/review/evidence/M0020/deck-impress-svg.pptx'; candidateSha256=(Get-FileHash $impressOutput -Algorithm SHA256).Hash.ToUpperInvariant(); applicationVersion=$impress.applicationVersion; receipt='artifacts/review/evidence/M0020/impress-tier3-evidence.json'; receiptSha256=(Get-FileHash $impressReceipt -Algorithm SHA256).Hash.ToUpperInvariant(); svg=$impressSvg; nativeRenderings=$impress.inspection.nativeRenderings }
    )
    humanReviewId='HR-M0020-01'; humanReviewStatus='pending'
}
$record | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath (Join-Path $root 'pptx-tier3-evidence.json') -Encoding utf8
Write-Output 'M0020 PowerPoint-origin/PowerPoint and Impress-origin/Impress validations passed independently; HR-M0020-01 remains pending.'
