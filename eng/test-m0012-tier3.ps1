param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0012'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$word = Get-Command winword.exe -ErrorAction SilentlyContinue
$lo = Get-Command soffice.exe -ErrorAction SilentlyContinue
if ($null -eq $word -and (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')) { $word = Get-Item 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
if ($null -eq $lo -and (Test-Path 'C:\Program Files\LibreOffice\program\soffice.exe')) { $lo = Get-Item 'C:\Program Files\LibreOffice\program\soffice.exe' }
if ($null -eq $word) { throw 'M0012 Tier 3 blocked: Microsoft Word is unavailable in this interactive user session.' }
if ($null -eq $lo) { throw 'M0012 Tier 3 blocked: LibreOffice Writer is unavailable in this environment.' }
if ([Environment]::Is64BitProcess -eq $false) { throw 'M0012 Tier 3 requires an interactive Windows x64 session.' }
$sdk = (& dotnet --list-sdks) -join "`n"
if ($sdk -notmatch '10\.0\.') { throw 'M0012 Tier 3 requires the .NET 10 SDK.' }

$fixtureRoot = Join-Path $repo 'tests/fixtures/m0012'
$provenancePath = Join-Path $fixtureRoot 'provenance.json'
if (-not (Test-Path -LiteralPath $provenancePath)) { throw 'M0012 Tier 3 blocked: committed application-produced authoring fixture provenance is absent; do not substitute synthetic OOXML or the M0010 fixture matrix.' }
$provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json
foreach ($fixture in $provenance.fixtures) {
    $path = Join-Path $repo ([string]$fixture.path)
    if (-not (Test-Path -LiteralPath $path)) { throw "M0012 Tier 3 fixture is missing: $path" }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne ([string]$fixture.sha256).ToUpperInvariant()) { throw "M0012 Tier 3 fixture hash mismatch: $path" }
}

$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.exe'
if (-not (Test-Path -LiteralPath $cli)) { throw "Build the Word-enabled CLI before Tier 3: $cli" }
$temp = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0012-tier3-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $temp | Out-Null
try {
    # The fixture manifest identifies the Word-origin UI-name/alias case, a style-numbering
    # relationship case, and a real prototype case. Each is authored then finalized through
    # both installed renderers by the committed fixture runner.
    $runner = Join-Path $PSScriptRoot 'run-m0012-tier3.ps1'
    if (-not (Test-Path -LiteralPath $runner)) { throw 'M0012 Tier 3 blocked: application-produced fixtures are present without their authoring/finalization runner.' }
    & $runner -RepositoryRoot $repo -Cli $cli -EvidenceRoot $EvidenceRoot -TemporaryRoot $temp -Fixtures $provenance.fixtures -WordVersion $word.VersionInfo.FileVersion -LibreOfficeVersion $lo.VersionInfo.FileVersion
    if ($LASTEXITCODE -ne 0) { throw 'M0012 Tier 3 authoring compatibility runner failed.' }
}
finally { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
