$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path $repo 'tests/fixtures/m0018'
$expected = @('powerpoint-origin', 'impress-origin')
$actual = @(Get-ChildItem -LiteralPath $fixtureRoot -Filter '*-origin.pptx' -File | ForEach-Object { $_.BaseName })
$actualNames = (@($actual | Sort-Object) -join '|')
$expectedNames = (@($expected | Sort-Object) -join '|')
if ($actualNames -cne $expectedNames) {
    throw "M0018 requires exactly two canonical application-origin fixtures; found: $($actual -join ', ')."
}
$layouts = @('content', 'two-columns', 'text-image', 'table')
foreach ($id in $expected) {
    $path = Join-Path $fixtureRoot "$id.pptx"
    $provenancePath = Join-Path $fixtureRoot "$id.provenance.json"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or -not (Test-Path -LiteralPath $provenancePath -PathType Leaf)) { throw "M0018 fixture or provenance is missing: $id" }
    $record = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json -DateKind String
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()
    $application = if ($id -eq 'powerpoint-origin') { 'Microsoft PowerPoint' } else { 'LibreOffice Impress' }
    if ([string]$record.id -cne $id -or [string]$record.originApplication -cne $application) { throw "M0018 fixture provenance has the wrong identity/application: $id" }
    if ($hash -cne [string]$record.sha256) { throw "M0018 fixture SHA-256 does not match committed bytes: $id" }
    if (-not [bool]$record.createdInApplication -or -not [bool]$record.notOpenXmlGeneratedOrRoundTripped -or -not [bool]$record.retainedNativeMasterThemeLayouts -or -not [bool]$record.syntheticAndRedistributionSafe) { throw "M0018 fixture provenance does not establish native application origin: $id" }
    if ([string]::IsNullOrWhiteSpace([string]$record.applicationVersion) -or [string]::IsNullOrWhiteSpace([string]$record.platform) -or [string]::IsNullOrWhiteSpace([string]$record.creationMethod)) { throw "M0018 fixture provenance is incomplete: $id" }
    foreach ($timestamp in @('createdAtUtc', 'saveDateUtc')) {
        $date = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse([string]$record.$timestamp, [ref]$date) -or $date.Offset -ne [TimeSpan]::Zero) { throw "M0018 fixture $timestamp must be UTC: $id" }
    }
    if ((@($record.layouts) -join '|') -cne ($layouts -join '|')) { throw "M0018 fixture does not identify the required four layouts: $id" }
    $archive = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $slides = @($archive.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object FullName)
        if ($slides.Count -ne 4) { throw "M0018 fixture must contain exactly four prototype slides: $id" }
        $texts = foreach ($entry in $slides) { $reader = [IO.StreamReader]::new($entry.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }
        foreach ($layout in $layouts) { if (-not ($texts -match [regex]::Escape("{{slide:$layout}}"))) { throw "M0018 fixture is missing prototype marker {{slide:$layout}}: $id" } }
        if (-not ($texts -match '\{\{table\}\}')) { throw "M0018 fixture is missing its editable table marker: $id" }
    } finally { $archive.Dispose() }
    Write-Output "$id fixture verified: $hash"
}
