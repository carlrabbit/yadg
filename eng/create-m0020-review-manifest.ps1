param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0020'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = [IO.Path]::GetFullPath($EvidenceRoot)
function Hash([string] $path) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing M0020 review input: $path" }; (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() }
function Rel([string] $path) { [IO.Path]::GetRelativePath($repo,[IO.Path]::GetFullPath($path)).Replace('\','/') }
function RenderingList($source) {
    $items = @()
    if ($null -eq $source) { return ,$items }
    foreach ($kind in @('prototype','candidate')) {
        $value = $source.$kind
        if ($value -is [System.Array]) { foreach ($item in $value) { $items += $item } }
    }
    foreach ($kind in @('prototypeSlideImages','candidateSlideImages')) { foreach ($item in $source.$kind) { $items += $item } }
    return ,$items
}
$tier3Path = Join-Path $root 'pptx-tier3-evidence.json'
$tier4Path = Join-Path $root 'tier4-evidence.json'
$tier3 = Get-Content -LiteralPath $tier3Path -Raw | ConvertFrom-Json -DateKind String
$tier4 = Get-Content -LiteralPath $tier4Path -Raw | ConvertFrom-Json -DateKind String
if ($tier3.result -cne 'passed' -or $tier3.gate -cne 'VAL-03' -or $tier4.result -cne 'passed' -or $tier4.gate -cne 'VAL-04') { throw 'M0020 review manifest requires successful current VAL-03 and VAL-04 receipts.' }
$candidates = @()
foreach ($path in $tier3.paths) {
    $candidatePath = Join-Path $repo $path.candidate
    $fixturePath = Join-Path $repo $path.fixture
    $renderings = RenderingList $path.nativeRenderings
    $candidates += [ordered]@{
        path=$path.candidate; sha256=Hash $candidatePath; format='PPTX'; originApplication=$path.origin; nativeApplication=$path.origin
        fixture=[ordered]@{path=$path.fixture;sha256=Hash $fixturePath;provenancePath=($path.fixture -replace '\.pptx$','.provenance.json');provenanceSha256=Hash (Join-Path $repo ($path.fixture -replace '\.pptx$','.provenance.json'))}
        source=[ordered]@{path=$tier3.sourcePath;sha256=$tier3.sourceSha256}; config=[ordered]@{path=$tier3.configPath;sha256=$tier3.configSha256}
        tier3Receipt=[ordered]@{path=$path.receipt;sha256=Hash (Join-Path $repo $path.receipt)}
        tier4Receipt=[ordered]@{path='artifacts/review/evidence/M0020/tier4-evidence.json';sha256=Hash $tier4Path}
        installedCandidate=($tier4.nativeReceipts | Where-Object { $_.candidate -like "*$(if ($path.origin -eq 'Microsoft PowerPoint') {'powerpoint'} else {'impress'})*" } | Select-Object -First 1)
        applicationVersion=$path.applicationVersion; svg=$path.svg; nativeRenderings=$renderings
    }
}
if ($candidates.Count -ne 2) { throw 'M0020 review requires exactly two distinct PPTX candidates.' }
$packagePath = Join-Path $repo $tier4.packagePath
$manifest = [ordered]@{
    schemaVersion=1; milestone='M0020'; humanReviewId='HR-M0020-01'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim(); candidateCount=2
    candidates=$candidates
    tier3Receipt=[ordered]@{path='artifacts/review/evidence/M0020/pptx-tier3-evidence.json';sha256=Hash $tier3Path}
    tier4Receipt=[ordered]@{path='artifacts/review/evidence/M0020/tier4-evidence.json';sha256=Hash $tier4Path}
    package=[ordered]@{id=$tier4.packageId;version=$tier4.packageVersion;path=$tier4.packagePath;sha256=Hash $packagePath}
    humanReviewStatus='pending'
}
$manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $root 'review-manifest.json') -Encoding utf8
Write-Output 'M0020 human-review manifest created for the two native-origin SVG deck candidates; approval remains pending.'
