param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0019'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$root = [IO.Path]::GetFullPath($EvidenceRoot)
function Hash-File([string] $path) {
    $full = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "M0019 manifest input is missing: $full" }
    return (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash.ToUpperInvariant()
}
function Relative-File([string] $path) { return [IO.Path]::GetRelativePath($repo,[IO.Path]::GetFullPath($path)).Replace('\','/') }

$wordCandidate = Join-Path $root 'finalized-word.docx'
$deckCandidate = Join-Path $root 'deck-impress.pptx'
$receipts = [ordered]@{}
foreach ($spec in @(
    @{ name='VAL03'; file='mermaid-tier3-evidence.json' },
    @{ name='VAL04'; file='docx-tier3-evidence.json' },
    @{ name='VAL05'; file='pptx-tier3-evidence.json' }
)) {
    $path = Join-Path $root $spec.file
    $record = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -DateKind String
    if ([string]$record.result -cne 'passed') { throw "$($spec.name) has not passed; refusing to create a review manifest." }
    $receipts[$spec.name] = [ordered]@{ path=Relative-File $path; sha256=Hash-File $path }
}
$val03 = Get-Content -LiteralPath (Join-Path $root 'mermaid-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$val04 = Get-Content -LiteralPath (Join-Path $root 'docx-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$val05 = Get-Content -LiteralPath (Join-Path $root 'pptx-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$val06 = Get-Content -LiteralPath (Join-Path $root 'tier4-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$toolPath = Join-Path $repo ([string]$val06.packagePath)

$manifest = [ordered]@{
    schemaVersion=1; milestone='M0019'; humanReviewId='HR-M0019-01'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim()
    candidates=@(
        [ordered]@{ path='artifacts/review/evidence/M0019/finalized-word.docx'; sha256=Hash-File $wordCandidate; format='DOCX'; originApplication='Microsoft Word'; finalizedBy='Microsoft Word' },
        [ordered]@{ path='artifacts/review/evidence/M0019/deck-impress.pptx'; sha256=Hash-File $deckCandidate; format='PPTX'; originApplication='LibreOffice Impress'; generatedBy='Yadg'; reopenedBy='LibreOffice Impress' }
    )
    inputs=[ordered]@{
        wordSource=[ordered]@{ path='eng/test-m0017-tier3.ps1'; contentSha256=[string]$val04.sourceSha256; sourceDefinitionSha256=Hash-File (Join-Path $repo 'eng/test-m0017-tier3.ps1') }
        wordTemplate=[ordered]@{ path=[string]$val04.wordOriginTemplatePath; sha256=[string]$val04.wordOriginTemplateSha256 }
        presentationSource=[ordered]@{ path=[string]$val05.sourcePath; sha256=[string]$val05.sourceSha256 }
        directConfig=[ordered]@{ path=[string]$val05.configPath; sha256=[string]$val05.configSha256 }
        impressTemplate=[ordered]@{ path=[string]$val05.templatePath; sha256=[string]$val05.templateSha256 }
    }
    tool=[ordered]@{ packageId=[string]$val06.packageId; packageVersion=[string]$val06.packageVersion; packagePath=Relative-File $toolPath; packageSha256=Hash-File $toolPath }
    applications=[ordered]@{ word=[string]$val04.wordVersion; writer=[string]$val04.writerVersion; impress=[string]$val05.impressVersion; powerPoint=[string]$val05.powerPointVersion }
    producer=[ordered]@{ bunVersion=[string]$val03.bunVersion; package=[string]$val03.mermaidPackage; packageVersion=[string]$val03.mermaidVersion }
    tier3Receipts=$receipts
}
$path = Join-Path $root 'review-manifest.json'
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $path -Encoding utf8
Write-Output "M0019 review manifest created: $path"
