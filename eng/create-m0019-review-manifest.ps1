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
function Read-Receipt([string] $name) {
    $path = Join-Path $root $name
    $record = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -DateKind String
    if ([string]$record.result -cne 'passed') { throw "M0019 receipt did not pass: $name" }
    return [ordered]@{ path=Relative-File $path; sha256=Hash-File $path }
}
function Rendering-Set($record) {
    $items = @()
    foreach ($key in @('prototype','candidate')) {
        $item = $record.inspection.nativeRenderings.$key
        if ($item -is [System.Array]) { continue }
        $items += [ordered]@{ kind=$key; path=[string]$item.path; sha256=[string]$item.sha256; pageCount=[int]$item.pageCount; renderedBy=[string]$item.renderedBy }
    }
    foreach ($key in @('prototypeSlideImages','candidateSlideImages')) {
        foreach ($item in $record.inspection.nativeRenderings.$key) {
            $items += [ordered]@{ kind=$key; slide=[int]$item.slide; path=[string]$item.path; sha256=[string]$item.sha256; renderedBy=[string]$item.renderedBy }
        }
    }
    foreach ($kind in @('prototype','candidate')) {
        $slides = $record.inspection.nativeRenderings.$kind
        if ($slides -is [System.Array]) {
            foreach ($item in $slides) {
                $items += [ordered]@{ kind="$kind-slide"; slide=[int]$item.slide; path=[string]$item.path; sha256=[string]$item.sha256; renderedBy='Microsoft PowerPoint native slide export' }
            }
        }
    }
    return ,$items
}

$wordCandidate = Join-Path $root 'finalized-word.docx'
$powerPointCandidate = Join-Path $root 'deck-powerpoint.pptx'
$impressCandidate = Join-Path $root 'deck-impress.pptx'
$mermaid = Get-Content -LiteralPath (Join-Path $root 'mermaid-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$word = Get-Content -LiteralPath (Join-Path $root 'docx-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$powerPoint = Get-Content -LiteralPath (Join-Path $root 'powerpoint-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$impress = Get-Content -LiteralPath (Join-Path $root 'impress-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$val05 = Get-Content -LiteralPath (Join-Path $root 'pptx-tier3-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$val06 = Get-Content -LiteralPath (Join-Path $root 'tier4-evidence.json') -Raw | ConvertFrom-Json -DateKind String
$toolPath = Join-Path $repo ([string]$val06.packagePath)

$receipts = [ordered]@{
    VAL03=(Read-Receipt 'mermaid-tier3-evidence.json')
    VAL04=(Read-Receipt 'docx-tier3-evidence.json')
    VAL05=(Read-Receipt 'pptx-tier3-evidence.json')
    VAL05PowerPoint=(Read-Receipt 'powerpoint-tier3-evidence.json')
    VAL05Impress=(Read-Receipt 'impress-tier3-evidence.json')
}

$manifest = [ordered]@{
    schemaVersion=2; milestone='M0019'; humanReviewId='HR-M0019-01'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim()
    candidateCount=3
    candidates=@(
        [ordered]@{ path='artifacts/review/evidence/M0019/finalized-word.docx'; sha256=Hash-File $wordCandidate; format='DOCX'; originApplication='Microsoft Word'; finalizedBy='Microsoft Word'; nativeTier3Receipt=$receipts.VAL04 }
        [ordered]@{ path='artifacts/review/evidence/M0019/deck-powerpoint.pptx'; sha256=Hash-File $powerPointCandidate; format='PPTX'; originApplication='Microsoft PowerPoint'; generatedBy='Yadg'; reopenedBy='Microsoft PowerPoint'; nativeTier3Receipt=$receipts.VAL05PowerPoint; nativeRenderings=(Rendering-Set $powerPoint) }
        [ordered]@{ path='artifacts/review/evidence/M0019/deck-impress.pptx'; sha256=Hash-File $impressCandidate; format='PPTX'; originApplication='LibreOffice Impress'; generatedBy='Yadg'; reopenedBy='LibreOffice Impress'; nativeTier3Receipt=$receipts.VAL05Impress; nativeRenderings=(Rendering-Set $impress) }
    )
    inputs=[ordered]@{
        wordSource=[ordered]@{ path='eng/test-m0017-tier3.ps1'; contentSha256=[string]$word.sourceSha256; sourceDefinitionSha256=Hash-File (Join-Path $repo 'eng/test-m0017-tier3.ps1') }
        wordTemplate=[ordered]@{ path=[string]$word.wordOriginTemplatePath; sha256=[string]$word.wordOriginTemplateSha256 }
        presentationSource=[ordered]@{ path=[string]$powerPoint.sourcePath; sha256=[string]$powerPoint.sourceSha256 }
        directConfig=[ordered]@{ path=[string]$powerPoint.configPath; sha256=[string]$powerPoint.configSha256 }
        powerpointTemplate=[ordered]@{ path=[string]$powerPoint.templatePath; sha256=[string]$powerPoint.templateSha256; provenancePath='tests/fixtures/m0018/powerpoint-origin.provenance.json'; provenanceSha256=Hash-File (Join-Path $repo 'tests/fixtures/m0018/powerpoint-origin.provenance.json') }
        impressTemplate=[ordered]@{ path=[string]$impress.templatePath; sha256=[string]$impress.templateSha256; provenancePath='tests/fixtures/m0018/impress-origin.provenance.json'; provenanceSha256=Hash-File (Join-Path $repo 'tests/fixtures/m0018/impress-origin.provenance.json') }
    }
    tool=[ordered]@{ packageId=[string]$val06.packageId; packageVersion=[string]$val06.packageVersion; packagePath=Relative-File $toolPath; packageSha256=Hash-File $toolPath }
    applications=[ordered]@{ word=[string]$word.wordVersion; writer=[string]$word.writerVersion; impress=[string]$impress.applicationVersion; powerPoint=[string]$powerPoint.applicationVersion }
    producer=[ordered]@{ bunVersion=[string]$mermaid.bunVersion; package=[string]$mermaid.mermaidPackage; packageVersion=[string]$mermaid.mermaidVersion }
    tier3Receipts=$receipts
    visualReview=[ordered]@{ status='executor-renderings-reviewed; human-approval-pending'; summaryReceipt='artifacts/review/evidence/M0019/pptx-tier3-evidence.json'; nativePathRenderings=@('PowerPoint prototype and candidate PNG slides','Impress prototype and candidate PNG slides','Impress native PDF exports'); visualFitAcceptedByAutomation=$false }
}
$path = Join-Path $root 'review-manifest.json'
$manifest | ConvertTo-Json -Depth 14 | Set-Content -LiteralPath $path -Encoding utf8
Write-Output "M0019 review manifest created for three candidates: $path"
