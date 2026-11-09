param([string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0019'))

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-tools.json') -Raw | ConvertFrom-Json
$bunVersion = [string]$manifest.bun.version
$cacheRoot = Join-Path ([IO.Path]::GetTempPath()) 'yadg-m0008-tools'
$bun = Join-Path $cacheRoot "bun-$bunVersion-windows-x64/bun.exe"
if (-not (Test-Path -LiteralPath $bun -PathType Leaf)) { & (Join-Path $PSScriptRoot 'test-m0008-tier3.ps1') -CacheRoot $cacheRoot }
if ((& $bun --version).Trim() -ne $bunVersion) { throw "M0019 DOCX Tier 3 requires pinned Bun $bunVersion." }
$globalPath = Join-Path $repo 'global.json'
if (Test-Path -LiteralPath $globalPath) { throw 'Refusing to replace an existing repository global.json.' }
Set-Content -LiteralPath $globalPath -Value '{"sdk":{"version":"10.0.403","rollForward":"disable","allowPrerelease":false}}' -Encoding ascii
try {
    & (Join-Path $PSScriptRoot 'test-m0017-tier3.ps1') -M0019Mode -MermaidExecutable $bun
    if ($LASTEXITCODE -ne 0) { throw 'M0019 real Word/Writer authoring and finalization failed.' }
} finally { Remove-Item -LiteralPath $globalPath -Force -ErrorAction SilentlyContinue }

$sourceEvidencePath = Join-Path $repo 'artifacts/review/evidence/M0017/tier3-evidence.json'
$sourceEvidence = Get-Content -LiteralPath $sourceEvidencePath -Raw | ConvertFrom-Json -DateKind String
$wordRun = @($sourceEvidence.runs | Where-Object { $_.renderer -eq 'word' -and $_.result -eq 'passed' })
$writerRun = @($sourceEvidence.runs | Where-Object { $_.renderer -eq 'libreoffice' -and $_.result -eq 'passed' })
if ($wordRun.Count -ne 1 -or $writerRun.Count -ne 1) { throw 'M0019 DOCX Tier 3 requires successful native Word and Writer runs.' }
New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
$candidate = Join-Path $EvidenceRoot 'finalized-word.docx'
$sourceArtifact = Join-Path $repo ([string]$wordRun[0].finalized)
Copy-Item -LiteralPath $sourceArtifact -Destination $candidate -Force
$sourceEvidenceHash = (Get-FileHash -LiteralPath $sourceEvidencePath -Algorithm SHA256).Hash.ToUpperInvariant()
$record = [ordered]@{
    milestone='M0019'; gate='VAL-04'; result='passed'; repositoryRevision=(& git -C $repo rev-parse HEAD).Trim()
    wordApplication='Microsoft Word'; wordVersion=[string]$sourceEvidence.wordVersion
    writerApplication='LibreOffice Writer'; writerVersion=[string]$sourceEvidence.libreOfficeVersion
    wordOriginTemplatePath=[string]$sourceEvidence.templatePath; wordOriginTemplateSha256=[string]$sourceEvidence.templateSha256
    sourceSha256=[string]$sourceEvidence.sourceSha256; sourceIncludesValues=$true; mermaidProducer="$($manifest.mermaidCli.package)@$($manifest.mermaidCli.version)"
    wordFinalizedPath='artifacts/review/evidence/M0019/finalized-word.docx'; wordFinalizedSha256=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToUpperInvariant()
    writerFinalizedPath=[string]$writerRun[0].finalized; writerFinalizedSha256=[string]$writerRun[0].finalizedSha256
    underlyingM0017Receipt='artifacts/review/evidence/M0017/tier3-evidence.json'; underlyingM0017ReceiptSha256=$sourceEvidenceHash
    resultDetails='Both real runtimes finalized generated DOCX files and checked source values, the Mermaid-produced image, formatting, references and links.'
}
$record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'docx-tier3-evidence.json') -Encoding utf8
Write-Output 'M0019 real Word and Writer Tier 3 passed with source-value substitution and pinned Mermaid producer configuration.'
