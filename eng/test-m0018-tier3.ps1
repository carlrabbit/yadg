param(
    [string] $TemplatePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/impress-origin.pptx'),
    [string] $SourcePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/review-source.md'),
    [string] $CandidatePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/review-sample.pptx'),
    [string] $EvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/impress-tier3-evidence.json')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$template = [IO.Path]::GetFullPath($TemplatePath)
$source = [IO.Path]::GetFullPath($SourcePath)
$candidate = [IO.Path]::GetFullPath($CandidatePath)
$evidence = [IO.Path]::GetFullPath($EvidencePath)
$provenancePath = Join-Path (Split-Path -Parent $template) 'impress-origin.provenance.json'
$provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json -DateKind String
$soffice = Join-Path $env:ProgramFiles 'LibreOffice/program/soffice.exe'
$python = Join-Path $env:ProgramFiles 'LibreOffice/program/python.exe'
$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.dll'
if (-not (Test-Path -LiteralPath $soffice -PathType Leaf) -or -not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'M0018 Tier 3 is blocked: real LibreOffice Impress and its UNO Python runtime are required.' }
if (-not (Test-Path -LiteralPath $cli -PathType Leaf)) { throw "Build the Release CLI before Tier 3: $cli" }
if ((Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant() -cne [string]$provenance.sha256) { throw 'Impress-origin fixture bytes do not match their provenance hash.' }
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Presentation source is missing: $source" }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $candidate),(Split-Path -Parent $evidence) | Out-Null

$revision = (& git -C $repo rev-parse HEAD).Trim()
$templateHash = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant()
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToUpperInvariant()
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo($soffice).ProductVersion
$profile = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0018-impress-' + [guid]::NewGuid().ToString('N'))
$profileUri = ([uri]([IO.Path]::GetFullPath($profile) + [IO.Path]::DirectorySeparatorChar)).AbsoluteUri.TrimEnd('/')
$openCopy = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0018-impress-output-' + [guid]::NewGuid().ToString('N') + '.pptx')
$port = Get-Random -Minimum 22000 -Maximum 42000
$profileCleaned = $false
$failure = $null
$result = 'failed'
$process = $null
$inspectScript = Join-Path $PSScriptRoot 'inspect-m0018-impress-tier3.py'
try {
    & dotnet $cli presentation check --template $template --source $source
    if ($LASTEXITCODE -ne 0) { throw 'Impress-origin fixture failed presentation preflight.' }
    & dotnet $cli presentation build --template $template --source $source --output $candidate
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw 'Impress-origin review candidate build failed.' }
    Copy-Item -LiteralPath $candidate -Destination $openCopy

    $process = Start-Process -FilePath $soffice -ArgumentList @("-env:UserInstallation=`"$profileUri`"", '--norestore', "--accept=socket,host=localhost,port=$port;urp;StarOffice.ServiceManager", '--impress') -PassThru
    $env:YADG_M0018_IMPRESS_PORT = [string]$port
    $env:YADG_M0018_IMPRESS_VERSION = [string]$version
    $env:YADG_M0018_PAUSE_FOR_VISUAL_INSPECTION = '1'
    & $python $inspectScript $template $source $openCopy $candidate $evidence $repo $revision
    if ($LASTEXITCODE -ne 0) { throw 'Real LibreOffice Impress did not open and inspect the generated candidate cleanly.' }
    if (-not $process.WaitForExit(30000)) { Stop-Process -Id $process.Id -Force; $process.WaitForExit() }

    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $resolvedProfile = [IO.Path]::GetFullPath($profile)
    if (-not $resolvedProfile.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedProfile) -notmatch '^yadg-m0018-impress-[0-9a-f]{32}$') { throw 'Refusing to remove an Impress profile outside the expected M0018 temporary path.' }
    if (Test-Path -LiteralPath $resolvedProfile) { Remove-Item -LiteralPath $resolvedProfile -Recurse -Force }
    $profileCleaned = -not (Test-Path -LiteralPath $resolvedProfile)
    if (-not $profileCleaned) { throw 'Isolated Impress profile cleanup failed.' }
    Remove-Item -LiteralPath $openCopy -Force
    $record = Get-Content -LiteralPath $evidence -Raw | ConvertFrom-Json -DateKind String
    $record | Add-Member -NotePropertyName profileCleaned -NotePropertyValue $profileCleaned -Force
    $record | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $evidence -Encoding utf8
    $result = [string]$record.result
    if ($result -ne 'passed') { throw 'LibreOffice Impress candidate inspection returned a failed result.' }
} catch {
    $failure = $_.Exception.Message
    if (Test-Path -LiteralPath $evidence) {
        try {
            $record = Get-Content -LiteralPath $evidence -Raw | ConvertFrom-Json -DateKind String
            $record.result = 'failed'
            if ([string]::IsNullOrWhiteSpace([string]$record.failure)) { $record.failure = $failure }
            elseif ([string]$record.failure -notlike "*$failure*") { $record.failure = "$($record.failure); $failure" }
            $record.profileCleaned = $profileCleaned
            $record | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $evidence -Encoding utf8
        } catch {}
    } else {
        [ordered]@{
            milestone='M0018'; gate='VAL-03'; result='failed'; failure=$failure
            repositoryRevision=$revision; application='LibreOffice Impress'; applicationVersion=$version
            templatePath=[IO.Path]::GetRelativePath($repo,$template); templateSha256=$templateHash
            sourcePath=[IO.Path]::GetRelativePath($repo,$source); sourceSha256=$sourceHash
            outputPath=[IO.Path]::GetRelativePath($repo,$candidate)
            outputSha256=if (Test-Path -LiteralPath $candidate) { (Get-FileHash $candidate -Algorithm SHA256).Hash.ToUpperInvariant() } else { $null }
            openResult='failed'; profileCleaned=$false
        } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $evidence -Encoding utf8
    }
} finally {
    Remove-Item Env:YADG_M0018_IMPRESS_PORT -ErrorAction SilentlyContinue
    Remove-Item Env:YADG_M0018_IMPRESS_VERSION -ErrorAction SilentlyContinue
    Remove-Item Env:YADG_M0018_PAUSE_FOR_VISUAL_INSPECTION -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $openCopy) { Remove-Item -LiteralPath $openCopy -Force -ErrorAction SilentlyContinue }
    $profileName = [IO.Path]::GetFileName([IO.Path]::GetFullPath($profile))
    $ownedProcesses = Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'soffice.bin' -or $_.Name -eq 'soffice.exe') -and $_.CommandLine -and $_.CommandLine.Contains($profileName) }
    $ownedProcesses | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ((Test-Path -LiteralPath $profile) -and -not $profileCleaned) {
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
        $resolvedProfile = [IO.Path]::GetFullPath($profile)
        if ($resolvedProfile.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedProfile) -match '^yadg-m0018-impress-[0-9a-f]{32}$') {
            Remove-Item -LiteralPath $resolvedProfile -Recurse -Force -ErrorAction SilentlyContinue
            $profileCleaned = -not (Test-Path -LiteralPath $resolvedProfile)
        }
    }
}
if ($result -ne 'passed') { throw "M0018 Impress Tier 3 failed: $failure. Receipt: $evidence" }
Write-Output "M0018 Impress-origin Tier 3 passed: $evidence"
