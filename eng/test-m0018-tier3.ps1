param(
    [Parameter(Mandatory=$true)][string] $TemplatePath,
    [Parameter(Mandatory=$true)][string[]] $SourcePath,
    [Parameter(Mandatory=$true)][ValidateSet('Microsoft PowerPoint','LibreOffice Impress')][string] $TemplateOriginApplication,
    [Parameter(Mandatory=$true)][string] $TemplateOriginVersion,
    [Parameter(Mandatory=$true)][string] $TemplateCreationMethod,
    [Parameter(Mandatory=$true)][string] $TemplateCreatedAtUtc,
    [Parameter(Mandatory=$true)][switch] $ConfirmTemplateSyntheticAndRedistributionSafe,
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/review-sample.pptx'),
    [string] $EvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0018/tier3-evidence.json')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$template = [IO.Path]::GetFullPath($TemplatePath)
$sources = @($SourcePath | ForEach-Object { [IO.Path]::GetFullPath($_) })
$output = [IO.Path]::GetFullPath($OutputPath)
$evidence = [IO.Path]::GetFullPath($EvidencePath)
$cli = Join-Path $repo 'src/Yadg.Cli/bin/Release/net10.0/yadg.dll'
$soffice = Join-Path $env:ProgramFiles 'LibreOffice/program/soffice.exe'
if (-not (Test-Path -LiteralPath $template -PathType Leaf) -or [IO.Path]::GetExtension($template) -ine '.pptx') { throw 'M0018 Tier 3 requires an existing PPTX prototype created and saved in PowerPoint or Impress.' }
if ($sources.Count -eq 0 -or @($sources | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) { throw 'M0018 Tier 3 requires one or more existing Markdown sources.' }
if ([string]::IsNullOrWhiteSpace($TemplateOriginVersion) -or [string]::IsNullOrWhiteSpace($TemplateCreationMethod) -or -not $ConfirmTemplateSyntheticAndRedistributionSafe) { throw 'M0018 Tier 3 requires complete synthetic, redistribution-safe template provenance.' }
$createdAt = [DateTimeOffset]::MinValue
if (-not [DateTimeOffset]::TryParse($TemplateCreatedAtUtc, [ref]$createdAt) -or $createdAt.Offset -ne [TimeSpan]::Zero) { throw 'TemplateCreatedAtUtc must be a UTC timestamp with a zero offset.' }
if (-not (Test-Path -LiteralPath $cli -PathType Leaf)) { throw "Build the Release CLI before Tier 3: $cli" }
if (-not (Test-Path -LiteralPath $soffice -PathType Leaf)) { throw 'M0018 Tier 3 is blocked: LibreOffice Impress is unavailable in this Windows session.' }

$outputDirectory = Split-Path -Parent $output
$evidenceDirectory = Split-Path -Parent $evidence
New-Item -ItemType Directory -Force -Path $outputDirectory, $evidenceDirectory | Out-Null
$arguments = @($cli, 'presentation', 'check', '--template', $template)
foreach ($source in $sources) { $arguments += @('--source', $source) }
& dotnet @arguments
if ($LASTEXITCODE -ne 0) { throw 'M0018 Tier 3 candidate failed direct presentation pre-flight.' }
$arguments = @($cli, 'presentation', 'build', '--template', $template)
foreach ($source in $sources) { $arguments += @('--source', $source) }
$arguments += @('--output', $output)
& dotnet @arguments
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $output -PathType Leaf)) { throw 'M0018 Tier 3 candidate build failed.' }

$profile = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0018-impress-' + [guid]::NewGuid().ToString('N'))
$profileUri = ([uri]([IO.Path]::GetFullPath($profile) + [IO.Path]::DirectorySeparatorChar)).AbsoluteUri.TrimEnd('/')
$version = (& $soffice --version | Out-String).Trim()
$process = Start-Process -FilePath $soffice -ArgumentList @("-env:UserInstallation=`"$profileUri`"", '--norestore', '--impress', "`"$output`"") -PassThru
Write-Host 'Inspect the opened presentation in LibreOffice Impress. Confirm it opens without repair and check native text, picture, table, prototype styling and static elements. Do not treat this step as the blocking human review record.'
$opened = Read-Host 'After closing Impress, enter yes only if you observed a clean open without repair; otherwise enter no'
if (-not $process.HasExited) { try { $process.WaitForExit(30000) | Out-Null } catch {} }
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$resolvedProfile = [IO.Path]::GetFullPath($profile)
if (-not $resolvedProfile.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedProfile) -notmatch '^yadg-m0018-impress-[0-9a-f]{32}$') { throw 'Refusing to remove a profile outside the expected M0018 temporary path.' }
if (Test-Path -LiteralPath $resolvedProfile) { Remove-Item -LiteralPath $resolvedProfile -Recurse -Force }
if ($opened -cne 'yes') { throw 'M0018 Tier 3 failed: reviewer did not confirm a clean Impress open.' }

$record = [ordered]@{
    milestone = 'M0018'
    repositoryRevision = (& git -C $repo rev-parse HEAD).Trim()
    application = 'LibreOffice Impress'
    applicationVersion = $version
    templateOriginApplication = $TemplateOriginApplication
    templateOriginVersion = $TemplateOriginVersion
    templateCreationMethod = $TemplateCreationMethod
    templateCreatedAtUtc = $createdAt.ToString('O')
    platform = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
    syntheticAndRedistributionSafe = $true
    templatePath = [IO.Path]::GetRelativePath($repo, $template)
    templateSha256 = (Get-FileHash -LiteralPath $template -Algorithm SHA256).Hash.ToUpperInvariant()
    sourcePaths = @($sources | ForEach-Object { [IO.Path]::GetRelativePath($repo, $_) })
    sourceSha256 = @($sources | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToUpperInvariant() })
    outputPath = [IO.Path]::GetRelativePath($repo, $output)
    outputSha256 = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToUpperInvariant()
    openResult = 'operator-confirmed-open-without-repair'
    profileCleaned = -not (Test-Path -LiteralPath $resolvedProfile)
}
$record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $evidence -Encoding utf8
Write-Output "M0018 real Impress open recorded: $evidence"
