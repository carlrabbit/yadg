$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$soffice = Join-Path $env:ProgramFiles 'LibreOffice/program/soffice.exe'
$python = Join-Path $env:ProgramFiles 'LibreOffice/program/python.exe'
if (-not (Test-Path -LiteralPath $soffice -PathType Leaf)) { throw 'LibreOffice Impress is unavailable.' }
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'LibreOffice bundled UNO Python is unavailable.' }
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo($soffice).ProductVersion
if ([string]::IsNullOrWhiteSpace($version)) { throw 'Could not read the installed LibreOffice product version.' }
$profile = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0018-fixture-' + [guid]::NewGuid().ToString('N'))
$profileName = [IO.Path]::GetFileName($profile)
$profileUri = ([uri]([IO.Path]::GetFullPath($profile) + [IO.Path]::DirectorySeparatorChar)).AbsoluteUri.TrimEnd('/')
$port = Get-Random -Minimum 22000 -Maximum 42000
$process = $null
$env:YADG_M0018_IMPRESS_VERSION = $version
$env:YADG_M0018_IMPRESS_PORT = [string]$port
try {
    $process = Start-Process -FilePath $soffice -ArgumentList @("-env:UserInstallation=`"$profileUri`"", '--norestore', "--accept=socket,host=localhost,port=$port;urp;StarOffice.ServiceManager", '--impress') -PassThru -WindowStyle Hidden
    & $python (Join-Path $PSScriptRoot 'create-m0018-impress-fixture.py')
    if ($LASTEXITCODE -ne 0) { throw 'LibreOffice Impress did not create the canonical fixture through UNO.' }
    Write-Output "Created Impress-origin fixture with LibreOffice $version."
} finally {
    Remove-Item Env:YADG_M0018_IMPRESS_VERSION -ErrorAction SilentlyContinue
    Remove-Item Env:YADG_M0018_IMPRESS_PORT -ErrorAction SilentlyContinue
    $ownedProcesses = @(Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'soffice.bin' -or $_.Name -eq 'soffice.exe') -and $_.CommandLine -and $_.CommandLine.Contains($profileName) })
    foreach ($owned in $ownedProcesses) { Stop-Process -Id $owned.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($process -and -not $process.HasExited) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue }
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    while (@(Get-CimInstance Win32_Process | Where-Object { ($_.Name -eq 'soffice.bin' -or $_.Name -eq 'soffice.exe') -and $_.CommandLine -and $_.CommandLine.Contains($profileName) }).Count -gt 0 -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 250 }
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $resolvedProfile = [IO.Path]::GetFullPath($profile)
    if (-not $resolvedProfile.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedProfile) -cne $profileName -or $profileName -notmatch '^yadg-m0018-fixture-[0-9a-f]{32}$') { throw 'Refusing to remove an Impress fixture profile outside its task-owned temporary directory.' }
    if (Test-Path -LiteralPath $resolvedProfile) { Remove-Item -LiteralPath $resolvedProfile -Recurse -Force }
}
