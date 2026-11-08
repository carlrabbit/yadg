$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$soffice = Join-Path $env:ProgramFiles 'LibreOffice/program/soffice.exe'
$python = Join-Path $env:ProgramFiles 'LibreOffice/program/python.exe'
if (-not (Test-Path -LiteralPath $soffice -PathType Leaf)) { throw 'LibreOffice Impress is unavailable.' }
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'LibreOffice bundled UNO Python is unavailable.' }
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo($soffice).ProductVersion
if ([string]::IsNullOrWhiteSpace($version)) { throw 'Could not read the installed LibreOffice product version.' }
$env:YADG_M0018_IMPRESS_VERSION = $version
& $python (Join-Path $PSScriptRoot 'create-m0018-impress-fixture.py')
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
