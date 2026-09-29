$ErrorActionPreference = 'Stop'
dotnet restore "$PSScriptRoot/../Yadg.slnx"
$vsMsbuild = @(
    'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe'
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($vsMsbuild) { & $vsMsbuild "$PSScriptRoot/../Yadg.slnx" /t:Build /p:Configuration=Release /p:Restore=false /v:minimal } else { dotnet build "$PSScriptRoot/../Yadg.slnx" --no-restore --configuration Release }
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
dotnet test "$PSScriptRoot/../Yadg.slnx" --no-build --configuration Release
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& "$PSScriptRoot/verify-m0012-fixtures.ps1"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
$repo = Split-Path -Parent $PSScriptRoot
$trackedOperational = & git -C $repo ls-files | Where-Object { $_ -match '^(\.execution|\.review|artifacts)/' }
if ($trackedOperational) { throw "Operational state is tracked: $($trackedOperational -join ', ')" }
Write-Output 'Repository hygiene verified: execution, review, and artifacts state is untracked.'
