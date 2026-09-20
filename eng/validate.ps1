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
