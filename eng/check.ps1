$ErrorActionPreference = 'Stop'
$project = "$PSScriptRoot/../src/Yadg.Cli/Yadg.Cli.csproj"
$vsMsbuild = @(
    'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe',
    'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe'
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($vsMsbuild) { & $vsMsbuild $project /t:Build /p:Configuration=Release /v:minimal; if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }; dotnet "$PSScriptRoot/../src/Yadg.Cli/bin/Release/net10.0-windows/yadg.dll" check @args }
else { dotnet run --project $project -- check @args }
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
