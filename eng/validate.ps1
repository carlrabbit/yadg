$ErrorActionPreference = 'Stop'
dotnet restore "$PSScriptRoot/../Yadg.slnx"
dotnet build "$PSScriptRoot/../Yadg.slnx" --no-restore --configuration Release
dotnet test "$PSScriptRoot/../Yadg.slnx" --no-build --configuration Release
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
