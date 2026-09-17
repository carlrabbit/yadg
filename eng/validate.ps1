$ErrorActionPreference = 'Stop'
dotnet restore "$PSScriptRoot/../Yadg.sln"
dotnet build "$PSScriptRoot/../Yadg.sln" --no-restore --configuration Release
dotnet test "$PSScriptRoot/../Yadg.sln" --no-build --configuration Release
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
