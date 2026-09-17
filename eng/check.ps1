$ErrorActionPreference = 'Stop'
dotnet run --project "$PSScriptRoot/../src/Yadg.Cli/Yadg.Cli.csproj" -- check @args
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
