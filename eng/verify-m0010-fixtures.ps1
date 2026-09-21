$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $repo 'tests/fixtures/m0010'
$records = @()
foreach ($file in @('word-origin','libreoffice-origin')) {
    $record = Get-Content (Join-Path $dir "$file.provenance.json") -Raw | ConvertFrom-Json
    $record | Add-Member -NotePropertyName id -NotePropertyValue $file
    $path = Join-Path $repo ([string]$record.path)
    $actual = (Get-FileHash $path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne ([string]$record.sha256).ToUpperInvariant()) { throw "Fixture hash mismatch: $path" }
    $records += $record
}
[pscustomobject]@{ fixtures = $records } | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $dir 'provenance.json')
Write-Output 'M0010 fixture provenance and hashes verified.'
