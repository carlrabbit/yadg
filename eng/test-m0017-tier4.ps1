param(
    [string] $PackagePath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/package/Yadg.1.2.0.nupkg'),
    [string] $EvidenceRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) 'artifacts/review/evidence/M0017')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$package = [IO.Path]::GetFullPath($PackagePath)
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) { throw "M0017 Tier 4 package is missing: $package" }
$sha256 = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToUpperInvariant()
$packageDir = Split-Path -Parent $package
$isolated = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0017-tool-' + [guid]::NewGuid().ToString('N'))
$workspace = Join-Path ([IO.Path]::GetTempPath()) ('yadg-m0017-consumer-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $isolated, $workspace, (Join-Path $workspace 'YadgTemplates'), (Join-Path $workspace 'images'), $EvidenceRoot | Out-Null
try {
    & dotnet tool install --tool-path $isolated --add-source $packageDir Yadg --version 1.2.0
    if ($LASTEXITCODE -ne 0) { throw 'M0017 isolated dotnet tool install failed.' }
    $installed = Join-Path $isolated 'yadg.exe'
    if (-not (Test-Path -LiteralPath $installed)) { throw 'The isolated tool install did not create yadg.exe.' }
    $template = Join-Path $EvidenceRoot 'word-authored-reference-template.docx'
    if (-not (Test-Path -LiteralPath $template)) { throw 'Run M0017 Tier 3 first; its Word-authored template is needed for the consumer scenario.' }
    Copy-Item -LiteralPath $template -Destination (Join-Path $workspace 'YadgTemplates/template.docx')
    $diagram = Join-Path $EvidenceRoot 'system-flow.png'
    if (-not (Test-Path -LiteralPath $diagram)) { throw 'Run M0017 Tier 3 first; its representative flow image is needed for the consumer scenario.' }
    Copy-Item -LiteralPath $diagram -Destination (Join-Path $workspace 'images/flow.png')
    Set-Content -LiteralPath (Join-Path $workspace 'YADG.md') -Encoding utf8 -Value @('---','yadg:','  version: 1','values:','  story-value: "M0017 Tier 4"','references:','  headings: number','  figures: label-number','  tables: label-number','---')
    Set-Content -LiteralPath (Join-Path $workspace 'content.md') -Encoding utf8 -Value @('# Architecture {#architecture}', '', 'Read [the architecture explanation](#architecture), [the system flow](#system-flow), [the interface matrix](#interface-matrix), and [the external guide](https://example.test/yadg).', '', 'See [@system-flow] and [@interface-matrix].', '', '![System flow](images/flow.png){#system-flow}', '', '| Component | Protocol |', '|---|---|', '| Gateway | HTTPS |', '| Worker | HTTPS |', '{#interface-matrix caption="Interface matrix"}', '', '- [x] outer task', '  1. ordered child', '     - nested unordered', '', '> quoted **text**', '>', '> > nested quote', '', '```text', '  literal spacing', '{#literal-code-line}', '```', '', '~~strike~~ `inline-code`', '', '+---+---+---+', '| AAAAA | B |', '+---+---+ B +', '| D | E | B |', '+ D +---+---+', '| D | CCCCC |', '+---+---+---+', '{#merged-grid}')

    $commands = @()
    foreach ($arguments in @(@('check','--workspace',$workspace), @('build','--workspace',$workspace))) {
        $start = [Diagnostics.ProcessStartInfo]::new($installed)
        $start.UseShellExecute = $false
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($argument in $arguments) { $start.ArgumentList.Add([string]$argument) }
        $process = [Diagnostics.Process]::Start($start)
        if ($null -eq $process) { throw "Unable to invoke isolated yadg $($arguments[0])." }
        try {
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit()
            $commands += [ordered]@{ command = 'yadg ' + ($arguments -join ' '); exitCode = $process.ExitCode; stdout = $stdout.GetAwaiter().GetResult(); stderr = $stderr.GetAwaiter().GetResult() }
            if ($process.ExitCode -ne 0) { throw "M0017 installed yadg $($arguments[0]) failed: $($commands[-1].stderr)" }
        }
        finally { $process.Dispose() }
    }
    $authored = Get-ChildItem -LiteralPath (Join-Path $workspace 'YadgPreWords') -Filter '*.docx' | Select-Object -First 1
    if ($null -eq $authored) { throw 'Isolated installed tool produced no authored DOCX.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($authored.FullName)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        $reader = [IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        $ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $text = ($xml.SelectNodes('//w:t',$ns) | ForEach-Object { $_.InnerText }) -join ''
        foreach ($expected in @('the architecture explanation','the system flow','the interface matrix','the external guide')) { if (-not $text.Contains($expected,[StringComparison]::Ordinal)) { throw "Installed package output lost '$expected'." } }
        $instructions = @($xml.SelectNodes('//w:instrText',$ns) | ForEach-Object { $_.InnerText })
        if (@($instructions | Where-Object { $_ -match 'REF\s+' }).Count -lt 2) { throw 'Installed package output is missing one or more generated REF fields.' }
        $links = @($xml.SelectNodes('//w:hyperlink[@w:anchor]',$ns))
        if ($links.Count -lt 3) { throw 'Installed package output is missing semantic local hyperlinks.' }
        foreach ($expected in @('outer task','ordered child','nested unordered','nested quote','literal spacing','{#literal-code-line}','strike','inline-code','AAAAA','CCCCC')) { if (-not $text.Contains($expected,[StringComparison]::Ordinal)) { throw "M0017 installed package output lost '$expected'." } }
        if (@($xml.SelectNodes('//w:strike',$ns)).Count -lt 1 -or @($xml.SelectNodes('//w:gridSpan[@w:val="2"]',$ns)).Count -lt 1 -or @($xml.SelectNodes('//w:vMerge',$ns)).Count -lt 1) { throw 'M0017 installed package output lost strike or grid merge structures.' }
    }
    finally { $zip.Dispose() }

    $evidence = [ordered]@{ milestone = 'M0017'; packagePath = [IO.Path]::GetRelativePath($repo, $package).Replace([char]92,[char]47); packageSha256 = $sha256; packageVersion = '1.2.0'; isolatedToolPath = 'temporary dotnet tool path'; isolatedWorkspace = 'temporary consumer workspace'; installedCommand = $installed; commands = $commands; authoredDocxSha256 = (Get-FileHash -LiteralPath $authored.FullName -Algorithm SHA256).Hash.ToUpperInvariant(); result = 'passed' }
    $evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'tier4-evidence.json') -Encoding utf8
    Write-Output "M0017 isolated packed-tool check/build passed for Yadg 1.2.0 ($sha256)."
}
finally {
    Remove-Item -LiteralPath $isolated -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $workspace -Recurse -Force -ErrorAction SilentlyContinue
}
