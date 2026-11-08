param(
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/powerpoint-origin.pptx')
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$output = [IO.Path]::GetFullPath($OutputPath)
$fixtureDirectory = Join-Path $repo 'tests/fixtures/m0018'
$provenancePath = Join-Path $fixtureDirectory 'powerpoint-origin.provenance.json'
if (Test-Path -LiteralPath $output) { throw "Refusing to overwrite an existing canonical fixture: $output" }
if (Test-Path -LiteralPath $provenancePath) { throw "Refusing to overwrite existing fixture provenance: $provenancePath" }
New-Item -ItemType Directory -Force -Path $fixtureDirectory | Out-Null

$app = $null
$presentation = $null
$appWasRunning = $false
$saved = $false
try {
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $appWasRunning = $true }
    catch { $app = New-Object -ComObject PowerPoint.Application }
    $app.Visible = -1
    $presentation = $app.Presentations.Add(-1)
    $master = $presentation.Designs.Item(1).SlideMaster
    $layouts = $master.CustomLayouts
    $presentation.PageSetup.SlideSize = 15 # ppSlideSizeWide, PowerPoint's native widescreen page.

    function Set-RoleSamples($shape, [string] $marker, [bool] $includeHeading, [bool] $includeBullet) {
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add($marker)
        $lines.Add('ordinary paragraph example')
        if ($includeHeading) { $lines.Add('content heading example') }
        if ($includeBullet) { $lines.Add('list formatting example') }
        $shape.TextFrame.WordWrap = -1
        $shape.TextFrame.TextRange.Text = $lines -join "`r"
        $paragraph = $shape.TextFrame.TextRange.Paragraphs(2)
        $paragraph.Font.Size = 18
        if ($includeHeading) {
            $heading = $shape.TextFrame.TextRange.Paragraphs(3)
            $heading.Font.Size = 28
        }
        if ($includeBullet) {
            $bulletIndex = if ($includeHeading) { 4 } else { 3 }
            $bullet = $shape.TextFrame.TextRange.Paragraphs($bulletIndex)
            $bullet.Font.Size = 18
            $bulletFormat = $bullet.ParagraphFormat.Bullet
            $bulletFormat.Visible = $true
            $bulletFormat.Type = 1
            $bulletFormat.Character = 42
            $bullet.IndentLevel = 1
        }
    }

    function Set-Title($slide) {
        $title = $slide.Shapes.Placeholders.Item(1)
        $title.TextFrame.TextRange.Text = '{{title}}'
    }

    function Add-PrototypeMarker($slide, [string] $layoutId) {
        $marker = $slide.Shapes.AddTextbox(1, 820, 8, 132, 16)
        $marker.Name = "YADG prototype $layoutId"
        $marker.TextFrame.TextRange.Text = "{{slide:$layoutId}}"
        $marker.TextFrame.TextRange.Font.Size = 6
        $marker.TextFrame.TextRange.Font.Color.RGB = 8421504
        $marker.Fill.Visible = 0
        $marker.Line.Visible = 0
    }

    function Add-StaticFooter($slide) {
        $footer = $slide.Shapes.AddTextbox(1, 40, 515, 500, 16)
        $footer.Name = 'Native template footer'
        $footer.TextFrame.TextRange.Text = 'M0018 compatibility fixture · PowerPoint native template'
        $footer.TextFrame.TextRange.Font.Size = 9
        $footer.TextFrame.TextRange.Font.Italic = -1
        $footer.Fill.Visible = 0
        $footer.Line.Visible = 0
    }

    function Clear-BodyPlaceholders($slide) {
        for ($index = $slide.Shapes.Count; $index -ge 1; $index--) {
            $shape = $slide.Shapes.Item($index)
            try {
                if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -ne 1) { $shape.Delete() }
            } catch {}
        }
    }

    function Add-TextRegion($slide, [string] $name, [string] $marker, [single] $left, [single] $top, [single] $width, [single] $height) {
        $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $shape.Name = $name
        Set-RoleSamples $shape $marker $true $true
        return $shape
    }

    function Add-TitleAndContent($layout, [string] $layoutId, [string] $titleText, [bool] $heading, [bool] $bullet) {
        $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layout)
        Set-Title $slide
        $slide.Shapes.Placeholders.Item(1).Name = "YADG title $layoutId"
        Clear-BodyPlaceholders $slide
        Add-PrototypeMarker $slide $layoutId
        Add-StaticFooter $slide
        return $slide
    }

    $contentSlide = Add-TitleAndContent $layouts.Item(2) 'content' 'Narrative overview' $true $true
    Add-TextRegion $contentSlide 'YADG content region' '{{content}}' 48 138 864 315 | Out-Null

    $columnsSlide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layouts.Item(4))
    Set-Title $columnsSlide
    $columnsSlide.Shapes.Placeholders.Item(1).Name = 'YADG title two-columns'
    Clear-BodyPlaceholders $columnsSlide
    Add-TextRegion $columnsSlide 'YADG left region' '{{left}}' 48 138 420 315 | Out-Null
    Add-TextRegion $columnsSlide 'YADG right region' '{{right}}' 492 138 420 315 | Out-Null
    Add-PrototypeMarker $columnsSlide 'two-columns'
    Add-StaticFooter $columnsSlide

    $visualSlide = Add-TitleAndContent $layouts.Item(2) 'text-image' 'Architecture map' $true $true
    Add-TextRegion $visualSlide 'YADG text region' '{{text}}' 42 138 420 315 | Out-Null
    $image = $visualSlide.Shapes.AddShape(1, 500, 155, 390, 275)
    $image.Name = 'YADG image region'
    $image.TextFrame.TextRange.Text = '{{image}}'
    $image.TextFrame.TextRange.Font.Size = 18
    $image.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $image.Fill.Visible = 0
    $image.Line.Visible = -1

    $tableSlide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layouts.Item(6))
    Set-Title $tableSlide
    $tableSlide.Shapes.Placeholders.Item(1).Name = 'YADG title table'
    Add-PrototypeMarker $tableSlide 'table'
    Add-StaticFooter $tableSlide
    $frame = $tableSlide.Shapes.AddTable(2, 2, 48, 145, 864, 300)
    $frame.Name = 'YADG native editable table'
    $frame.Table.Rows.Item(1).Height = 64
    $frame.Table.Rows.Item(2).Height = 64
    $frame.Table.Columns.Item(1).Width = 460
    $frame.Table.Columns.Item(2).Width = 404
    $tableValues = @(@('{{table}}','header formatting example'),@('body formatting example','body formatting example'))
    for ($row = 1; $row -le 2; $row++) {
        for ($column = 1; $column -le 2; $column++) {
            $cell = $frame.Table.Cell($row, $column).Shape
            $cell.TextFrame.TextRange.Text = $tableValues[$row - 1][$column - 1]
            $cell.TextFrame.TextRange.Font.Size = 16
            if ($row -eq 1) { $cell.TextFrame.TextRange.Font.Bold = -1 }
            else { $cell.TextFrame.TextRange.Font.Italic = -1 }
        }
    }

    $presentation.SaveAs($output, 24) # ppSaveAsOpenXMLPresentation
    $saved = $true
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Get-Process POWERPNT | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1 -ExpandProperty Path)).ProductVersion
    $created = [DateTimeOffset]::UtcNow.ToString('O')
    $record = [ordered]@{
        id = 'powerpoint-origin'
        path = [IO.Path]::GetRelativePath($repo, $output)
        originApplication = 'Microsoft PowerPoint'
        applicationVersion = $version
        platform = [Runtime.InteropServices.RuntimeInformation]::OSDescription
        createdAtUtc = $created
        saveDateUtc = $created
        creationMethod = 'Created as a new presentation through the real desktop Microsoft PowerPoint COM object model; native default master/theme/layouts and placeholders were retained and the four prototype slides were composed in PowerPoint before saving as PPTX.'
        syntheticAndRedistributionSafe = $true
        createdInApplication = $true
        notOpenXmlGeneratedOrRoundTripped = $true
        retainedNativeMasterThemeLayouts = $true
        layouts = @('content','two-columns','text-image','table')
        sha256 = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash.ToUpperInvariant()
    }
    $record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $provenancePath -Encoding utf8
    Write-Output "Created PowerPoint-origin fixture: $output"
    Write-Output "Provenance: $provenancePath"
} finally {
    if ($presentation) { try { $presentation.Close() } catch {} }
    if ($app -and -not $appWasRunning) { try { $app.Quit() } catch {} }
    if (-not $saved -and (Test-Path -LiteralPath $output)) { Remove-Item -LiteralPath $output -Force }
}
