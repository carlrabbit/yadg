param(
    [string] $OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/fixtures/m0018/powerpoint-origin.pptx'),
    [string] $ProvenancePath
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$output = [IO.Path]::GetFullPath($OutputPath)
$fixtureDirectory = Join-Path $repo 'tests/fixtures/m0018'
$provenancePath = if ([string]::IsNullOrWhiteSpace($ProvenancePath)) { Join-Path (Split-Path -Parent $output) 'powerpoint-origin.provenance.json' } else { [IO.Path]::GetFullPath($ProvenancePath) }
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
    # The prototype coordinates below are authored for a 960 x 540 pt widescreen
    # canvas. Set explicit custom dimensions instead of relying on enum 15, which
    # PowerPoint interprets as a smaller 720 x 405 pt page on this Office build.
    $presentation.PageSetup.SlideSize = 7 # ppSlideSizeCustom
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540

    function Copy-TextFormatting($source, $target) {
        foreach ($property in @('Name', 'Size', 'Bold', 'Italic', 'Underline')) {
            try { $target.Font.$property = $source.Font.$property }
            catch { throw "Could not copy native PowerPoint $property formatting from the application exemplar." }
        }
        try { $target.Font.Color.RGB = $source.Font.Color.RGB }
        catch { throw 'Could not copy native PowerPoint font color from the application exemplar.' }
    }

    function Set-RoleSamples($shape, [string] $marker, $titleExemplar, [bool] $includeHeading, [bool] $includeBullet) {
        $lines = [Collections.Generic.List[string]]::new()
        $lines.Add($marker)
        $lines.Add('ordinary paragraph example')
        if ($includeHeading) { $lines.Add('content heading example') }
        if ($includeBullet) { $lines.Add('list formatting example') }
        $shape.TextFrame.WordWrap = -1
        $shape.TextFrame.TextRange.Text = $lines -join "`r"
        if ($includeHeading) {
            $heading = $shape.TextFrame.TextRange.Paragraphs(3)
            Copy-TextFormatting $titleExemplar $heading
        }
        if ($includeBullet) {
            $bulletIndex = if ($includeHeading) { 4 } else { 3 }
            $bullet = $shape.TextFrame.TextRange.Paragraphs($bulletIndex)
            # A content placeholder supplies PowerPoint's native outline-level
            # bullet, indent and font. Use its second level as the nested-list exemplar.
            $bullet.IndentLevel = 2
            # Materialize the placeholder's inherited bullet on this paragraph
            # so the example remains visible to Open XML consumers as well.
            $bullet.ParagraphFormat.Bullet.Visible = -1
        }
    }

    function Set-Title($slide) {
        $title = $slide.Shapes.Placeholders.Item(1)
        $title.TextFrame.TextRange.Text = '{{title}}'
        return $title
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

    function Get-BodyPlaceholders($slide) {
        $result = [Collections.Generic.List[object]]::new()
        foreach ($shape in $slide.Shapes) {
            if ($shape.Type -ne 14) { continue }
            $placeholderType = [int]$shape.PlaceholderFormat.Type
            if ($placeholderType -in @(2, 7)) { $result.Add($shape) }
        }
        return $result.ToArray()
    }

    function Set-TextRegion($shape, [string] $name, [string] $marker, $titleExemplar) {
        $shape.Name = $name
        Set-RoleSamples $shape $marker $titleExemplar.TextFrame.TextRange.Paragraphs(1) $true $true
    }

    function Add-TitleAndContent($layout, [string] $layoutId) {
        $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layout)
        $title = Set-Title $slide
        $slide.Shapes.Placeholders.Item(1).Name = "YADG title $layoutId"
        $body = @(Get-BodyPlaceholders $slide)
        if ($body.Count -ne 1) { throw "Expected one native PowerPoint body placeholder for layout '$layoutId'; found $($body.Count)." }
        Set-TextRegion $body[0] "YADG $layoutId region" (('{{' + $layoutId + '}}')) $title
        Add-PrototypeMarker $slide $layoutId
        Add-StaticFooter $slide
        return [pscustomobject]@{ Slide = $slide; Title = $title; Body = $body[0] }
    }

    $contentPrototype = Add-TitleAndContent $layouts.Item(2) 'content'
    $contentPrototype.Body.Name = 'YADG content region'

    $columnsSlide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layouts.Item(4))
    $columnsTitle = Set-Title $columnsSlide
    $columnsSlide.Shapes.Placeholders.Item(1).Name = 'YADG title two-columns'
    $columnPlaceholders = @(Get-BodyPlaceholders $columnsSlide)
    if ($columnPlaceholders.Count -ne 2) { throw "Expected two native PowerPoint body placeholders for 'two-columns'; found $($columnPlaceholders.Count)." }
    Set-TextRegion $columnPlaceholders[0] 'YADG left region' '{{left}}' $columnsTitle
    Set-TextRegion $columnPlaceholders[1] 'YADG right region' '{{right}}' $columnsTitle
    Add-PrototypeMarker $columnsSlide 'two-columns'
    Add-StaticFooter $columnsSlide

    $visualPrototype = Add-TitleAndContent $layouts.Item(2) 'text-image'
    $visualSlide = $visualPrototype.Slide
    $visualRegion = $visualPrototype.Body
    $visualRegion.Name = 'YADG text region'
    # Keep the native content placeholder and its outline formatting; adjust
    # only its geometry to make room for the separate editable picture region.
    $visualRegion.Left = 42
    $visualRegion.Top = 138
    $visualRegion.Width = 420
    $visualRegion.Height = 315
    Set-RoleSamples $visualRegion '{{text}}' $visualPrototype.Title.TextFrame.TextRange.Paragraphs(1) $true $true
    $image = $visualSlide.Shapes.AddShape(1, 500, 155, 390, 275)
    $image.Name = 'YADG image region'
    $image.TextFrame.TextRange.Text = '{{image}}'
    $nativeImageLabelExample = $visualRegion.TextFrame.TextRange.Paragraphs(2)
    Copy-TextFormatting $nativeImageLabelExample $image.TextFrame.TextRange
    try { $image.TextFrame.TextRange.ParagraphFormat.Alignment = $nativeImageLabelExample.ParagraphFormat.Alignment }
    catch { throw 'Could not copy native PowerPoint paragraph alignment to the image-region marker.' }
    $image.Fill.Visible = 0
    $image.Line.Visible = -1

    $tableSlide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, $layouts.Item(6))
    $tableTitle = Set-Title $tableSlide
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
        creationMethod = 'Created as a new presentation through the real desktop Microsoft PowerPoint COM object model with a 960 x 540 pt custom widescreen canvas matching the authored native shape geometry; native default master/theme/layouts and title/body placeholders were retained, heading examples copied native title formatting, a native theme table was authored, and the four prototype slides were saved from PowerPoint as PPTX.'
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
