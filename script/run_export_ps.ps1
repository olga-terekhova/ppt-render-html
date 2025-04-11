function Export-SlideWithLinkData {
    param (
        [string]$pptxPath,        # Path to the PowerPoint file
        [string]$outputPath,      # Path to save exported results
        [int]$slideNumber,        # Number of the slide to process
        [string]$templatePath     # Path to the templates folder
    )

    # Create PowerPoint COM object
    $ppApp = New-Object -ComObject PowerPoint.Application
    $ppApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Open the presentation
    $presentation = $ppApp.Presentations.Open($pptxPath)

    # Use the slide number provided
    $slideIndex = $slideNumber
    $slide = $presentation.Slides.Item($slideIndex)

    # Set DPI factor
    $dpiFactor = 2

    # Get slide dimensions
    $slideWidth = $presentation.PageSetup.SlideWidth
    $slideHeight = $presentation.PageSetup.SlideHeight

    # Export slide as PNG
    $imgPath = Join-Path $outputPath "slide_$slideIndex.png"
    $slide.Export($imgPath, "PNG", $slideWidth * $dpiFactor, $slideHeight * $dpiFactor)

    # Build JS metadata
    $linkData = @"
const slideMetadata = {
  width: $slideWidth,
  height: $slideHeight,
  links: [
"@

    $hasLink = $false
    foreach ($shape in $slide.Shapes) {
        if ($shape.Type -eq 13) {
            $hyperlink = $shape.ActionSettings.Item(1).Hyperlink.Address
            if ($hyperlink) {
                $hasLink = $true
                $linkData += @"
    { x: $($shape.Left), y: $($shape.Top), w: $($shape.Width), h: $($shape.Height), url: '$hyperlink' },
"@
            }
        }
    }

    if ($hasLink) {
        $linkData = $linkData.TrimEnd(",`r`n") + "`r`n"
    }

    $linkData += @"
  ]
};
"@

    # Save slide metadata
    $jsonJSPath = Join-Path $outputPath "slide_${slideIndex}_data.js"
    $linkData | Out-File -FilePath $jsonJSPath -Encoding utf8

    # === Process slide.html template ===
    $slideHtmlPath = Join-Path $templatePath "slide.html"
    $slideHtmlContent = Get-Content $slideHtmlPath -Raw
    $slideHtmlContent = $slideHtmlContent -replace "slideCanvas", "slideCanvas_$slideIndex"
    $slideHtmlContent = $slideHtmlContent -replace "slideData\.js", "slide_${slideIndex}_data.js"
    $slideHtmlContent = $slideHtmlContent -replace "mainScript\.js", "mainScript_${slideIndex}.js"
    $newSlideHtmlPath = Join-Path $outputPath "slide_${slideIndex}.html"
    $slideHtmlContent | Out-File -FilePath $newSlideHtmlPath -Encoding utf8

    # === Process mainScript.js template ===
    $scriptTemplatePath = Join-Path $templatePath "mainScript.js"
    $scriptContent = Get-Content $scriptTemplatePath -Raw
    $scriptContent = $scriptContent -replace "slideCanvas", "slideCanvas_$slideIndex"
    $scriptContent = $scriptContent -replace "slide_1\.png", "slide_${slideIndex}.png"
    $newScriptPath = Join-Path $outputPath "mainScript_${slideIndex}.js"
    $scriptContent | Out-File -FilePath $newScriptPath -Encoding utf8

    # Clean up PowerPoint
    $presentation.Close()
    $ppApp.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppApp) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Host "Slide $slideIndex exported as PNG and metadata. HTML and JS customized from templates."
}

# === Example usage ===
$pptxPath = "C:\temp\HomeSite.pptx"
$outputPath = "C:\temp\output"
$templatePath = "C:\temp\templates"
$slideNumber = 2

Export-SlideWithLinkData -pptxPath $pptxPath -outputPath $outputPath -slideNumber $slideNumber -templatePath $templatePath
