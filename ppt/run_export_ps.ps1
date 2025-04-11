function Export-SlideWithLinkData {
    param (
        [string]$pptxPath,  # Path to the PowerPoint file
        [string]$outputPath  # Path to save exported results
    )

    # Create PowerPoint COM object
    $ppApp = New-Object -ComObject PowerPoint.Application
    $ppApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Open the presentation
    $presentation = $ppApp.Presentations.Open($pptxPath)

    # Get the active slide (use the slide index directly)
    $slideIndex = 2  
    $slide = $presentation.Slides.Item($slideIndex)  # Use slide index directly

    # Set DPI factor (1.5x or 2x for HiDPI)
    $dpiFactor = 2

    # Get slide width and height
    $slideWidth = $presentation.PageSetup.SlideWidth
    $slideHeight = $presentation.PageSetup.SlideHeight

    # Export slide as PNG with high resolution
    $imgPath = Join-Path $outputPath "slide_$slideIndex.png"
    $slide.Export($imgPath, "PNG", $slideWidth * $dpiFactor, $slideHeight * $dpiFactor)

    # Start JS metadata structure for link data
    $linkData = @"
const slideMetadata = {
  width: $slideWidth,
  height: $slideHeight,
  links: [
"@

    # Collect link areas from slide shapes
    $hasLink = $false
    foreach ($shape in $slide.Shapes) {
        if ($shape.Type -eq 13) {  # If shape is a picture (msoPicture)
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
        # Remove last comma from linkData
        $linkData = $linkData.TrimEnd(",`r`n") + "`r`n"
    }

    # Close the JS object structure
    $linkData += @"
  ]
};
"@

    # Save metadata to slideData.js file
    $jsonJSPath = Join-Path $outputPath "slideData.js"
    $linkData | Out-File -FilePath $jsonJSPath -Encoding utf8

    # Clean up and close PowerPoint application
    $presentation.Close()
    $ppApp.Quit()

    # Release COM objects to ensure PowerPoint closes fully
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppApp) | Out-Null

    # Clean up
    [GC]::Collect()  # Force garbage collection
    [GC]::WaitForPendingFinalizers()  # Wait for finalizer threads to complete

    Write-Host "Slide exported as PNG and metadata saved to slideData.js"
}

# Example usage:
$pptxPath = "C:\temp\HomeSite.pptx"
$outputPath = "C:\temp"
Export-SlideWithLinkData -pptxPath $pptxPath -outputPath $outputPath
