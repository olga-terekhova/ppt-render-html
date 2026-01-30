param (
    [Alias("p")][string]$pptxPath
)


function Export-SlideWithLinkData {
    param (
        [object]$slide,
        [int]$slideIndex,
        [string]$outputPath,
	[string]$baseName,
        [single]$slideWidth,
        [single]$slideHeight,
        [int]$dpiFactor
    )

    # Export slide as PNG
    $imgFileName = $baseName + "_" + $slideIndex + ".png"
    $imgPath = Join-Path $outputPath $imgFileName
    $slide.Export($imgPath, "PNG", $slideWidth * $dpiFactor, $slideHeight * $dpiFactor)

    Write-Host "Processed slide $slideIndex, exported to $imgFileName"
}


function Export-Presentation {
    param (
        [string]$pptxPath
    )

    # Create PowerPoint COM object
    $ppApp = New-Object -ComObject PowerPoint.Application
    $ppApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Open the presentation
    $presentation = $ppApp.Presentations.Open($pptxPath)

    $outputPath = Split-Path -Parent $pptxPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pptxPath)

    # Get slide dimensions and DPI settings
    $slideWidth = $presentation.PageSetup.SlideWidth
    $slideHeight = $presentation.PageSetup.SlideHeight
    $dpiFactor = 2

    # Process each slide
    $totalSlides = $presentation.Slides.Count
    for ($i = 1; $i -le $totalSlides; $i++) {
        $slide = $presentation.Slides.Item($i)
        Export-SlideWithLinkData -slide $slide -slideIndex $i -outputPath $outputPath -baseName $baseName -slideWidth $slideWidth -slideHeight $slideHeight -dpiFactor $dpiFactor
    }

    # Clean up PowerPoint
    $presentation.Close()
    $ppApp.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppApp) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Host "Presentation processed and output saved to $outputPath"
}


if ($pptxPath) {
    Export-Presentation -pptxPath $pptxPath 
}