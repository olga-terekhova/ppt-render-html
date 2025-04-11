function Resolve-SubAddressToSlideIndex {
    param (
        [Microsoft.Office.Interop.PowerPoint.Presentation]$presentation,
        [string]$subAddress
    )

    foreach ($s in $presentation.Slides) {
        if ($s.Name -eq $subAddress) {
            return $s.SlideIndex
        }
        try {
            if ([int]$subAddress -eq $s.SlideIndex) {
                return $s.SlideIndex
            }
        } catch {
            # Not an integer, skip
        }
    }

    return $null
}

function Export-SlideWithLinkData {
    param (
        [object]$slide,
        [int]$slideIndex,
        [string]$outputPath,
        [string]$templatePath,
        [single]$slideWidth,
        [single]$slideHeight,
        [int]$dpiFactor
    )

    # Export slide as PNG
    $imgPath = Join-Path $outputPath "slide_$slideIndex.png"
    $slide.Export($imgPath, "PNG", $slideWidth * $dpiFactor, $slideHeight * $dpiFactor)


    # Ungroup grouped shapes and build a flat shape list
    $flatShapes = @()

    foreach ($shape in $slide.Shapes) {
        if ($shape.Type -eq 6) {  # msoGroup
            try {
                $ungrouped = $shape.Ungroup()
                foreach ($s in $ungrouped) {
                    $flatShapes += $s
                }
            } catch {
                Write-Warning "Failed to ungroup a shape on slide $slideIndex"
            }
        } else {
            $flatShapes += $shape
        }
    }

    # Build JS metadata
    $linkData = @"
const slideMetadata = {
  width: $slideWidth,
  height: $slideHeight,
  links: [
"@

      $hasLink = $false
      foreach ($shape in $flatShapes) {
      $actionSetting = $shape.ActionSettings.Item(1)  # ppMouseClick
      $hyperlink = $actionSetting.Hyperlink
	
        
      if ($actionSetting.Action -eq 7 -and $hyperlink) {  # ppActionHyperlink = 7
        $url = $null

            if ($hyperlink.Address) {
                # External link
                $url = $hyperlink.Address
            } elseif ($hyperlink.SubAddress) {
                # Parse SubAddress: expect something like "265,10,Slide 10"
                $parts = $hyperlink.SubAddress -split ","
                if ($parts.Count -ge 2 -and $parts[1] -match '^\d+$') {
                    $targetIndex = [int]$parts[1]
                    $url = "slide_$targetIndex.html"
                }
            }

            if ($url) {
                $hasLink = $true
                $linkData += @"
    { x: $($shape.Left), y: $($shape.Top), w: $($shape.Width), h: $($shape.Height), url: '$url' },
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

    Write-Host "Processed slide $slideIndex"
}

function Export-PresentationWithLinkData {
    param (
        [string]$pptxPath,
        [string]$outputPath,
        [string]$templatePath
    )

    # Create PowerPoint COM object
    $ppApp = New-Object -ComObject PowerPoint.Application
    $ppApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Open the presentation
    $presentation = $ppApp.Presentations.Open($pptxPath)

    # Get slide dimensions and DPI settings
    $slideWidth = $presentation.PageSetup.SlideWidth
    $slideHeight = $presentation.PageSetup.SlideHeight
    $dpiFactor = 2

    # Copy style.css
    Copy-Item -Path (Join-Path $templatePath "style.css") -Destination $outputPath -Force

    # Process each slide
    $totalSlides = $presentation.Slides.Count
    for ($i = 1; $i -le $totalSlides; $i++) {
        $slide = $presentation.Slides.Item($i)
        Export-SlideWithLinkData -slide $slide -slideIndex $i -outputPath $outputPath -templatePath $templatePath -slideWidth $slideWidth -slideHeight $slideHeight -dpiFactor $dpiFactor
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

# === Example usage ===
$pptxPath = "C:\temp\HomeSite.pptx"
$outputPath = "C:\temp\output"
$templatePath = "C:\temp\templates"

Export-PresentationWithLinkData -pptxPath $pptxPath -outputPath $outputPath -templatePath $templatePath
