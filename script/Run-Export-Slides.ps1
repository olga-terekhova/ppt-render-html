# Parameters stored in a file-style script
$pptxPath = "C:\temp\HomeSite.pptx"
$outputPath = "C:\temp\output"
$templatePath = "C:\temp\templates"

# Run the main export script with the parameters
& "$PSScriptRoot\Export-Slides.ps1" -p $pptxPath -o $outputPath -t $templatePath