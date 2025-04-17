param (
    [Alias("p")][string]$paramPath
)

# Default to ./param/Params.ps1 if not provided
if (-not $paramPath) {
    $paramPath = Join-Path -Path $PSScriptRoot -ChildPath "param\Params.ps1"
}

# Check if parameter file exists
if (-not (Test-Path $paramPath)) {
    Write-Error "Parameter file not found at: $paramPath"
    exit 1
}

# Source the parameter file to get $pptxPath, $outputPath, $templatePath
. $paramPath


# Ensure variables are set
if (-not ($pptxPath -and $outputPath -and $templatePath)) {
    Write-Error "One or more required parameters are missing in $paramPath"
    exit 1
}

# Run the main export script with the parameters
& "$PSScriptRoot\Export-Slides.ps1" -p $pptxPath -o $outputPath -t $templatePath