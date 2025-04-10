Add-type -AssemblyName office
$application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$scriptDir = $PSScriptRoot
$filename = "Save_to_canvas.pptm"
$path = Join-Path $scriptDir $filename 
$presentation = $application.Presentations.open($path)

$presentation.application.Run("Save_to_canvas.pptm!ExportSlideWithLinkData") 
$application.Quit()     