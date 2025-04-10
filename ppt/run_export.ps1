Add-type -AssemblyName office
$application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

$scriptDir = $PSScriptRoot
$filename = "Save_to_canvas.pptm"
$path = Join-Path $scriptDir $filename 
$presentation = $application.Presentations.open($path)

$presentation.application.Run("Save_to_canvas.pptm!ExportSlideWithLinkData") 
$application.Quit()     

#Popup box to show completion - you would remove this if using task scheduler 
#$wshell = New-Object -ComObject Wscript.Shell $wshell.Popup("Operation Completed",0,"Done",0x1)  

#exit 