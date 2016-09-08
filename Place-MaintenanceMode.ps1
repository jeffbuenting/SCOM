
#---------------------------------------------------------------------------------------
# Main
#---------------------------------------------------------------------------------------

Import-Module -Name "\\vbgov.com\deploy\Disaster_Recovery\SCOM\Scripts\SCOMPSModule\SCOMPSModule.psm1" -argumentlist 'vbas022'



Write-Host "Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")