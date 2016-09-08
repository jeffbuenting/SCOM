#------------------------------------------------------------------------------
# Install-SCOMPrereqs
#
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# Main
#------------------------------------------------------------------------------

Import-Module ServerManager

add-windowsfeature net-framework
add-windowsfeature web-server,web-static-content,web-default-doc,web-http-errors,web-asp-net,web-asp,web-net-ext,web-cgi,web-isapi-ext,web-isapi-filter,web-http-logging,web-request-monitor,web-windows-auth,web-filtering,web-mgmt-tools,web-mgmt-compat,web-metabase,web-wmi,web-lgcy-mgmt-console,web-lgcy-scripting

import-module '\\vbgov.com\deploy\Disaster_Recovery\ActiveDirectory\Scripts\LocalUsersAndComputersModule\LocalUsersAndComputersModule'

# ----- Add Server Local Admin-U group to the Local Administrators Group
$LocalAdmins = Get-LocalGroupMember 'Administrators'
if ( $LocalAdmins  -cnotcontains "comitscom05" ) {   	# ----- Add the group 
	Add-LocalGroupMember -ADGroup 'comitscom05' -localgroup 'Administrators'
	
}


		Write-Host "Press any key to continue ..."

		$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
# ----- Create Temp Directory

if ( (Test-Path "C:\Temp") -eq $False ) {
	Set-Location c:\ 
	New-Item -name temp -ItemType directory -ErrorAction SilentlyContinue
}

Copy-Item '\\vbgov.com\deploy\Disaster_Recovery\SCOM\Scripts\Install-SCOMPrereqs\dotNetFx40_Full_x86_x64.exe' 'C:\Temp'
Copy-Item '\\vbgov.com\deploy\Disaster_Recovery\SCOM\Scripts\Install-SCOMPrereqs\reportviewer.exe' 'C:\Temp'

# ----- Install .Net 4.0
$Parameters = '/q /norestart'
[System.Diagnostics.Process]::Start( 'c:\temp\dotNetFx40_Full_x86_x64.exe',$Parameters )
$Job = (Get-Process | where { $_.processname -eq 'dotNetFx40_Full_x86_x64' })
$Job.waitforexit()

# ----- Install Report Viewer Controls
[System.Diagnostics.Process]::Start( 'c:\temp\ReportViewer.exe',$Parameters )
$Job = (Get-Process | where { $_.processname -eq 'ReportViewer.exe' })
$Job.waitforexit()

# ----- Enable ISAPI and CGI Restrictions
Import-Module webadministration
$frameworkPath = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\aspnet_isapi.dll"  
$isapiConfiguration = get-webconfiguration "/system.webServer/security/isapiCgiRestriction/add[@path='$FrameworkPath']/@allowed"  
if (!$isapiConfiguration.value){  
    set-webconfiguration "/system.webServer/security/isapiCgiRestriction/add[@path='$FrameworkPath']/@allowed" -value "True" -PSPath:IIS:\  
}  

# ----- Cleanup
Remove-Item -Path 'c:\Temp\*.*'