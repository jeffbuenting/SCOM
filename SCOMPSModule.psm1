#---------------------------------------------------------------------------------------
# SCOMPSModule.psm1
#
# Contains SCOM comandlets
#---------------------------------------------------------------------------------------

Param ( [string]$ManagementServer = "VBAS022",
		[Bool]$Debug = $False)

# --------------------------------------------------------------------------------------
# Function Set-SCOMMaintenanceMode
#
# Places a computer and its associated Health Service and Health Watcher Services into Maintenance Mode
# http://blogs.msdn.com/b/boris_yanushpolsky/archive/2007/07/25/putting-a-computer-into-maintenance-mode.aspx
# --------------------------------------------------------------------------------------

<#
	.SYNOPSIS
		Sets the object and its associated Health agents into Maintenance Mode.
		
	.DESCRIPTION
		Sets the object and its associated Health agent and Health Watcher agent to maintenance mode.  Prevents any of them from throwing alerts
		
	.PARAMETER ComputerPrincipalName
		Specifies the Computers FQDN
		
	.PARAMETER NumberofHoursinMaintenanceMode
		Specifies the number of hours in maintenance mode
	
	.PARAMETER Comment
		Specifies why the object has been placed in maintenance mode
		
	.INPUTS
		None
	
	.OUTPUTS
		None
		
	.EXAMPLE
		set-maintenancemode 'VBVS0030.VBGOV.COM' .25 'Patching'
		
	.LINK
		http://blogs.msdn.com/b/boris_yanushpolsky/archive/2007/07/25/putting-a-computer-into-maintenance-mode.aspx
#>

Function Set-SCOMMaintenanceMode {
	
	param( 	[String]$computerPrincipalName,
			[String]$numberOfHoursInMaintenanceMode,
			[String]$comment )
   
	# ----- Get Computer Object
	$computerClass = get-monitoringclass -name:Microsoft.Windows.Computer
  	$computerCriteria = "PrincipalName='" + $computerPrincipalName + "'"
  	$computer = get-monitoringobject -monitoringclass:$computerClass -criteria:$computerCriteria
 
	# ----- Get Associated Health Service Object
	$healthServiceClass = get-monitoringclass -name:Microsoft.SystemCenter.HealthService
  	$healthServices = $computer.GetRelatedMonitoringObjects($healthServiceClass)
	$healthService = $healthServices[0]
	  
  	# ----- Get Associated Health Service Watcher Object
	$healthServiceWatcherClass = get-monitoringclass -name:Microsoft.SystemCenter.HealthServiceWatcher
  	$healthServiceCriteria = "HealthServiceName='" + $computerPrincipalName + "'"
	$healthServiceWatcher = get-monitoringobject -monitoringclass:$healthServiceWatcherClass -criteria:$healthServiceCriteria
   
	# ----- Set Maintenance Mode Length
	$startTime = [System.DateTime]::Now
  	$endTime = $startTime.AddHours($numberOfHoursInMaintenanceMode)
 
	# ----- Put COmputer Object in Maint Mode 
	if ( $computer.InMaintenanceMode -ne $true ) {
			"Putting " + $computerPrincipalName + " into maintenance mode"
 			try {
					New-MaintenanceWindow -startTime:$startTime -endTime:$endTime -monitoringObject:$computer -comment:$comment -ErrorAction SilentlyContinue
				}
				catch {
					Write-Host "Error:" -ForegroundColor Red
					"$_.Exception"
					$_.FullyQualifiedErrorID
					$_.Exception.gettype().Fullname
					"----"
					$_ | FL *
					"----"
					$computerClass
					$computerCriteria
					$Computer
			}
		}
		else {
			$computerPrincipalName + " Already in maintenance Mode"
	}
 
	# ----- Put Associated Health Service into Maint Mode 
	if ( $HealthService.InMaintenanceMode -ne $true ) {
			try {
					Write-Host "Putting the associated health service into maintenance mode"
					New-MaintenanceWindow -startTime:$startTime -endTime:$endTime -monitoringObject:$healthService -comment:$comment -ErrorAction SilentlyContinue
				}
				catch [System.InvalidOperationException]
					{
						Write-Host "     Associated Health Service already in maintenance Mode"
				}
				catch {
					Write-Host "Error:" -ForegroundColor Red
					"$_.Exception"
					$_.FullyQualifiedErrorID
					$_.Exception.gettype().Fullname
					"----"
					$_ | FL *
			}
		}
		else {
			Write-Host "Associated Health Service already in maintenance Mode"
	}
 
 	# ----- Put Associated health Service Watcher into Maint Mode
 	if ( $HealthServiceWatcher.InMaintenanceMode -ne $true ) {
			try {
					write-host "Putting the associated health service watcher into maintenance mode"
 					New-MaintenanceWindow -startTime:$startTime -endTime:$endTime -monitoringObject:$healthServiceWatcher -comment:$comment -ErrorAction SilentlyContinue
				}
				catch [System.InvalidOperationException]
					{
						Write-Host "     Associated Health Watcher Service already in maintenance Mode"
				}
				catch {
					Write-Host "Error:" -ForegroundColor Red
					"$_.Exception"
					$_.FullyQualifiedErrorID
					$_.Exception.gettype().Fullname
					"----"
					$_ | FL *
			}
		}
		else {
			Write-Host "Associated Health Service Watcher already in maintenance Mode"
	}
}

# ----------------------------------------------------------------------------- 
# Function: Set-RegString 
# Description: Create/Update the specified registry string value 
# Return Value: True/false respectively 
#------------------------------------------------------------------------------

function Set-RegString{ 
    param( 
        [string]$server = ".", 
        [string]$hive, 
        [string]$keyName, 
        [string]$valueName, 
        [string]$value    
    ) 
    $hives = [enum]::getnames([Microsoft.Win32.RegistryHive]) 

    if($hives -notcontains $hive){ 
        write-error "Invalid hive value"; 
        return; 
    } 
    $regHive = [Microsoft.Win32.RegistryHive]$hive; 
    $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regHive,$server); 
    $subKey = $regKey.OpenSubKey($keyName,$true); 

    if(!$subKey){ 
        write-error "The specified registry key does not exist."; 
        return; 
    } 
    $subKey.SetValue($valueName, $value, [Microsoft.Win32.RegistryValueKind]::String); 
    if($?) {$true} else {$false} 
} 

# ------------------------------------------------------------------------------- 
# Function: New-RegSubKey 
# Description: Create the registry key 
# Return Value: True/false respectively 
#--------------------------------------------------------------------------------

function New-RegSubKey{ 
    param( 
        [string]$server = ".", 
        [string]$hive, 
        [string]$keyName 
    ) 

    $hives = [enum]::getnames([Microsoft.Win32.RegistryHive]) 

    if($hives -notcontains $hive){ 
        write-error "Invalid hive value"; 
        return; 
    } 
    $regHive = [Microsoft.Win32.RegistryHive]$hive; 
    $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regHive,$server); 
    [void]$regKey.CreateSubKey($keyName); 
    if($?) {$true} else {$false} 
} 

#-------------------------------------------------------------------------------------
# MAIN
#-------------------------------------------------------------------------------------

if ( $Debug -eq $True ) { 
	Write-Host "Debugging SCOM PS Module..."
	Write-Host "  Management Server --- $ManagementServer"
	Write-Host "  Debug --- $Debug"
}

# ----- Check if Module is already imported


if ( (Get-PSSnapin -Name Microsoft.EnterpriseManagement.OperationsManager.Client -ErrorAction SilentlyContinue) -eq $null ){
	if ( $Debug -eq $True ) { 
		Write-Host "Installing SCOM Snapin..."
	}

	Try {
			try {
					Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client -ErrorAction Stop	
				}
				catch [System.Management.Automation.ActionPreferenceStopException] {
					throw $_.Exception
			}
		}
		catch [System.Management.Automation.Runspaces.PSSnapInException]{
			if ( $Debug -eq $True ) { 
				Write-Host "Error ----> $_" -ForegroundColor red
			}
			Write-Host "need to register snapin first..."
			
			try {
					Write-Host "Registering SCOM Snapin so it can be laoded..."
					# ----- Copy needed files for the SCOM snapin
					MD 'c:\Temp\SCOM' -force
					Copy-Item "\\vbas022.vbgov.com\c$\Program Files\System Center Operations Manager 2007\Microsoft.EnterpriseManagement.OperationsManager.Client*.*" "c:\temp\SCOM" -ErrorAction SilentlyContinue
					Copy-item "\\vbas022\c$\Program Files\System Center Operations Manager 2007\SDK Binaries\*.*" "C:\temp\SCOM" -ErrorAction SilentlyContinue
				}
				catch [System.IO.IOException] {
					Write-Host "Error --> $_" -ForegroundColor Red
				}
				
				Write-Host "Editing Registry"
				# ----- Manually register the Snapin
				New-RegSubKey . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'ApplicationBase' 'c:\temp\SCOM' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'AssemblyName' 'Microsoft.EnterpriseManagement.OperationsManager.ClientShell, Version=6.0.4900.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'ModuleName' 'C:\temp\SCOM\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.dll' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'PowerShellVersion' '1.0' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Vendor' 'Microsoft Corporation' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Version' '6.0.4900.0' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Description' 'Microsoft Operations Manager Shell Snapin' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Types' 'C:\temp\SCOM\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.Types.ps1xml' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Formats' 'C:\temp\SCOM\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.Format.ps1xml' | out-null 
 		}
		
}
get-Module
if ( $Debug -eq $True ) { 
	Write-Host "Mapping SCOM Drive..."
}

# ----- Create a drive that maps to the root of the provider namespace. 
New-PSDrive -Name: Monitoring -PSProvider: OperationsManagerMonitoring -Root: \  -Scope global | Out-Null
cd Monitoring:\
 
# ----- create a connection to the Management Group you intend to manage. 
$ManamgementServer = Get-RootManagementServer 

New-ManagementGroupConnection $ManagementServer | out-null
cd $ManagementServer	

if ( $Debug -eq $True ) { 
	Write-Host "SCOM PS Module load Complete..."
}

#--------------------------------------------------------------------------------------------
# Expose the functions to powershell
#--------------------------------------------------------------------------------------------

Export-ModuleMember Set-SCOMMaintenanceMode


