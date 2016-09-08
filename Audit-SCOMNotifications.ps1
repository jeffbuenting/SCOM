#------------------------------------------------------------------------------
# Scom Powershell script to audit Subscription Notfications
#------------------------------------------------------------------------------

Param(	$AlertName )
#		$Resolution,
#		$Source, 
#		$Path,
#		$LastModBy,
#		$LastModTime,
#		$Description,
#	$NotificationGUID )

#-----------------------------------------------------------------------------------------
# Function Setup-SCOMPowershellEnvironment 
#
# Sets up Powershell to use SCOM comandlets.
# http://derekhar.blogspot.com/2007/07/operation-manager-command-shell-on-any.html
# http://blogs.msdn.com/b/scshell/archive/2007/01/03/have-your-powershell-and-our-cmdlets-too.aspx
# ----------------------------------------------------------------------------------------

Function Setup-SCOMPowershellEnvironment {

	#----- Set up SCOM Environment
	if ( (Get-PSSnapin -Name Microsoft.EnterpriseManagement.OperationsManager.Client -ErrorAction SilentlyContinue) -eq $null )	{
		if ( Test-Path -Path "C:\Program Files\System Center Operations Manager 2007" ) {	# ------ Checking if SCOM Console is installed
				Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client
			}
			else {
				# ----- Copy needed files for the SCOM snapin
				Copy-Item "\\vbas022.vbgov.com\c$\Program Files\System Center Operations Manager 2007\Microsoft.EnterpriseManagement.OperationsManager.Client*.*" "c:\temp"
				Copy-item "\\vbas022\c$\Program Files\System Center Operations Manager 2007\SDK Binaries\*.*" "C:\temp"
				
				# ----- Manually register the Snapin
				New-RegSubKey . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'ApplicationBase' 'c:\temp' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'AssemblyName' 'Microsoft.EnterpriseManagement.OperationsManager.ClientShell, Version=6.0.4900.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'ModuleName' 'C:\temp\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.dll' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'PowerShellVersion' '1.0' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Vendor' 'Microsoft Corporation' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Version' '6.0.4900.0' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Description' 'Microsoft Operations Manager Shell Snapin' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Types' 'C:\temp\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.Types.ps1xml' | out-null
				Set-RegString . 'LocalMachine' 'SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.EnterpriseManagement.OperationsManager.Client' 'Formats' 'C:\temp\Microsoft.EnterpriseManagement.OperationsManager.ClientShell.Format.ps1xml' | out-null 
				
			    Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client 
		}
	}
	
	# ----- Create a drive that maps to the root of the provider namespace. 
	New-PSDrive -Name: Monitoring -PSProvider: OperationsManagerMonitoring -Root: \ -Scope script | out-null
	cd Monitoring:\
	 
	# ----- create a connection to the Management Group you intend to manage. 
	New-ManagementGroupConnection vbas022  | out-null
	cd vbas022	 
}


#-----------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------



Write-Host "Audit-SCOMNotification.ps1"
Write-Host "AlertName is: ",$AlertName
Write-eventlog -logname 'Operations Manager' -source 'Health Service Script' -eventID 9990 -Entrytype Information -message "Audit-SCOMNotification.ps1 Starting`n`nParameters:`n  AlertName: $AlertName`n`n"


#
#
#Setup-SCOMPowershellEnvironment
#
#$Subscription = Get-NotificationSubscription -Id $NotificationGUID
#
#$Today = Get-Date -Format yyyy-MMM-dd
#
#$ExcelFile = "C:\scomlogs\notifications"+$Today+".xlsx"
#
#
#
## ----- File exist ( C:\scomlogs\notifications$Today.log )
#
#if ( Test-Path -Path $ExcelFile ) {
#		# ----- Get the Spreadsheet
#		$AuditTrail = New-Object -ComObject excel.application
#		
#		$WorkBook = $AuditTrail.WorkBooks.Open( $ExcelFile )
#		$Worksheet = $WorkBook.Worksheets.Item(1)
#		
#		# ----- Find next empty row
#		$R = 1
#		do {$R++}until( $Worksheet.Cells.Item($R,1).Text -eq "" )
#	}
#	Else {
#		# ----- Create the file
#		$AuditTrail = New-Object -ComObject excel.application
#				
#		$WorkBook=$AuditTrail.WorkBooks.Add()
#		$WorkSheet = $WorkBook.Worksheets.Item(1)
#		
#		$WorkSheet.Saveas( $ExcelFile )
#						
#		# ----- Empty row is first row
#		$R = 1
#}
#
## ----- Add info to the spreadsheet
#
#$Worksheet.Cells.Item( $R,1 ) = (Get-Date)
#$Worksheet.Cells.Item( $R,2 ) = $SendTo
#$Worksheet.Cells.Item( $R,3 ) = $AlertName
#$Worksheet.Cells.Item( $R,4 ) = $Resolution
#$Worksheet.Cells.Item( $R,5 ) = $Source
#$Worksheet.Cells.Item( $R,6 ) = $lastModBy
#$Worksheet.Cells.Item( $R,7 ) = $LastModTime
#$Worksheet.Cells.Item( $R,8 ) = $Description
#
#
# #----- Save the Spreadsheet
#
#$Workbook.Save( )
#$AuditTrail.Quit()
#
## ----- Cleanup old spreadsheets older than one week
#
#Get-ChildItem c:\scomlogs | where { $_.lastwritetiem -lt (date).adddays(-7) } | remove-item


#   $Data/Context/DataItem/ResolutionStateName$ $Data/Context/DataItem/ManagedEntityPath$\$Data/Context/DataItem/ManagedEntityDisplayName$  $Data/Context/DataItem/ManagedEntityPath$ $Data/Context/DataItem/LastModifiedBy$ $Data/Context/DataItem/LastModifiedLocal$ $Data/Context/DataItem/AlertDescription$ $Data/Context/DataItem/AlertId$

#  $Data/Context/DataItem/ResolutionStateName$ $Data/Context/DataItem/ManagedEntityPath$\$Data/Context/DataItem/ManagedEntityDisplayName$  $Data/Context/DataItem/ManagedEntityPath$ $Data/Context/DataItem/LastModifiedBy$ $Data/Context/DataItem/LastModifiedLocal$ $Data/Context/DataItem/AlertDescription$ $Data/Context/DataItem/AlertId$
