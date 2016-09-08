#-----------------------------------------------------------------------------------------
# Capture Alerts CPU Utilization  
#
#-----------------------------------------------------------------------------------------



#---------------------------------------------------------------------------------------
# MAIN
#---------------------------------------------------------------------------------------

#----- Set up SCOM Environment
Add-PSSnapin Microsoft.EnterpriseManagement.OperationsManager.Client -ErrorAction SilentlyContinue
# ----- Create a drive that maps to the root of the provider namespace. 
New-PSDrive -Name: Monitoring -PSProvider: OperationsManagerMonitoring -Root: \ | Out-Null
cd Monitoring:\
 
# ----- create a connection to the Management Group you intend to manage. 
New-ManagementGroupConnection vbas022 | out-null
cd vbas022	

#Get Current Date and Time
$Date = Get-Date
$TimeBiasInMinutes = ($Date - $Date.ToUniversalTime()).TotalMinutes
$Date = $Date.ToShortDateString()
$Date = $Date -replace("/", "-")

#Set the Start date and time
#StartDate is the previous business day from today
$PreviousBusinessDay = Get-Date
DO{
	$PreviousBusinessDay = $PreviousBusinessDay.AddDays(-1)
} until ($PreviousBusinessDay.DayofWeek -inotmatch "saturday" -AND $PreviousBusinessDay.DayofWeek -inotmatch "sunday")

$StartDate = Get-Date -Date $PreviousBusinessDay -Hour 18 -Minute 0 -Second 0
$StartDate = $StartDate.ToUniversalTime()

$StartDate

#Set the end date and time
$EndDate = Get-Date -Date $Date -Hour 7 -Minute 30 -Second 0
$EndDate = $EndDate.ToUniversalTime()

#If End time is in the future, use current date and time instead
if ($EndDate -gt $Date) {$EndDate = $Date}

$EndDate

$Alerts = Get-Alert | Where-Object {$_.Name -eq "Total CPU Utilization Percentage is too high" -and ($_.TimeRaised -ge $StartDate -and $_.TimeRaised -le $EndDate)}

#$Alerts | Select-Object name,principalname,parameters

$P=@()
foreach ( $A in $Alerts ) {
	$A.parameters
}

$P