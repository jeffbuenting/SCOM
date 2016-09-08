@Echo off
REM -- Configure Powershell
C:\Windows\System32\WindowsPowerShell\v1.0\powershell -command "& {set-executionpolicy unrestricted}"

REM -- Complete the rest in Powershell x86
C:\Windows\System32\WindowsPowerShell\v1.0\powershell -executionpolicy bypass -command "& '\\vbgov.com\deploy\Disaster_Recovery\SCOM\Scripts\place-maintenancemode.ps1'

pause