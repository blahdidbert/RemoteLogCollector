#
#  This script exports consolidated and filtered event logs to CSV
#  Author: Michael Karsyan, FSPro Labs, eventlogxp.com (c) 2016
#  Modified for use by: Bryan Bowie, bryanbowie.info 2016
#  Version 3.0
# Added more comments to the sections
# Added all lognames line 20
# Added more to Eventtypes line 23
# Added the creation of Test folder if it doesn't already exist line 30
# Added -Message -UserName -TimeGenerated -TimeWritten to line 51


#Setting the variable 'EventAgeDays' to the number of day's worth of logs to obtain; here 7 days.
Set-Variable -Name EventAgeDays -Value 7

# Setting variables to the machine name you are wanting logs from. Repalce SERV(N).
Set-Variable -Name CompArr -Value @("DESKTOP-A7CLLQH")

# Setting the variable to check for application and systems logs.
# Need to run "Get-EventLog -list" on machine and correct below names if needed.
Set-Variable -Name LogNames -Value @("Application", "Security", "System", "HardwareEvents", "Internet Explorer", "Key Management Service", "OAlerts", "PreEmptive", "Windows PowerShell")

# Setting the variable to catch all log types.
Set-Variable -Name EventTypes -Value @("Error", "Warning", "Information", "FailureAudit", "SuccessAudit")

# Setting variable to export folder location. Change as necessary. 
# This folder path is is local to the machine this script is run from.
Set-Variable -Name ExportFolder -Value "C:\TEST\"

#Create Export folder and suppress exception if folder exists.
mkdir $ExportFolder -Force

# Consolidated error log into one location.
$el_c = @()
#Get today's date. 
$now=get-date
#Set date to today and subtract EventAgeDays.
$startdate=$now.adddays(-$EventAgeDays)

# Since Windows freaks out over ":" using dashes instead.
$ExportFile=$ExportFolder + "el" + $now.ToString("yyyy-MM-dd---hh-mm-ss") + ".csv"

foreach($comp in $CompArr)
{
  foreach($log in $LogNames)
  {
    Write-Host Processing $comp\$log
    $el = get-eventlog -ComputerName $comp -log $log -After $startdate -EntryType $EventTypes 
    $el_c += $el  # Consolidating information.
  }
}
$el_sorted = $el_c | Sort-Object TimeGenerated    # Sort by time.

#Exporting to CSV
Write-Host Exporting to $ExportFile
$el_sorted|Select EntryType, TimeGenerated, TimeWritten, Source, EventID, MachineName, Message, UserName| Export-CSV $ExportFile -NoTypeInfo  
Write-Host Done!