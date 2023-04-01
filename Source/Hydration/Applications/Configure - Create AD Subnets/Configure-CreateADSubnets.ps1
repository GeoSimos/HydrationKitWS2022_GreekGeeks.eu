<#
Solution: Hydration
Purpose: Used to create AD Sites and Subnets
Version: 1.2 - January 10, 2013

This script is provided "AS IS" with no warranties, confers no rights and 
is not supported by the authors or Deployment Artist. 

Author - Johan Arwidmark
    Twitter: @jarwidmark
    Blog   : http://deploymentresearch.com
#>


# Determine where to do the logging 
$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
$logPath = $tsenv.Value("LogPath") 
$logFile = "$logPath\$($myInvocation.MyCommand).log" 

# Start the logging 
Start-Transcript $logFile 
Write-Host "Logging to $logFile" 

# Create Empty AD Sites (sites with no domain controllers, for lab purpose only)
New-ADReplicationSite -Name Cologne
New-ADReplicationSite -Name Minneapolis

# Create AD Subnets 
New-ADReplicationSubnet -Name "10.0.1.0/24" -Site Thessaloniki
New-ADReplicationSubnet -Name "10.0.2.0/24" -Site Cologne
New-ADReplicationSubnet -Name "10.0.3.0/24" -Site Minneapolis

# Stop logging 
Stop-Transcript

