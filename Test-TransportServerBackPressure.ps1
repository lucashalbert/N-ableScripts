<#
.SYNOPSIS
Test-TransportServerBackPressure.ps1 - Script to check Transport Servers for Back Pressure.
.DESCRIPTION 
Checks the event logs of Hub Transport servers for "back pressure" events.
.OUTPUTS
Results are output to the PowerShell window.
.PARAMETER server
Perform a check of a single server
.PARAMETER hours
Specify the number of hours to look back in the event log
.EXAMPLE
.\Test-TransportServerBackPressure.ps1
Checks all Hub Transport servers in the organization and outputs the results to the shell window.
.EXAMPLE
.\Test-TransportServerBackPressure.ps1 -server HO-EX2010-MB1
Checks the server HO-EX2010-MB1 and outputs the results to the shell window.
.EXAMPLE
.\Test-TransportServerBackPressure.ps1 -server HO-EX2010-MB1 -hours 24
Checks the server HO-EX2010-MB1 for the previous 24 hours and outputs the results to the shell window.
.LINK
http://exchangeserverpro.com/powershell-script-check-hub-transport-servers-for-back-pressure-events
.NOTES
Written by: Paul Cunningham
Find me on:
* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp
Modified by: Lucas Halbert
Date: 02/14/2019
Changes: Add the 'hours' parameter to allow specification of hours to look back in the event log.
         Add state output.
Find me on:
* Website:  https://www.lhalbert.xyz/
* Twitter:  https://twitter.com/lucashalbert
* LinkedIn: https://www.linkedin.com/in/lucashalbert/
* Github:   https://github.com/lucashalbert
#>

#requires -version 2

Param(
	[Parameter( Mandatory=$false)]
	[string]$server,
    [Parameter( Mandatory=$false)]
    [int]$hours
)

# Add default log time if not set
if (!($hours))
{
    $hours = 72
}

#Add Exchange 2010 snapin if not already loaded
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}

#...................................
# Script
#...................................


#Check if a single server was specified
if ($server)
{
	#Run for single specified server
	try
	{
		[array]$servers = @(Get-ExchangeServer $server -ErrorAction Stop)
	}
	catch
	{
		#Couldn't find Exchange server of that name
		Write-Warning $_.Exception.Message
		
		#Exit because single server was specified and couldn't be found in the organization
		EXIT
	}
}
else
{
	#Get list of Hub Transport servers in the organization
	[array]$servers = @(Get-ExchangeServer | Where-Object {$_.IsHubTransportServer})
}

#Check each server
foreach($server in $servers)
{
    # Calculate Date
    $Begin = (Get-Date).AddHours(-$hours)

	#$events = @(Invoke-Command –Computername $server –ScriptBlock { Get-EventLog -LogName Application | Where-Object {$_.Source -eq "MSExchangeTransport" -and $_.Category -eq "ResourceManager"} })
    # Add begin and end log search time
    #$events = @(Invoke-Command –Computername $server –ScriptBlock { param($Begin) Get-EventLog -After $Begin -Before (Get-Date) -LogName Application | Where-Object {$_.Source -eq "MSExchangeTransport" -and $_.Category -eq "ResourceManager"} } -ArgumentList $Begin)
	$events = @(Get-EventLog -After $Begin -Before (Get-Date) -LogName Application | Where-Object {$_.Source -eq "MSExchangeTransport" -and $_.Category -eq "ResourceManager"})
    $count = $events.count

	if ($count -lt 1)
	{
		$Output = "OK: $server has no back pressure events found."
        Write-Host $Output
        $State = "0"
	}
	else
	{
		$lastevent = $events | Select-Object -First 1

		$now = Get-Date
		$timewritten = $lastevent.TimeWritten
		$ago = "{0:N0}" -f ($now - $timewritten).TotalHours 
		
		switch ($lastevent.EventID)
		{
			"15006"
            {
                $BPstate = "Critical (Diskspace)"
                $State = "1"
            }
			"15007"
            {
                $BPstate = "Critical (Memory)"
                $CState = "1"
            }
			default
            {
                $BPstate = $lastevent.ReplacementStrings[1]
                #Unknown status
                $State = "2"
            }
		}

		$Output = "$server is $BPstate as of $ago hours ago"
        Write-Host $Output
        Exit

	}
}
