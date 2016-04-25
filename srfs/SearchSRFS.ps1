if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

$Date = (Get-Date).ToString('yyyyMMdd')


$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$filename = "$ScriptDirectory\Relay_IP_Results_$Date.txt"

$Search_IP = read-host "Enter the IP to search for"

$RemoteIPs = @()

$Server = gc env:computername
$Conns = @(Get-ReceiveConnector "Srfs*" )
ForEach ($Conn in $Conns){
	$RemoteIPs += (Get-ReceiveConnector -Identity $Conn ).RemoteIPRanges |Select -Property *, @{l="Connector";e={$Conn.name}} 
}

$results1 = $RemoteIPs |Where {$_.lowerbound -like $Search_IP.trim()}

$results = $RemoteIPs -match $Search_IP.trim()

if ($results) {
Write-host "$($Search_IP.trim()) was found on `"$($results1.Connector)`"" -foregroundcolor green
}
Else {
Write-host "$($Search_IP.trim()) was not found on SRFS Connectors" -foregroundcolor red
}