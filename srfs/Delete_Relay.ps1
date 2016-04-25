if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

$Date = (Get-Date).ToString('yyyyMMdd')

# Added to prevent truncation in the results
$FormatEnumerationLimit =-1


$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$filename = "$ScriptDirectory\Relay_IP_Results_$Date.txt"
$IPs_toRemove = Get-Content .\RemoveIPs.txt | select -unique

$RemoteIPs = @()

$Server = gc env:computername
$Conns = @(Get-ReceiveConnector "srfs*" )
ForEach ($Conn in $Conns){
	$RemoteIPs += (Get-ReceiveConnector -Identity $Conn ).RemoteIPRanges |Select -Property *, @{l="Connector";e={$($Conn.name)}} 
}


ForEach ($IP in $IPs_toRemove) {

	$results = @($RemoteIPs -match ($IP))

	if ($results) {
		$RC_Conn = $results[0].Connector

		$Connector = Get-ReceiveConnector -Identity "$($RC_Conn)"
		Write-Host "BEFORE :There are $(($Connector.RemoteIPRanges).count) entries on $($Connector.name)" -ForegroundColor Green
		Write-Host "Removing ... $IP on $RC_Conn" -foregroundcolor green
		$IP | foreach {$Connector.RemoteIPRanges -= "$_"}
		Set-ReceiveConnector "$($RC_Conn)" -RemoteIPRanges $Connector.RemoteIPRanges
		Write-Host "AFTER: There are $(($Connector.RemoteIPRanges).count) entries on $($Connector.name)" -ForegroundColor Green

	}
	Else {
		Write-host "$($IP) was not found on SRFS Connectors" -foregroundcolor red
	}
}
