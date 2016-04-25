 # Load EX2010 Snapin
 if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

$ConnectorID = "SRFS-Connector4"

# Added to prevent truncation in the results
$FormatEnumerationLimit =-1

$Connector = Get-ReceiveConnector -Identity $ConnectorID
Write-Host "BEFORE :There are $(($Connector.RemoteIPRanges).count) entries on $($Connector.name)" -ForegroundColor Green

$IPs_toADD = Get-Content .\IPs.txt | select -unique
Write-Host "There are $($IPs_toADD.count) entries to be added" -ForegroundColor Yellow

$IPs_toADD| foreach {$Connector.RemoteIPRanges += $_}

Set-ReceiveConnector $ConnectorID -RemoteIPRanges $Connector.RemoteIPRanges

$Connector = Get-ReceiveConnector -Identity $ConnectorID

Write-Host "AFTER: There are $(($Connector.RemoteIPRanges).count) entries on $($Connector.name)" -ForegroundColor Green