$Conn_Prefix ="SRFS-Connector"
$LastConn_Num = [int](Get-ReceiveConnector "SRFS*" |Where {$_.name -ne "SRFS-Connector-Throttled"} |sort name -desc|select -first 1).tostring().split("\")[1].substring(14)
$New_Conn_Name = "$Conn_Prefix$($LastConn_Num+1)"
$title = $Null

$prompt = "Create New SRFS Connector named: $($New_Conn_Name)?"

$Yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes','Continue'

$No = New-Object System.Management.Automation.Host.ChoiceDescription '&No','Exit'

$options = [System.Management.Automation.Host.ChoiceDescription[]] ($Yes,$No)


$choice = $host.ui.PromptForChoice($title,$prompt,$options,1)
If ($choice -eq 0){
	# Create new SRFS Receive Connector
	New-ReceiveConnector -Name $($New_Conn_Name) -Usage Internal -RemoteIPRanges 127.0.0.4 
	Set-ReceiveConnector -Identity $($New_Conn_Name) -MaxMessageSize 30720KB -ProtocolLoggingLevel Verbose -MessageRateLimit 600 -fqdn "srfs.vanderbilt.edu" -MaxInboundConnection 15000 -PermissionGroups "AnonymousUsers,ExchangeServers" -AuthMechanism "Tls,ExternalAuthoritative"
	# Set Anonymous Permissions to Accept any Recipients
	Get-ReceiveConnector "$($New_Conn_Name)"| Add-ADPermission -User "NT AUTHORITY\ANONYMOUS LOGON" -ExtendedRights "Ms-Exch-SMTP-Accept-Any-Recipient"
	
	Write-Host "Verify $($New_Conn_Name) Connector Permissions" -ForegroundColor yellow
	Get-ReceiveConnector "$($New_Conn_Name)" |get-adpermission | select user,extendedrights
}