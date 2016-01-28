Param (
	[Parameter(Position=0, Mandatory=$False, HelpMessage="The mailbox to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String] $Mailbox,

	[Parameter(Position=1, Mandatory=$False, HelpMessage="The database to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String] $Database,

	[Parameter(Position=2, Mandatory=$False, HelpMessage="The file containing users to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String] $SourceFile,

	[Parameter(Position=3, Mandatory=$False, HelpMessage="The services to be disabled")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("All","Activesync","MAPI","OWA","POP","IMAP","EWS","EWSMAC")]
	[String] $Services = "All"

)


if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}


Function DisableMailboxServices ($Alias,$Service)
{
	Write-Host "Working on ... " $Alias

	
Switch($Service){
		"Activesync" {set-CASMailbox $Alias -ActiveSyncEnabled:$False}
		"MAPI" {set-CASMailbox $Alias -MAPIEnabled:$False}
		"OWA" {set-CASMailbox $Alias -OWAEnabled:$False }
		"POP" {set-CASMailbox $Alias -PopEnabled:$False }
		"IMAP" {set-CASMailbox $Alias -ImapEnabled:$False }
		"EWS" {set-CASMailbox $Alias -EwsEnabled:$False }
		"EWSMAC" {set-CASMailbox $Alias -EwsAllowEntourage:$False -EwsAllowMacOutlook:$False }

		default {set-CASMailbox $Alias -ActiveSyncEnabled:$False -ImapEnabled:$False -MAPIEnabled:$False -EwsEnabled:$False -OWAEnabled:$False -PopEnabled:$False -EwsAllowEntourage:$False -EwsAllowMacOutlook:$False}
	}
#Report the Results

$Result = get-casmailbox $Alias |select name,ActiveSyncEnabled,MAPIEnabled,OWAEnabled,ImapEnabled,PopEnabled,EwsEnabled,EwsAllowMacOutlook,EwsAllowEntourage
	
Return $Result
}

$dt = Get-Date -format "yyyyMMdd_hhmm"
$CSVFile = ($MyInvocation.MyCommand.Name).Replace(".ps1","_$dt.csv")

$Report = @()

if ($Mailbox) {
	$mbx = get-mailbox -id $Mailbox 
	Foreach($Svc in $Services) { 
		$Report += DisableMailboxServices ($mbx.Alias) $Svc
		get-casmailbox $mbx.Alias |select name,ActiveSyncEnabled,MAPIEnabled,OWAEnabled,ImapEnabled,PopEnabled,EwsEnabled,EwsAllowMacOutlook,EwsAllowEntourage|fl
	}
	 
}
if ($Database) {
	$MBXs = @( get-mailbox -database $Database )
	Foreach($mbx in $MBXs){
		Foreach($Svc in $Services) { 
			$Report += DisableMailboxServices $mbx.Alias $Svc
			get-casmailbox $mbx.Alias |select name,ActiveSyncEnabled,MAPIEnabled,OWAEnabled,ImapEnabled,PopEnabled,EwsEnabled,EwsAllowMacOutlook,EwsAllowEntourage|fl
					
		}
	}
}

If ($SourceFile)
{
	# Read the users to move from a CSV file containing the alias of the users in a column named 'VUNetID'
	Import-Csv $SourceFile | foreach { 
		Foreach($Svc in $Services) { 
			$Report += DisableMailboxServices $_.VUNetID $Svc
			get-casmailbox $_.VUNetID |select name,ActiveSyncEnabled,MAPIEnabled,OWAEnabled,ImapEnabled,PopEnabled,EwsEnabled,EwsAllowMacOutlook,EwsAllowEntourage|fl

		}
	}
}

$Report|select -property *|Export-Csv -NoTypeInformation -Path .\$CSVFile  