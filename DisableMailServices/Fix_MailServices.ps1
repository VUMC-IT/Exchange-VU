Param (
	[Parameter(Position=0, Mandatory=$False,ValueFromPipeline=$true, HelpMessage="The mailbox to have an OOO message")]
	[ValidateNotNullOrEmpty()]
	[String] $Mailbox,

	[Parameter(Position=1, Mandatory=$False, HelpMessage="The file containing users to have OOO messages processed")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript({Test-Path $_ -PathType 'Leaf'})] 
	[String] $SourceFile
)

if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

$dt = Get-Date -format "yyyyMMdd_hhmm"
$CSVFile = ($MyInvocation.MyCommand.Name).Replace(".ps1","_$dt.csv")

$Report = @()


If ($Mailbox)
{
Write-Host "Working on ... " $Mailbox
	 
		set-CASMailbox $Mailbox -ActiveSyncEnabled:$True -ImapEnabled:$True -MAPIEnabled:$True -EwsEnabled:$True -OWAEnabled:$True -PopEnabled:$True -EwsAllowEntourage:$True -EwsAllowMacOutlook:$True -EwsAllowOutlook:$True
	
}

If ($SourceFile)
{
Write-Host "Working on ... " $_.VUNetID
	Import-Csv $SourceFile | foreach { 
		set-CASMailbox $_.VUNetID -ActiveSyncEnabled:$True -ImapEnabled:$True -MAPIEnabled:$True -EwsEnabled:$True -OWAEnabled:$True -PopEnabled:$True -EwsAllowEntourage:$True -EwsAllowMacOutlook:$True -EwsAllowOutlook:$True
	}
}