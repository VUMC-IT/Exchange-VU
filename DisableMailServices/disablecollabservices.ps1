<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.92
	 Created on:   	8/25/2015 8:31 AM
	 Created by:   	guy shepperd / Thomas Chandler
	 Organization: 	VUIT
	 Filename:     	disablecollabservices.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
Param (
	[Parameter(Position = 0, Mandatory = $False, HelpMessage = "The mailbox to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String]$Mailbox,
	[Parameter(Position = 1, Mandatory = $False, HelpMessage = "The database to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String]$Database,
	[Parameter(Position = 2, Mandatory = $False, HelpMessage = "The file containing users to have mail services disabled")]
	[ValidateNotNullOrEmpty()]
	[String]$SourceFile,
	[Parameter(Position = 3, Mandatory = $False, HelpMessage = "The services to be disabled")]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("All", "Activesync", "MAPI", "OWA", "POP", "IMAP", "EWS", "EWSMAC", "LYNC")]
	[String]$Services = "All"
	
)
#######This block is for communication w the test lab#######
#$cred = get-credential -Message "Login with username@vanderbilt.edu"
#session information for implicit remote to CAS
#$exsess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://vuit-hcwnem101q.ds-test.vandytestlab.com/PowerShell/ -Authentication Kerberos -Credential $cred
#session information/options for implicit remote to Lync pool
#$lyncOptions = New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck
#$lyncsess = New-PSSession -ConnectionUri https://hcfepool2013.vanderbilt.edu/ocsPowerShell -SessionOption $lyncOptions -Authentication NegotiateWithImplicitCredential
#Import exchange commands
#Import-PSSession -Session $exsess
#Import Lync commands
#Import-PSSession -Session $lyncsess
############################################################
############Tests access to logging location################
$outfile = "\\groups-nas2.its.vanderbilt.edu\IARC\Logs\PSadmin\test.txt"
try
{
	Get-Date -format d | Add-Content $outfile
}

Catch
{
	Write-Warning "Unable to access logging location...Exiting"
	Start-Sleep -s 5
	Break
}
############Establishes implicit session w production env#####
$cred = get-credential -Message "Login with username@vanderbilt.edu"
#session information for implicit remote to CAS
$exsess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://vuit-hcwnem133/PowerShell/ -Authentication Kerberos -Credential $cred
#session information/options for implicit remote to Lync pool
$lyncOptions = New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck
$lyncsess = New-PSSession -ConnectionUri https://hcfepool2013.vanderbilt.edu/ocsPowerShell -SessionOption $lyncOptions -Authentication NegotiateWithImplicitCredential
#Import exchange commands
Import-PSSession -Session $exsess
#Import Lync commands
Import-PSSession -Session $lyncsess
##########################################################

Function DisableMailboxServices ($Alias, $Service)
{
	Write-Host "Working on ... " $Alias
	
	
	Switch ($Service)
	{
		"Activesync" { set-CASMailbox $Alias -ActiveSyncEnabled:$False }
		"MAPI" { set-CASMailbox $Alias -MAPIEnabled:$False }
		"OWA" { set-CASMailbox $Alias -OWAEnabled:$False }
		"POP" { set-CASMailbox $Alias -PopEnabled:$False }
		"IMAP" { set-CASMailbox $Alias -ImapEnabled:$False }
		"EWS" { set-CASMailbox $Alias -EwsEnabled:$False }
		"EWSMAC" { set-CASMailbox $Alias -EwsAllowEntourage:$False -EwsAllowMacOutlook:$False }
		"LYNC" {
				revoke-CSClientCertificate -identity $Alias
				disable-csuser -identity $Alias
				}
		
		default
		{
			set-CASMailbox $Alias -ActiveSyncEnabled:$False -ImapEnabled:$False -MAPIEnabled:$False -EwsEnabled:$False -OWAEnabled:$False -PopEnabled:$False -EwsAllowEntourage:$False -EwsAllowMacOutlook:$False
			revoke-CSClientCertificate -identity $Alias
			disable-csuser -identity $Alias
		}
	}
	#Report the Results
	
	$mailresult = get-casmailbox $alias
	$lyncresult = get-csuser $alias
	$result = new-object -TypeName PSObject -Property @{
		name = $mailresult.name
		ActiveSyncEnabled = $mailresult.ActivesyncEnabled
		MAPIEnabled = $mailresult.MAPIEnabled
		OWAEnabled = $mailresult.OWAEnabled
		ImapEnabled = $mailresult.ImapEnabled
		PopEnabled = $mailresult.PopEnabled
		EwsEnabled = $mailresult.EWSEnabled
		EwsAllowMacOutlook = $mailresult.EWSAllowMacOutlook
		EwsAllowEntourage = $mailresult.EWSAllowEntourage
		LyncEnabled = $lyncresult.Enabled
	}
	return $result | select name, ActiveSyncEnabled, MAPIEnabled, OWAEnabled, ImapEnabled, PopEnabled, EwsEnabled, EwsAllowMacOutlook, EwsAllowEntourage, LyncEnabled
}

$dt = Get-Date -format "yyyyMMdd_hhmm"
$CSVFile = ($MyInvocation.MyCommand.Name).Replace(".ps1", "_$dt.csv")
$admin = [Environment]::UserName
$reason = read-host "Enter Incident or short justification"
$Report = @()

if ($Mailbox)
{
	$mbx = get-mailbox -id $Mailbox
	Foreach ($Svc in $Services)
	{
		$rslt = DisableMailboxServices ($mbx.Alias) $Svc
		$Report += $rslt
		$rslt
	}
	
}
if ($Database)
{
	$MBXs = @(get-mailbox -database $Database)
	Foreach ($mbx in $MBXs)
	{
		Foreach ($Svc in $Services)
		{
			$rslt = DisableMailboxServices ($mbx.Alias) $Svc
			$Report += $rslt
			$rslt
		}
	}
}

If ($SourceFile)
{
	# Read the users to move from a CSV file containing the alias of the users in a column named 'VUNetID'
	Import-Csv $SourceFile | foreach {
		Foreach ($Svc in $Services)
		{
			$rslt = DisableMailboxServices $_.VUNetID $Svc
			$Report += $rslt
			$rslt
		}
	}
}
$Report | Add-member -Name Admin -Value $admin -MemberType NoteProperty -PassThru |
Add-member -Name Reason -Value $reason -MemberType NoteProperty -PassThru |
select Admin, reason, name, ActiveSyncEnabled, MAPIEnabled, OWAEnabled, ImapEnabled, PopEnabled, EwsEnabled, EwsAllowMacOutlook, EwsAllowEntourage, LyncEnabled |
Export-Csv -NoTypeInformation -Path "\\groups-nas2.its.vanderbilt.edu\IARC\Logs\PSadmin\$CSVFile"
