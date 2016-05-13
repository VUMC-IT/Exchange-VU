<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.120
	 Created on:   	5/11/2016 3:10 PM
	 Created by:   	chanct 
	 Organization: 	 
	 Filename: blockspammer.ps1     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

try { get-exchangeserver | Out-Null }
catch
{
	$exsess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://vuit-hcwnem133/PowerShell/ -Authentication Kerberos
	Import-PSSession -Session $exsess
}

$user = Read-Host -Prompt "Type Vunet of user to block"

add-adgroupmember -identity "vmail_internalblocked" -members $user -WhatIf
Write-Host ("Adding " + $user + " to internal block list`r`n") -ForegroundColor Yellow
add-adgroupmember -identity "vmail_externalblocked" -members $user -WhatIf
Write-Host ("Adding " + $user + " to external block list`r`n") -ForegroundColor Yellow
New-MailboxRepairRequest -Mailbox $user -corruptiontype ProvisionedFolder, SearchFolder, AggregateCounts, Folderview -whatif
Write-Host ("Starting mailbox repair on " + $user + ". This will sever all existing client connections`r`n") -ForegroundColor Yellow
Write-Host ("The newly spawned window will contain the inbox rules for " + $user + ". Please examine for suspicious rules - IE Send all mail to Deleted Items. 
    Ctrl-click any rules to be removed and press OK`r`n") -ForegroundColor Yellow
$smtp = Get-mailbox $user | select primarysmtpaddress
$smtp = $smtp.primarysmtpaddress
$badrules = get-inboxrule -mailbox $smtp | select ruleidentity, description | Out-GridView -passthru
$badrules.RuleIdentity | remove-inboxrule -mailbox $user -whatif
Write-Host ("Removed malicious rules from " + $user + "`r`n") -ForegroundColor Yellow
Write-Host ("Remember to log the block in sharepoint: 
    https://vanderbilt365.sharepoint.com/sites/VUIT/Hosting_services/Collaboration/_layouts/15/start.aspx#/Lists/BlockedEmailTracking/Active%20Email%20Blocks.aspx") -ForegroundColor Yellow
