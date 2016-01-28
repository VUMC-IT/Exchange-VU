<#
.SYNOPSIS
   <A brief description of the script>
.DESCRIPTION
   <A detailed description of the script>
.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
#>


$WarningPreference = "SilentlyContinue"

#Begin customization-------------------------
$SmtpServer = "srfs.vanderbilt.edu" #Enter FQDN of SMTP server
$SmtpFrom = "Messaging.Reports@vanderbilt.edu" #Enter sender email address
#$SmtpTo = "Tony.Hortert@vanderbilt.edu" #Enter one or more recipient addresses in an array
$SmtpTo = "e.zafar@vanderbilt.edu","taj.wolff@vanderbilt.edu"
$SmtpCC = "ECS@vanderbilt.edu" #Enter one or more recipient addresses in an array
$SmtpSubject = “Vmail Mailbox Quota Report” #Enter subject of message
#End customization---------------------------

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto

}

$Output = @()
$MailDBs =@()

$OutputFile_Path = "D:\Admins\Scripts\MBX_Quota\"
$TrackingOutputFile_Path = "D:\Admins\Scripts\Mailbox Folder Statistics\"
$dt = Get-Date -format "yyyyMMdd_hhmm"

$mailboxStats = get-mailboxserver |get-MailboxStatistics | where {"IssueWarning","ProhibitSend","MailboxDisabled" -contains $_.StorageLimitStatus -and $_.ObjectClass –eq “Mailbox”}

if ($mailboxStats -ne $null) {

	foreach ($mailboxStat in $mailboxStats) {

		$mailbox = Get-Mailbox -Identity $mailboxStat -WarningAction SilentlyContinue -Verbose:$false

		If ($Mailbox -ne $null) {
		if ($mailboxStat -ne $null) {
			$tisInBytes = $mailboxStat.TotalItemSize.Value.ToBytes()
		} else {
			$tisInBytes = "N/A"
		}

		$mbx = New-Object System.Management.Automation.PSObject
		$mbx | Add-Member -type noteProperty -name TotalItemSize -value $mailboxStat.TotalItemSize 
		$mbx | Add-Member -type noteProperty -name ItemCount -value $mailboxStat.ItemCount 
		$mbx | Add-Member -type noteProperty -name StorageLimitStatus -value $mailboxStat.StorageLimitStatus 

		$mbx | Add-Member -type noteProperty -name VUNetID -value $mailbox.Alias
		$mbx | Add-Member -type noteProperty -name DisplayName -value $mailbox.DisplayName
		$mbx | Add-Member -type noteProperty -name ServerName -value $mailbox.ServerName 
		$mbx | Add-Member -type noteProperty -name Department -value (get-user $mailbox.alias).department
		If ($mailbox.UseDatabaseQuotaDefaults -eq $true) 
		{
			$MailDBs = Get-MailboxDatabase $mailbox.Database |select identity,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota
			$mbx | Add-Member -type noteProperty -name ProhibitSendQuota -value $MailDBs.ProhibitSendQuota 
			$mbx | Add-Member -type noteProperty -PassThru -name ProhibitSendReceiveQuota -value $MailDBs.ProhibitSendReceiveQuota 
			$mbx | Add-Member -type noteProperty -name IssueWarningQuota -value $MailDBs.IssueWarningQuota 
			if (-not $MailDBs.ProhibitSendReceiveQuota.IsUnlimited) {
				$psrqInBytes = $MailDBs.ProhibitSendReceiveQuota.Value.ToBytes()
			} else {
				$psrqInBytes = "Unlimited"
			}
			
			}
		Else
		{
			$mbx | Add-Member -type noteProperty -name ProhibitSendQuota -value $mailbox.ProhibitSendQuota 
			$mbx | Add-Member -type noteProperty -PassThru -name ProhibitSendReceiveQuota -value $mailbox.ProhibitSendReceiveQuota 
			$mbx | Add-Member -type noteProperty -name IssueWarningQuota -value $mailbox.IssueWarningQuota 
			if (-not $mailbox.ProhibitSendReceiveQuota.IsUnlimited) {
				$psrqInBytes = $mailbox.ProhibitSendReceiveQuota.Value.ToBytes()
			} else {
				$psrqInBytes = "Unlimited"
			}
			
		}
		$mbx | Add-Member -type noteProperty -name UseDatabaseQuotaDefaults -value $mailbox.UseDatabaseQuotaDefaults 



		#Calculate mailbox usage report
		if (($tisInBytes -ne "N/A") -and ($psrqInBytes -ne "Unlimited")) {
			$usagePercent = '{0:P2}' -f ($tisInBytes / $psrqInBytes)
		} else {
			$usagePercent = "N/A"
		}
		$mbx | Add-Member -type noteProperty -name PercentUsed -value $usagePercent


		$Output += $mbx
	}
	}
}


# $Output|sort @{expression="StorageLimitStatus";Descending=$true},@{expression="DisplayName";Ascending=$true}|ft -autosize
$Grouped = $Output|Group-Object StorageLimitStatus|sort StorageLimitStatus


# Export those mailboxes greater than 95%
#$Output|where{$_.PercentUsed -ge 95 -and $_.ProhibitSendReceiveQuota -lt 4294967296}|Select VUNetID,@{Name="Quota";expression={($_.ProhibitSendReceiveQuota.Value.toMB()  + 1024)}},DisplayName,Department,@{Name="CurrentQuota";expression={($_.ProhibitSendReceiveQuota.Value.toMB())}},PercentUsed |Sort VUNetID |Export-CSV ($OutputFile_Path + "Quota_" + $dt +".csv") -notypeinformation
$Output|where{$_.PercentUsed -ge 95}|Select VUNetID,@{Name="Quota";expression={($_.ProhibitSendReceiveQuota.Value.toMB()  + 1024)}},DisplayName,Department,@{Name="CurrentQuota";expression={($_.ProhibitSendReceiveQuota.Value.toMB())}},PercentUsed |Sort VUNetID |Export-CSV ($OutputFile_Path + "Quota_" + $dt +".csv") -notypeinformation
$Output|where{$_.PercentUsed -ge 95}|Select VUNetID | Export-CSV ($TrackingOutputFile_Path + "Tracking_" + $dt +".csv") -notypeinformation


$head = ' 
<style> 
BODY{font-family:Verdana;} 
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;} 
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black; text-align: Center} 
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right}
TD.highlightyellow { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: yellow}
TD.highlightorange { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: orange}
TD.highlightred { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: red}
.center {text-align:center} 
</style>' 

$title = "VMail Mailbox Quota Report" 
# Convert to HTML and highlight rows  
# @{Name=””;expression={$_.}}
$MSGbody = "<H1>VMail Mailbox Quota Report</H1><H2>SUMMARY</H2>"
$MSGbody += $Grouped | Select @{Name="Storage Limit Status";expression={$_.Name}}, Count |ConvertTo-Html -Head $head -title $title
$MSGbody += '</p><H2>DETAIL</H2>' 
$MSGbody += $Output|sort @{expression="StorageLimitStatus";Descending=$true},@{expression="PercentUsed";Descending=$true}|`
 	select VUNetID,@{Name=”Display Name”;expression={$_.DisplayName}},Department,ItemCount,@{Name="Total Size (MB)";expression={$_.TotalItemSize.Value.ToMB()}},`
 	@{Name="Storage Limit Status";expression={$_.StorageLimitStatus}},`
 	@{Name="Use Database Quota Defaults";expression={$_.UseDatabaseQuotaDefaults}},`
 	@{Name="Issue Warning Quota (MB)";expression={if ($_.IssueWarningQuota.IsUnlimited -ne $true){$_.IssueWarningQuota.Value.ToMB()}Else{$_.IssueWarningQuota}}},`
 	@{Name="Prohibit Send Quota (MB)";expression={if ($_.ProhibitSendQuota.IsUnlimited -ne $true){$_.ProhibitSendQuota.Value.ToMB()}Else{$_.ProhibitSendQuota}}},`
 	@{Name="Prohibit Send Receive Quota(MB)";expression={if ($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){$_.ProhibitSendReceiveQuota.Value.ToMB()}Else{$_.ProhibitSendReceiveQuota}}}, `
 	@{Name="`% Used";expression={$_.PercentUsed}}|ConvertTo-Html
#$MSGbody += $Output|ConvertTo-Html

$HtmlBody =$MSGbody | %{ 
	If ($_ -Match ".*<td>MailboxDisabled</td>.*") { 
		$_ -Replace "<td>MailboxDisabled", "<td class='highlightred'>MailboxDisabled" 
		$script:bAlert = $true
	} 
	ElseIf ($_ -Match ".*<td>IssueWarning</td>.*") { 
		$_ -Replace "<td>IssueWarning", "<td class='highlightyellow'>IssueWarning" 
	} 
	ElseIf ($_ -Match ".*<td>ProhibitSend</td>.*") { 
		$_ -Replace "<td>ProhibitSend", "<td class='highlightorange'>ProhibitSend" 
		$script:bAlert = $true
	} 
	Else { $_ }
}

If ($script:bAlert -eq $true)
{	$SmtpSubject += " (Attention Required)"
	$SmtpPriority = 'High'}
Else
{
	$SmtpPriority = 'Normal'}

Send-MailMessage -From $SmtpFrom -To $SmtpTo -CC $SmtpCC -Subject $SmtpSubject -Body ($HtmlBody|out-string) -BodyAsHTML -SMTPserver $SmtpServer -DeliveryNotificationOption onFailure -Priority $SmtpPriority



