<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.93
	 Created on:   	9/17/2015 4:39 PM
	 Created by: tcc  	 
	 Organization: vuit	 
	 Filename: dailyquotaincrease    	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ea "SilentlyContinue"
function get-mbxstatreport
{ [cmdletBinding()]
	Param (
		[Parameter(Position = 0, Mandatory = $false, HelpMessage = "The vunet/mailbox to be reported")]
		#[ValidateNotNullOrEmpty()]
		[String]$mailbox,
		[Parameter(Position = 1, Mandatory = $False, HelpMessage = "The file containing multiple mailboxes to be reported")]
		#[ValidateNotNullOrEmpty()]
		[String]$SourceFile,
		[Parameter(Position = 2, Mandatory = $False, HelpMessage = "Email to deliver the report")]
		#[ValidateNotNullOrEmpty()]
		$emailto
	)
		
	function mbxstatreport
	{
		param ([string]$mailbox,
			[string]$emailto)
		$mbx = get-mailbox -id $mailbox | Select-Object  alias, displayname, primarysmtpaddress, ProhibitSendReceiveQuota, UseDatabaseQuotaDefaults
		$mbx | foreach {
			
			#$mailbox = $($_.alias)
			If ($_.UseDatabaseQuotaDefaults -eq $True)
			{
				$PRSQuota = 2048
			}
			Else
			{
				$PRSQuota = $($_.ProhibitSendReceiveQuota)
			}
			$displayname = $($_.displayname)
			
			Write-Host $mailbox, $PRSQuota, $displayname
			
			#$MSGbody=@()
			#$mbx = Get-Mailbox -id $mailbox | Select-Object alias, displayname, primarysmtpaddress,ProhibitSendReceiveQuota
			$mbx_stats = Get-MailboxStatistics -id $mailbox | Select-Object itemcount, totalitemsize
			
			
			If ($_.UseDatabaseQuotaDefaults -eq $True)
			{
				$PercentUsed = ("{0:P1}" -f (($mbx_stats.TotalItemSize.value.toMB()) / ($PRSQuota)))
			}
			Else
			{
				$PercentUsed = ("{0:P1}" -f (($mbx_stats.TotalItemSize.value.toMB()) / ($PRSQuota.value.toMB())))
			}
			
			Write-Host $PercentUsed
			
			$MBX_TotalitemsixeMB = ("{0:n1}" -f $mbx_stats.TotalItemSize.value.toMB())
			$mbx_fstats = Get-MailboxFolderStatistics -id $mailbox | where { $_.ItemsInFolder -gt 10 }`
			| Select-Object Name, FolderPath, ItemsInFolder,`
							@{ n = "Folder Size"; e = { (([math]::round(($_.foldersize.tokb()/1024), 2).tostring()) + " MB") } },
							@{ n = "FolderSize"; e = { (([math]::round(($_.foldersize.tokb()/1024), 2))) } },`
							@{ n = "Folder And Subfolder Size(MB)"; e = { (([math]::round(($_.FolderAndSubfolderSize.tokb()/1024), 2).tostring()) + " MB") } }`
			| Sort-Object FolderSize –Descending
			
			
			# $Output|sort @{expression="StorageLimitStatus";Descending=$true},@{expression="DisplayName";Ascending=$true}|ft -autosize
			#$Grouped = $Output|Group-Object StorageLimitStatus|sort StorageLimitStatus
			
			$head = '  
<style>
BODY{font-family:Verdana;}  
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}  
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black; text-align: Center}  
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: Left}
TD.highlightyellow { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: yellow}
TD.highlightorange { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: orange}
TD.highlightred { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: red}
.center {text-align:center} 
</style>'
			
			$title = "Mailbox Folder Statistics Report for $($mbx.displayname)"
			# Convert to HTML and highlight rows  
			
			$MSGbody = "<H1>Mailbox Folder Statistics Report</H1>"
			$MSGbody += "<p><th>VunetID : $mailbox</th>"
			$MSGbody += "<br><th>Display Name: $displayname</th>"
			$MSGbody += "<br><th>Total Mailbox Size =$($MBX_TotalitemsixeMB) MB</th>"
			$MSGbody += "<br><th>Current Quota : $($PercentUsed) used </th>"
			$MSGbody += "<br><th>Total Item Count = $($mbx_stats.itemcount) items</th></p>"
			$MSGbody += $Grouped | select @{ Name = "Storage Limit Status"; expression = { $_.Name } }, Count | ConvertTo-Html -Head $head -title $title
			$MSGbody += '</p>'
			$MSGbody += $mbx_fstats | Select-Object Name, FolderPath, ItemsInFolder, "Folder Size" | ConvertTo-Html
			
			####$MSGbody | Out-File d:\temp\output.htm
			
			$HtmlBody = $MSGbody
			
			$HtmlBody = $MSGbody | %{
				If ($_ -Match ".*<td>MailboxDisabled</td>.*")
				{
					$_ -Replace "<td>MailboxDisabled", "<td class='highlightred'>MailboxDisabled"
				}
				ElseIf ($_ -Match ".*<td>IssueWarning</td>.*")
				{
					$_ -Replace "<td>IssueWarning", "<td class='highlightyellow'>IssueWarning"
				}
				ElseIf ($_ -Match ".*<td>ProhibitSend</td>.*")
				{
					$_ -Replace "<td>ProhibitSend", "<td class='highlightorange'>ProhibitSend"
				}
				Else { $_ }
			}
			
			return $HtmlBody
		}
	}
	
	$WarningPreference = "SilentlyContinue"
	
	if ($mailbox)
	{
		$htmlbody = mbxstatreport $mailbox
	}
<#if ($SourceFile)
{
	mbxstatreport $mailbox, $emailto
}#>
	if ($emailto -eq "script")
	{ return $htmlbody }
	else
    {
	 $bodystr = $htmlbody | Out-String
     Send-MailMessage -To $emailto -From "messaging.reports@vanderbilt.edu" -Subject $title -body $bodystr -BodyAsHtml -SmtpServer "srfs.vanderbilt.edu"
     #sendmail -sendto $emailto -sendfrom "messaging.reports@vanderbilt.edu" -subject $title  -body $htmlbody 
    }
}

#saving working location of the script
$runfrom = Get-Location
#There is already a script that runs every day and outputs the over 95% quota users to a csv. Lets go get it.
Set-Location -Path "\\vuit-hcwnem197\D$\Admins\Scripts\MBX_Quota"

$todaysfile = get-childitem | where { $_.name -like "Quota_*" } | sort-object creationtime -Descending | Select Name -First 1
#test location
#$todaysfile = get-childitem | where { $_.name -like "test-quota*" } | sort-object creationtime -Descending | Select Name -First 1
$todaysfilename = $todaysfile.name
$destination = $runfrom.tostring() + "\todaysfile.csv"
copy-item $todaysfilename -Destination $destination
Set-Location $runfrom
#import the list
$todaysdata = Import-CSV -Path .\todaysfile.csv
$disabledlist = @()
$underlist = @()
$overlist = @()

#adjust the quotas
foreach ($Userobj in $todaysdata)
{
	Write-Host "Working on ... $($Userobj.VuNetid) - $($Userobj.Quota) " -NoNewline
	[int]$Quota = $Userobj.Quota
	[string]$MAXQuota = "$($Userobj.Quota)MB"
	If ($($Quota) -ge 2048)
	{
		[string]$IWQuota = (($MAXQuota - 300MB)/1MB)
		Write-Host "IW (300)= $($IWQuota)"
	}
	Else
	{
		[string]$IWQuota = [math]::truncate(($MAXQuota * .8)/1MB)
		Write-Host "IW (80%)= $($IWQuota)"
	}
	
	Set-Mailbox -identity $Userobj.VuNetid.Trim() -UseDatabaseQuotaDefaults $false -IssueWarningQuota "$($IWQuota)MB" -ProhibitSendQuota unlimited -ProhibitSendReceiveQuota $MAXQuota -RecoverableItemsWarningQuota 15GB -RecoverableItemsQuota 20GB -Confirm:$false
	$useremail = Get-Mailbox -identity $Userobj.VuNetid.Trim() | select windowsemailaddress | ft -HideTableHeaders | out-string
	#make a list of people under 5GB
	If (($Userobj.Quota -as [int]) -le 5120)
	{ $currentstats = new-object -TypeName PSObject -Property @{
			vunet = $Userobj.VunetID
			quota = $Userobj.Quota
			oldquota = $Userobj.CurrentQuota
			Name = $Userobj.Displayname
			email = $useremail.trim()
		}
		$underlist += $currentstats
	}
	#make a list of people over 5GB
	If (($Userobj.Quota -as [int]) -gt 5120)
	{ $currentstats = new-object -TypeName PSObject -Property @{
			vunet = $Userobj.VunetID
			quota = $Userobj.Quota
			oldquota = $Userobj.CurrentQuota
			Name = $Userobj.Displayname
			email = $useremail.trim()
		}
		$overlist += $currentstats
	}
	
}
#email todays results
$totallist = $underlist + $overlist
$reportintro = "The following mailbox quotas had exceeded 95% and were increased by 1GB:

		"
$reportdata = $totallist | convertto-html -Fragment
$body = ConvertTo-Html -Head $reportintro -Body $reportdata | Out-String

$SmtpFrom = "Messaging.Reports@vanderbilt.edu"
$SmtpTo = "vuit.collaboration@vanderbilt.edu"
#$SmtpTo = "thomas.chandler@vanderbilt.edu"
<#$SmtpCC =#> 
$SmtpSubject = "Daily Quota Adjustment Report"
Send-MailMessage -From $SmtpFrom -To $smtpto -Subject $SmtpSubject -body $body -BodyAsHtml -SmtpServer "srfs.vanderbilt.edu" 
#create Pesasus ticket for all users under 5GB
#cant pass vars to a spawned window. using environment variables
$env:desclist = $underlist | ForEach-Object { $_.vunet + "," + $_.oldquota }
$env:resolvlist = $underlist | ForEach-Object { $_.vunet + "," + $_.quota }

#spawn ps3 window; innoke web request needs at least ps3
powershell -version 3 -windowstyle normal -command { $psversiontable.psversion; .\create-pegasusincident.ps1 -incidenttype under -desclist $env:desclist -resolvlist $env:resolvlist; exit;}
#for EACH user over 5GB - get mbx stat report and make pegasus ticket

foreach ($user in $overlist)
{ #mailbox statistics report
	$report = get-mbxstatreport -mailbox $User.vunet -emailto script
		
		####$MSGbody | Out-File d:\temp\output.htm
		$oldquostr = [math]::round($User.oldquota/1000) | Out-String
		$newquostr = [math]::round($User.quota/1000) | Out-String
		$cleanreq = "Greetings,<br>
<br>
You are receiving this email because your mailbox is at or near its quota. We have proactively adjusted the quota from $oldquostr GB to $newquostr GB. An incident has been opened by us on your behalf. 
The report below can be used by you to aid with cleaning up your mailbox. Please refer any questions to your LAN Manager/Local Support Provider.<br>
<br>
Thanks, VUIT Collaboration Team <br>"

	$subjectname = $user.name	
	$SmtpSubject = "WARNING: Mailbox Cleanup Requested for $subjectname"
    $smtpbody = ConvertTo-Html -Head $cleanreq -Body $report | Out-String
    $smtpto = $user.email
    $smtpfrom = "Vuit.Collaboration@vanderbilt.edu"
	Send-MailMessage -From $SmtpFrom -To $smtpto -Subject $SmtpSubject -body $smtpbody -BodyAsHtml -SmtpServer "srfs.vanderbilt.edu" -Bcc "Vuit.Collaboration@vanderbilt.edu"
    #cant pass vars to a spawned window. using environment variables
	$env:PegasusName = $subjectname
    $env:PegasusVunet = $user.vunet
    $env:PegasusCurQuota = $oldquostr
    $env:PegasusQuota = $newquostr
    powershell -version 3 -windowstyle normal -command { $psversiontable.psversion; .\create-pegasusincident.ps1 -incidenttype over -Name $env:PegasusName -Vunet $env:PegasusVunet -CurrentQuota $env:PegasusCurQuota -Quota $env:PegasusQuota; exit;}
    
	}

#>

