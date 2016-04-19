Add-PSSnapin Quest.ActiveRoles.ADManagement -ea "SilentlyContinue"

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

function sendmail($body)
{
	$SmtpClient = new-object system.net.mail.smtpClient 
	$MailMessage = New-Object system.net.mail.mailmessage 
	$SmtpClient.Host = "srfs.vanderbilt.edu" 
	$mailmessage.from = "Messaging.Reports@vanderbilt.edu" 
	$mailmessage.To.add("VUIT.Incident.Response@Vanderbilt.edu")
	#$mailmessage.To.add("ITSITIRReport@intemail.email.vanderbilt.edu") 
	$mailmessage.To.add("ECS@vanderbilt.edu") 
	$mailmessage.Subject = “ITIR Quota Report” 
	$MailMessage.IsBodyHtml = $True
	$mailmessage.Body = $body

	$smtpclient.Send($mailmessage) 
}

$results = @()
$ReturnData = @()
$Database = @{n="Database";e={ $stats.database }} 
$DisplayName = @{n="DisplayName";e={ $stats.Displayname }}
$StorageLimitStatus = @{n="StorageLimitStatus";e={ $stats.StorageLimitStatus }}
$ItemCount = @{n="ItemCount";e={ $stats.ItemCount }}
$TotalItemSize = @{n="TotalItemSize(MB)";e={ $stats.TotalItemSize.Value.ToMB() }}
 
 $results += get-mailbox 'ITIR*' | foreach {
 $stats = get-mailboxstatistics $_
 $_ | select @{name="VUNet";expression={$_.Alias}},$DisplayName,$ItemCount,$TotalItemSize,ProhibitSendReceiveQuota,$StorageLimitStatus,@{n="Percent Used";e={"{0:P2}" -f ($stats.TotalItemSize.Value.ToMB() /$_.ProhibitSendReceiveQuota.Value.ToMB())}}
 }
$results += get-mailbox 'ITSI' | foreach {
 $stats = get-mailboxstatistics $_
 $_ | select @{name="VUNet";expression={$_.Alias}},$DisplayName,$ItemCount,$TotalItemSize,ProhibitSendReceiveQuota,$StorageLimitStatus,@{n="Percent Used";e={"{0:P2}" -f ($stats.TotalItemSize.Value.ToMB() /$_.ProhibitSendReceiveQuota.Value.ToMB())}}
} 
$results += get-mailbox 'swe0*' | foreach {
 $stats = get-mailboxstatistics $_
 $_ | select @{name="VUNet";expression={$_.Alias}},$DisplayName,$ItemCount,$TotalItemSize,ProhibitSendReceiveQuota,$StorageLimitStatus,@{n="Percent Used";e={"{0:P2}" -f ($stats.TotalItemSize.Value.ToMB() /$_.ProhibitSendReceiveQuota.Value.ToMB())}}
} 

$results += get-mailbox 'swe1*' | foreach {
 $stats = get-mailboxstatistics $_
 $_ | select @{name="VUNet";expression={$_.Alias}},$DisplayName,$ItemCount,$TotalItemSize,ProhibitSendReceiveQuota,$StorageLimitStatus,@{n="Percent Used";e={"{0:P2}" -f ($stats.TotalItemSize.Value.ToMB() /$_.ProhibitSendReceiveQuota.Value.ToMB())}}
} 

# Build the Email
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
$header = "<H1>ITIR Quota Report</H1>"  
$title = "ITIR Quota Report"  

$body = $results|select VUNet,DisplayName,ItemCount,"TotalItemSize(MB)",StorageLimitStatus,"Percent Used"| Sort-Object "Percent Used" -Descending 
# Convert to HTML and highlight rows
$HtmlBody = $body | ConvertTo-Html -Head $head -body $header -title $title| %{ 
  If ($_ -Match ".*<td>IssueWarning</td>.*") { 
    $_ -Replace "<td>", "<td class='highlightyellow'>" 
  } 
  ElseIf ($_ -Match ".*<td>MailboxDisabled</td>.*") { 
    $_ -Replace "<td>", "<td class='highlightred'>" 
  } 
  ElseIf ($_ -Match ".*<td>ProhibitSend</td>.*") { 
    $_ -Replace "<td>", "<td class='highlightorange'>" 
  } 
  Else { $_ }
}
# Send the Email 
Sendmail($HtmlBody)