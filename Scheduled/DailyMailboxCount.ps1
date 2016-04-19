#modified 5/28/13 by Chare
#removed Exchange 2007 Servers and code for calculating 2007 data

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

$GTCount = 0
$AllCount = @()
$MBCount = @()

$MBCount = @("")
$Environment = "Exchange 2010 Mailbox Count"
$serverlist = ("ITS-HCWNEM101","ITS-HCWNEM102","ITS-HCWNEM103","ITS-HCWNEM104","ITS-HCWNEM105","ITS-HCWNEM106","ITS-HCWNEM107","ITS-HCWNEM108")
$Totalmailboxcount = 0
Foreach ($Server in $Serverlist)
{	$Mailboxcount = (get-mailbox -server $Server -ResultSize Unlimited | Measure-Object)
	$Mailboxcount = $Mailboxcount.Count
	$Totalmailboxcount += $Mailboxcount
	$MBCount = New-Object PSObject
	$MBCount | Add-Member NoteProperty -Name "Server" -Value $Server
	$MBCount | Add-Member NoteProperty -Name "Count" -Value $Mailboxcount
	$AllCount += $MBCount
}
$MBCount = New-Object PSObject
$MBCount | Add-Member NoteProperty -Name "Server" -Value $Environment
$MBCount | Add-Member NoteProperty -Name "Count" -Value $Totalmailboxcount
$AllCount += $MBCount
$GTCount += $Totalmailboxcount



$MBCount = New-Object PSObject
$MBCount | Add-Member NoteProperty -Name "Server" -Value "Grand Total"
$MBCount | Add-Member NoteProperty -Name "Count" -Value $GTCount

$AllCount += $MBCount

$AllCount | format-table -autosize

  $head = '  
<style>  
BODY{font-family:Verdana;}  
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}  
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black; text-align: Center}  
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right}
.center {text-align:center} 
</style>' 
$header = "<H1>Exchange Mailbox Count Report</H1>"  
$title = "Exchange Mailbox Count Report"  
  
$msg =$AllCount |Select-Object Server,Count|ConvertTo-Html -head $head -body $header -title $title

function sendmail($body)
{
    $SmtpClient = new-object system.net.mail.smtpClient 
    $MailMessage = New-Object system.net.mail.mailmessage 
    $SmtpClient.Host = "srfs.vanderbilt.edu" 
    $mailmessage.from = "Messaging.Reports@vanderbilt.edu" 
    #$mailmessage.To.add("mark.gossard@vanderbilt.edu")
	$mailmessage.To.add("ECS@vanderbilt.edu")
	#$CC = "catherine.a.crimi@Vanderbilt.Edu","philip.neely@vanderbilt.edu"
 
    $CC | ForEach {$MailMessage.CC.add($_)}
	$mailmessage.Subject = “Exchange Mailbox Count Report” 
    $MailMessage.IsBodyHtml = $True
    $mailmessage.Body = $body
    
    $smtpclient.Send($mailmessage) 
}
sendmail $msg

#Write-Host "Done!"