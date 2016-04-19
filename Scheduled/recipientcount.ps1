# ==============================================================================================
# 
# Microsoft PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2009
# 
# NAME: 
# 
# AUTHOR: user , 
# DATE  : 8/13/2010
# 
# COMMENT: 
# 
# ==============================================================================================

$WarningPreference = "SilentlyContinue"

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}
    
if (-not (Get-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue))
	{Add-PSSnapin Quest.ActiveRoles.ADManagement}	
  

Function Convert-BytesToSize ($Size)
{

#Decide what is the type of size
Switch ($Size)
{
{$Size -gt 1PB}
{
#Write-Verbose “Convert to GB”
$NewSize = “$([math]::Round(($Size / 1PB),2))PB”
Break
}
{$Size -gt 1TB}
{
#Write-Verbose “Convert to GB”
$NewSize = “$([math]::Round(($Size / 1TB),2))TB”
Break
}
{$Size -gt 1GB}
{
#Write-Verbose “Convert to GB”
$NewSize = “$([math]::Round(($Size / 1GB),2))GB”
Break
}
{$Size -gt 1MB}
{
#Write-Verbose “Convert to MB”
$NewSize = “$([math]::Round(($Size / 1MB),2))MB”
Break
}
{$Size -gt 1KB}
{
#Write-Verbose “Convert to KB”
$NewSize = “$([math]::Round(($Size / 1KB),2))KB”
Break
}
Default
{
#Write-Verbose “Convert to Bytes”
$NewSize = “$([math]::Round($Size,2))Bytes”
Break
}
}
Return $NewSize

}

#Make $LogFile match the Powershell script name with a date and time stamp followed by ".log"
$dt = Get-Date -format "yyyyMMdd_hhmm"
$RawDataFile = ($MyInvocation.MyCommand.Name).Replace(".ps1","_RawData_$dt.log")
$ReportDataFile = ($MyInvocation.MyCommand.Name).Replace(".ps1","_Report_$dt.csv")

$All_Mailboxes = @()
 
$From = "Messaging.Reports@vanderbilt.edu"
$Title = "Vmail Details for $(Get-Date -Format d)"

Get-Mailbox -ResultSize Unlimited| foreach{ 	
	
$MBX   = $_ | select Name,Alias,ServerName,RecipientTypeDetails,ProhibitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults,IssueWarningQuota,ExchangeVersion,WhenMailboxCreated,CustomAttribute* 
$MBX | Add-Member -type noteProperty -name Division -value (get-qaduser $_.alias -IncludedProperties division -DontUseDefaultIncludedProperties).division
$MBX | Add-Member -type noteProperty -name Department -value (get-user $_.alias).department
	
if ($_.ProhibitSendReceiveQuota.IsUnlimited){
                $MBX  | Add-Member -type noteProperty -name RealQuota -value 2048 
		    }
            else {
                $MBX  | Add-Member -type noteProperty -name RealQuota -value $_.ProhibitSendReceiveQuota.value.toMB() 
            }
	Get-MailboxStatistics $_ | foreach{ 
	$MBX | add-member -type noteProperty -name DisplayName -value $_.DisplayName
	$MBX | add-member -type noteProperty -name ItemCount -value $_.ItemCount
	$MBX | add-member -type noteProperty -name StorageLimitStatus -value $_.StorageLimitStatus
	$MBX | add-member -type noteProperty -name TotalItemSize -value $_.TotalItemSize.value
	$MBX | add-member -type noteProperty -name TotalItemSize_MB -value $_.TotalItemSize.value.toMB()
	$MBX | add-member -type noteProperty -name TotalItemSize_GB -value $_.TotalItemSize.value.toGB()
	$MBX | add-member -type noteProperty -name DeletedItemCount -value $_.DeletedItemCount
	$MBX | add-member -type noteProperty -name TotalDeletedItemSize -value $_.TotalDeletedItemSize
	$MBX | add-member -type noteProperty -name DatabaseName -value $_.DatabaseName 
    $MBX | add-member -type noteProperty -name TotalMailboxUsage -value $($_.TotalDeletedItemSize.value + $_.TotalItemSize.value)
	$MBX | add-member -type noteProperty -name DisconnectDate -value $_.DisconnectDate 
	$MBX | add-member -type noteProperty -name DisconnectReason -value $_.DisconnectReason 
	}
#$MBX
$All_Mailboxes += $MBX
}

$Resource_MBX    = $All_Mailboxes | where { $_.RecipientTypeDetails -match "RoomMailbox" -or $_.RecipientTypeDetails -match "EquipmentMailbox" }
$Shared_MBX       = $All_Mailboxes | where { $_.RecipientTypeDetails -match "SharedMailbox" }
$User_MBX        = $All_Mailboxes | where { $_.RecipientTypeDetails -match "UserMailbox" }
$Created_Last7Days = $All_Mailboxes | where {$_.WhenMailboxCreated -ge (Get-Date).AddDays(-7)}



$Exchange2010_MBX = $All_Mailboxes | where {($_.ExchangeVersion.ExchangeBuild.Major -eq 14)}
$Exchange2007_MBX = $All_Mailboxes | where {($_.ExchangeVersion.ExchangeBuild.Major -eq 8)}



$DGs = (Get-DistributionGroup -resultsize unlimited).Count
 
$All_Ex_Servers = Get-Exchangeserver|sort name
$All_EX2007_Servers = $All_Ex_Servers | where {($_.AdminDisplayVersion -like "Version 8.*")}|sort name
$All_EX2010_Servers = $All_Ex_Servers | where {($_.AdminDisplayVersion -like "Version 14.*")}|sort name

# Handle the Edge Counts since they are not joined to the domain
#$2007Edges = New-Object System.Object
#$2007Edges | Add-Member -type NoteProperty -name Name -value "Edge"
#$2007Edges | Add-Member -type NoteProperty -name Count -value 2
# Handle the Edge Counts since they are not joined to the domain
$2010Edges = New-Object System.Object
$2010Edges | Add-Member -type NoteProperty -name Name -value "Edge"
$2010Edges | Add-Member -type NoteProperty -name Count -value 3


$All_EX2007_Roles = $All_EX2007_Servers|group-object serverrole|select name,count
$All_EX2007_Roles += $2007Edges 

$All_EX2010_Roles = $All_EX2010_Servers|group-object serverrole|select name,count
$All_EX2010_Roles += $2010Edges 


$AllMBX_Size = $All_Mailboxes|Measure-Object TotalItemSize -ave -min -max -sum
$AllMBX_Item = $All_Mailboxes|Measure-Object ItemCount -ave -min -max -sum
$AllMBX_Quota = $All_Mailboxes|where{$_.UseDatabaseQuotaDefaults -eq $false}|Measure-Object ProhibitSendReceiveQuota -ave -min -max -sum

$TotalQuota = $All_Mailboxes|Measure-Object RealQuota -ave -min -max -sum


$UserMBX_Size = $User_MBX|Measure-Object TotalItemSize -ave -min -max -sum
$UserMBX_Item = $User_MBX|Measure-Object ItemCount -ave -min -max -sum



$All_DBs = Get-MailboxDatabase|where {$_.recovery -eq $false}|sort server,name


$AllMBX_Quota

$LT_512MB_MBX = @()
$GT_512MB_LT_1GB_MBX = @()
$GT_1GB_LT_2GB_MBX = @()
$GT_2GB_LT_5GB_MBX = @()
$GT_5GB_LT_10GB_MBX = @()
$GT_10GB_MBX = @()

foreach ($Mbox in $All_Mailboxes) {

if ($Mbox.TotalItemSize -le 536870912) {
	$LT_512MB++
	$LT_512MB_MBX += $Mbox
}
elseif ($Mbox.TotalItemSize -gt 536870912 -and $Mbox.TotalItemSize -le 1073741824){
	$GT_512MB_LT_1GB++
	$GT_512MB_LT_1GB_MBX += $Mbox
}
elseif ($Mbox.TotalItemSize -gt 1073741824 -and $Mbox.TotalItemSize -le 2147483648){
	$GT_1GB_LT_2GB++
	$GT_1GB_LT_2GB_MBX += $Mbox
}
elseif ($Mbox.TotalItemSize -gt 2147483648 -and $Mbox.TotalItemSize -le 5368709120){
	$GT_2GB_LT_5GB++
	$GT_2GB_LT_5GB_MBX += $Mbox
}
elseif ($Mbox.TotalItemSize -gt 5368709120 -and $Mbox.TotalItemSize -le 10737418240){
	$GT_5GB_LT_10GB++
	$GT_5GB_LT_10GB_MBX += $Mbox
}
elseif ($Mbox.TotalItemSize -gt 10737418240){
	$GT_10GB++
	$GT_10GB_MBX += $Mbox
}

}

$LT_512MB_STAT = $LT_512MB_MBX |Measure-Object TotalItemSize -ave -min -max -sum
$GT_512MB_LT_1GB_STAT = $GT_512MB_LT_1GB_MBX |Measure-Object TotalItemSize -ave -min -max -sum
$GT_1GB_LT_2GB_STAT = $GT_1GB_LT_2GB_MBX |Measure-Object TotalItemSize -ave -min -max -sum
$GT_2GB_LT_5GB_STAT = $GT_2GB_LT_5GB_MBX |Measure-Object TotalItemSize -ave -min -max -sum
$GT_5GB_LT_10GB_STAT = $GT_5GB_LT_10GB_MBX |Measure-Object TotalItemSize -ave -min -max -sum
$GT_10GB_STAT = $GT_10GB_MBX |Measure-Object TotalItemSize -ave -min -max -sum


# Export Data to CSV

$All_Mailboxes|select -property * |export-csv -path .\$($RawDataFile) -notypeinformation

"EXServers,TotalMBXCount,UserMBXCount,SharedMBXCount,ResourceMBXCount,DistGrpCount,UserMBXAVGSize,UserMBXAVGItem,UserMBXMinSize,UserMBXMinItem,UserMBXMaxSize,UserMBXMaxItem,ALLMBXAVGSize,ALLMBXAVGItem,ALLMBXMinSize,ALLMBXMinItem,ALLMBXMaxSize,ALLMBXMaxItem,MBXSize512,MBXSize1GB,MBXSize2GB,MBXSize5GB,MBXSize10GB,MBXSizeGT10GB"| out-file $ReportDataFile -encoding ASCII

"$($All_Ex_Servers.Count),$($All_Mailboxes.Count),$($User_MBX.Count),$($Shared_MBX.Count),$($Resource_MBX.Count),$($DGs),$($UserMBX_Size.Average),$($UserMBX_Item.Average),$($UserMBX_Size.Minimum),$($UserMBX_Item.Minimum),$($UserMBX_Size.Maximum),$($UserMBX_Item.Maximum),$($AllMBX_Size.Average),$($AllMBX_Item.Average),$($AllMBX_Size.Minimum),$($AllMBX_Item.Minimum),$($AllMBX_Size.Maximum),$($AllMBX_Item.Maximum),$($LT_512MB),$($GT_512MB_LT_1GB),$($GT_1GB_LT_2GB),$($GT_2GB_LT_5GB),$($GT_5GB_LT_10GB),$($GT_10GB)"| out-file $ReportDataFile -append -encoding ASCII



# Build the Email Message
$msg  = "<html> <head>   
<style>  
BODY{font-family:Verdana;}  
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}  
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black; text-align: Center}  
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right}
.center {text-align:center}

li{background-repeat: no-repeat;background-position: 100% .4em;padding-right: .6em;margin: 1em 0;} 
</style> </head><body><H1>$Title</H1> " 

$msg  += "<ul>"
$msg  += "<li>Number of Exchange Servers in AD: $($All_Ex_Servers.Count)"
$msg  += "<ul>"
$msg  += "<li>Exchange 2007: $([int]$($All_EX2007_Servers.Count))</li>"
$msg  += $All_EX2007_Roles|ConvertTo-Html -Fragment 
$msg  += "<li>Exchange 2010: $([int]$($All_EX2010_Servers.Count)+5) </li>"
$msg  += $All_EX2010_Roles|ConvertTo-Html -Fragment
$msg  += "</ul></li></p>"
$msg  += "<li>Total Mailbox Count: $($All_Mailboxes.Count)"
$msg  += "<ul>"
$msg  += "<li>User Mailboxes: $($User_MBX.Count)</li>" 
$msg  += "<li>Shared Mailboxes: $($Shared_MBX.Count) </li>"
$msg  += "<li>Resource Mailboxes: $($Resource_MBX.Count) </li>"
$msg  += "</ul></li></p>"

$msg  += "Exchange 2007 MBX: $($Exchange2007_MBX.Count)"
$msg  += "<br>"
$msg  += "Exchange 2010 MBX: $($Exchange2010_MBX.Count)"
$msg  += "<br>"

$msg  += "MBX Created in Last 7 Days: $($Created_Last7Days.Count)"
$msg  += "<br>"

$msg  += "<li>Distribution Groups: $($DGs) </li></ul>"

$msg  += "Total Quota: $([math]::Round($TotalQuota.Sum/1TB,2))"
$msg  += "<br>"
$msg  += "</p>"

$msg  += "<table>"
$msg  += "<tr>"
$msg  += "<td>User Mailboxes</td>"
$msg  += "<td>Size</td>"
$msg  += "<td>Item</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Average</td>"
$msg  += "<td> $(Convert-BytesToSize $UserMBX_Size.Average) </td>"
$msg  += "<td> $([math]::Round($UserMBX_Item.Average,0)) </td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Minimum</td>"
$msg  += "<td> $(Convert-BytesToSize $UserMBX_Size.Minimum) </td>"
$msg  += "<td> $([math]::Round($UserMBX_Item.Minimum,0)) </td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Maximum</td>"
$msg  += "<td>$(Convert-BytesToSize $UserMBX_Size.Maximum)</td>"
$msg  += "<td>$([math]::Round($UserMBX_Item.Maximum,0))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Sum</td>"
$msg  += "<td>$(Convert-BytesToSize $UserMBX_Size.Sum)</td>"
$msg  += "<td>" + [string]::format('{0:N0}',$([math]::Round($UserMBX_Item.Sum,0))) +"</td>"
$msg  += "</tr>"

$msg  += "</table>"

$msg  += "<br>"

$msg  += "<table>"
$msg  += "<tr>"
$msg  += "<td>All Mailboxes</td>"
$msg  += "<td>Size</td>"
$msg  += "<td>Item</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Average</td>"
$msg  += "<td> $(Convert-BytesToSize $AllMBX_Size.Average) </td>"
$msg  += "<td> $([math]::Round($AllMBX_Item.Average,0)) </td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Minimum</td>"
$msg  += "<td> $(Convert-BytesToSize $AllMBX_Size.Minimum) </td>"
$msg  += "<td> $([math]::Round($AllMBX_Item.Minimum,0)) </td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Maximum</td>"
$msg  += "<td>$(Convert-BytesToSize $AllMBX_Size.Maximum)</td>"
$msg  += "<td>$([math]::Round($AllMBX_Item.Maximum,0))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Sum</td>"
$msg  += "<td>$(Convert-BytesToSize $AllMBX_Size.Sum)</td>"
$msg  += "<td>" + [string]::format('{0:N0}',$([math]::Round($AllMBX_Item.Sum,0))) + "</td>"
$msg  += "</tr>"

$msg  += "</table>"

$msg  += "<br>"

$msg  += "<table>"
$msg  += "<tr>"
$msg  += "<td>Mailbox Size</td>"
$msg  += "<td>Mailbox Count</td>"
$msg  += "<td>Percent</td>"
$msg  += "<td>Storage (GB)</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Less Than 512MB</td>"
$msg  += "<td> $($LT_512MB) </td>"
$percent1 = "{0:P2}" -f $($LT_512MB/$All_Mailboxes.Count)
$msg  += "<td>$percent1</td>"
$msg  += "<td>$([math]::Round($LT_512MB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Less Than 1GB</td>"
$msg  += "<td> $($GT_512MB_LT_1GB) </td>"
$percent2 = "{0:P2}" -f $($GT_512MB_LT_1GB/$All_Mailboxes.Count)
$msg  += "<td>$percent2</td>"
$msg  += "<td>$([math]::Round($GT_512MB_LT_1GB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Less than 2GB</td>"
$msg  += "<td> $($GT_1GB_LT_2GB) </td>"
$percent3 = "{0:P2}" -f $($GT_1GB_LT_2GB/$All_Mailboxes.Count)
$msg  += "<td>$percent3</td>"
$msg  += "<td>$([math]::Round($GT_1GB_LT_2GB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Less than 5GB</td>"
$msg  += "<td>$($GT_2GB_LT_5GB)</td>"
$percent4 = "{0:P2}" -f $($GT_2GB_LT_5GB/$All_Mailboxes.Count)
$msg  += "<td>$percent4</td>"
$msg  += "<td>$([math]::Round($GT_2GB_LT_5GB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Less than 10GB</td>"
$msg  += "<td>$($GT_5GB_LT_10GB)</td>"
$percent5 = "{0:P2}" -f $($GT_5GB_LT_10GB/$All_Mailboxes.Count)
$msg  += "<td>$percent5</td>"
$msg  += "<td>$([math]::Round($GT_5GB_LT_10GB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"

$msg  += "<tr>"
$msg  += "<td>Greater than 10GB</td>"
$msg  += "<td>$($GT_10GB)</td>"
$percent6 = "{0:P2}" -f $($GT_10GB/$All_Mailboxes.Count)
$msg  += "<td>$percent6</td>"
$msg  += "<td>$([math]::Round($GT_10GB_STAT.Sum/1GB,2))</td>"
$msg  += "</tr>"


$msg  += "</table>"


$msg  += "</body></html>"
  
function sendmail($body)
{
    $SmtpClient = new-object system.net.mail.smtpClient 
    $MailMessage = New-Object system.net.mail.mailmessage 
    $SmtpClient.Host = "srfs.vanderbilt.edu" 
    $mailmessage.from = $From 
    $mailmessage.To.add("mark.gossard@vanderbilt.edu")
    $mailmessage.To.add("chris.hare@vanderbilt.edu")
    $mailmessage.To.add("guy.shepperd@vanderbilt.edu")	
    #$mailmessage.To.add("ECS@vanderbilt.edu")
	#$CC = "ExchangeConsolidation@vanderbilt.edu","philip.neely@vanderbilt.edu","john.d.osborne@vanderbilt.edu"
 
    #$CC | ForEach {$MailMessage.CC.add($_)}
	$mailmessage.Subject = $Title 
    $MailMessage.IsBodyHtml = $True
    $mailmessage.Body = $body
    
    $smtpclient.Send($mailmessage) 
}

sendmail $msg