#Get the list of users to be migrated from this file
$Users = @(Import-Csv .\quota.csv)

foreach($User in $Users)
{
Write-Host "Working on ... $($User.VuNetid) - $($User.Quota) " -NoNewline 
[int]$Quota = $User.Quota
[string]$MAXQuota = "$($User.Quota)MB" 
If ($($Quota) -ge 2048) { 
	[string]$IWQuota = (($MAXQuota-300MB)/1MB)
	Write-Host "IW (300)= $($IWQuota)" 
}
Else {
	[string]$IWQuota = [math]::truncate(($MAXQuota * .8)/1MB)
	Write-Host "IW (80%)= $($IWQuota)" 
}

Set-Mailbox -identity $User.VuNetid.Trim() -UseDatabaseQuotaDefaults $false -IssueWarningQuota "$($IWQuota)MB" -ProhibitSendQuota unlimited -ProhibitSendReceiveQuota $MAXQuota -RecoverableItemsWarningQuota 15GB -RecoverableItemsQuota 20GB  -Confirm:$false 
}
