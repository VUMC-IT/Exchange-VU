#Get the list of users to be migrated from this file
$Users = @(Get-Content -Path  .\Remove_quota.csv)

foreach($User in $Users)
{
Set-Mailbox -identity $User.Trim() -UseDatabaseQuotaDefaults $false -IssueWarningQuota unlimited -ProhibitSendQuota unlimited -ProhibitSendReceiveQuota unlimited -Confirm:$false 
}
