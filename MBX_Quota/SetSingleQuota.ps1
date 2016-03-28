<#

Created by: Jeremy Pratt
Creation date: 03/27/16
Last edited: 03/28/16

This script does the following: 
1. Prompts for a single VUnetID
2. Returns the current quota and % used for that VUnetID
3. Prompts for new quota, in GB
4. Sets the new quota
5. Returns the new quota and % used for that user

#>

Function GetQuota
{
    Param (
	   [Parameter(Position=0, Mandatory=$False, HelpMessage="The mailbox you need to check")]
	   [ValidateNotNullOrEmpty()]
	   [String] $Mailbox
    )

    #Saves mailbox to variable
    $mbx = get-mailbox -id $Mailbox |Select-Object  alias, displayname, primarysmtpaddress,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults 

    
    $mbx | foreach {
    
        #Gets Prohibit Send/Receive Quota
        If ($_.UseDatabaseQuotaDefaults -eq $True){
            $PRSQuota = 2048
        }
        Else{
            $PRSQuota = $($_.ProhibitSendReceiveQuota)
        }

        #Writes mailbox quota stats to screen
        $displayname = $($_.displayname)
        Write-Host $mailbox, $PRSQuota, $displayname

        #Saves mailbox statistics to variable
        $mbx_stats = Get-MailboxStatistics -id $mailbox | Select-Object itemcount,totalitemsize

        #Does math for % quota used
        If ($_.UseDatabaseQuotaDefaults -eq $True){
            $PercentUsed =("{0:P1}" -f (($mbx_stats.TotalItemSize.value.toMB()) / ($PRSQuota)))
        }
        Else {
	       $PercentUsed =("{0:P1}" -f (($mbx_stats.TotalItemSize.value.toMB()) / ($PRSQuota.value.toMB())))
        }

        #Writes percentage
        Write-Host $PercentUsed
    }
}

Function SetQuota
{
    #Convert quota from GB to MB.
    $NewQ = ($NewQuota * 1024)

    #Convert to string for parameter
    [string]$MaxQuota = "$($NewQ)MB"
    
    #Determine Issue Warning Quota
    If ($NewQuota -lt 2){
        $IWQuota = ($NewQ * .8)
        $IWQuotaR = [math]::Round($IWQuota)
        [String]$IWQuotaS = "$($IWQuotaR)MB"
    }
    Else {
        $IWQuota = ($NewQ - 300)
        [string]$IWQuotaS = "$($IWQuota)MB"
    }
    
    #Update Quota.
    Write-Host ""
    Write-Host "Setting new quota for user" $($User)
    Write-Host "New Quota = " $MaxQuota
    Write-Host "New Warning Threshold = " $IWQuotaS
    Write-Host ""
    Set-Mailbox -identity $User.Trim() -UseDatabaseQuotaDefaults $false -IssueWarningQuota $IWQuotaS -ProhibitSendQuota unlimited -ProhibitSendReceiveQuota $MaxQuota -RecoverableItemsWarningQuota 15GB -RecoverableItemsQuota 20GB
}
   
#Prompt for user to be changed.
$User = Read-Host "Enter VUnetID"

#Print current quota
Write-Host ""
GetQuota -Mailbox $User

#Prompt for new quota
Write-Host ""
[int]$NewQuota = Read-Host "Enter New Quota (in GB)"


#Run Function to set quota
SetQuota

#Print new quota
GetQuota -Mailbox $User
Write-Host ""
