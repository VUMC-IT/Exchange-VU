if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}
    

$dt = Get-Date -format "yyyyMMdd_hhmm"
$OutputCSVFile = ($MyInvocation.MyCommand.Name).Replace(".ps1","_$dt.csv")



$DevicesToRemove = Get-ActiveSyncDevice -result unlimited | Get-ActiveSyncDeviceStatistics | where {$_.LastSuccessSync -le (Get-Date).AddDays("-30")}

$DevicesToRemove| Export-CSV .\$OutputCSVFile -notypeinformation

$DevicesToRemove | foreach-object {Remove-ActiveSyncDevice ([string]$_.Guid) -confirm:$false}