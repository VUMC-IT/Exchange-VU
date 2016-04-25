############Establishes implicit session w production env#####
#session information for implicit remote to CAS
$exsess = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://vuit-hcwnem133/PowerShell/ -Authentication Kerberos
#session information/options for implicit remote to Lync pool
#$lyncsess = New-PSSession -ConnectionUri https://hcfepool2013.vanderbilt.edu/ocsPowerShell -Authentication NegotiateWithImplicitCredential
#Import exchange commands
Import-PSSession -Session $exsess -AllowClobber
#Import Lync commands
#Import-PSSession -Session $lyncsess -allowclobber
##########################################################

$Users = Get-Content -Path .\FixImmunizationDisableMultiple-data.txt
$Report = @()
$dt = Get-Date -format "yyyyMMdd_hhmm"

ForEach ($User in $Users) {
    #Enable Mail Services 
    set-CASMailbox $User -ActiveSyncEnabled:$True -ImapEnabled:$True -MAPIEnabled:$True -EwsEnabled:$True -OWAEnabled:$True -PopEnabled:$True -EwsAllowEntourage:$True -EwsAllowMacOutlook:$True -EwsAllowOutlook:$True -MapiBlockOutlookRpcHttp:$False

    #Get results
    $mailresult = get-casmailbox $User
    $result = new-object -TypeName PSObject -Property @{
        name = $mailresult.name
        ActiveSyncEnabled = $mailresult.ActivesyncEnabled
        MAPIEnabled = $mailresult.MAPIEnabled
        OWAEnabled = $mailresult.OWAEnabled
        ImapEnabled = $mailresult.ImapEnabled
        PopEnabled = $mailresult.PopEnabled
        EwsEnabled = $mailresult.EWSEnabled
        EwsAllowMacOutlook = $mailresult.EWSAllowMacOutlook
        EwsAllowEntourage = $mailresult.EWSAllowEntourage
        EwsAllowOutlook = $mailresult.EwsAllowOutlook
        MapiBlockOutlookRpcHttp = $mailresult.MapiBlockOutlookRpcHttp
    }
	
    #Save results to $Report
    $Report += $result
}

#Report results to screen
$Report | select Name, ActiveSyncEnabled, MAPIEnabled, OWAEnabled, ImapEnabled, PopEnabled, EwsEnabled, EwsAllowMacOutlook, EwsAllowEntourage, EwsAllowOutlook, MapiBlockOutlookRpcHttp 

#Save results to log file
$Report | select Name, ActiveSyncEnabled, MAPIEnabled, OWAEnabled, ImapEnabled, PopEnabled, EwsEnabled, EwsAllowMacOutlook, EwsAllowEntourage, EwsAllowOutlook, MapiBlockOutlookRpcHttp | Out-File -FilePath ".\Logs\FixImmunizationDisableResults_$dt.txt" -Force
