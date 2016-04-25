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

$User = Read-Host "Enter VUnetID"

set-CASMailbox $User -ActiveSyncEnabled:$False -ImapEnabled:$False -MAPIEnabled:$False -EwsEnabled:$False -OWAEnabled:$False -PopEnabled:$False -EwsAllowEntourage:$False -EwsAllowMacOutlook:$False -EwsAllowOutlook:$False -MapiBlockOutlookRpcHttp:$True -confirm

Get-CSUser $User | Revoke-CsClientCertificate


#Report the Results
	
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
	
return $result | select name, ActiveSyncEnabled, MAPIEnabled, OWAEnabled, ImapEnabled, PopEnabled, EwsEnabled, EwsAllowMacOutlook, EwsAllowEntourage, EwsAllowOutlook, MapiBlockOutlookRpcHttp