$Path = "E:\TransportRoles\Logs\ProtocolLog\SmtpReceive"
$Search = Read-host "Enter IP to Search For"
$results = Get-ChildItem -Path $Path | Select-String -Pattern $Search
If($results.count -gt 1) {

Write-Host "The IP was found in the SmtpReceive Logs" -foregroundcolor green
$results | select line | Export-Csv -NoTypeInformation results.csv

}
Else {
Write-Host "The IP was not found in the SmtpReceive Logs" -foregroundcolor red
}



