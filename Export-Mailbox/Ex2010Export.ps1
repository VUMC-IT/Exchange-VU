#Updated 03/02/16 - Pratt, Jeremy
#Update checks for and exports archive databases as well, if the user has one. 


$export_mbx = Read-Host "Enter VUNet ID Here"

$Export_Path_Base = "\\groups-nas2.its.vanderbilt.edu\cis-swe\"

# $Export_Path_Base = "\\groups-nas2.its.vanderbilt.edu\iarc\Administrators\Common\Exchange_PST\"
# $Export_Path_Base = "\\groups-nas1\its_swe\"

if (!(Test-Path ($Export_Path_Base + $export_mbx) -pathType container)) {
	New-Item ($Export_Path_Base + $export_mbx) -ItemType Directory
	Start-Sleep -s 30
}

$Export_Path = ($Export_Path_Base + $export_mbx + "\" + $Export_mbx + ".pst")
$Archive_Path = ($Export_Path_Base + $export_mbx + "\" + $Export_mbx + "-Archive.pst")

$archive = Get-Mailbox -identity $export_mbx | select archivedatabase

If ($archive.archivedatabase -eq $null) {
    write-host "No archive database found" -foregroundcolor Green
    New-MailboxExportRequest -Mailbox $export_mbx -FilePath $Export_Path
    }
    else {
    New-MailboxExportRequest -Mailbox $export_mbx -FilePath $Export_Path
    New-MailboxExportRequest -Mailbox $export_mbx -Filepath $Archive_Path -IsArchive
    }
