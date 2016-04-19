
Param (
	[Parameter(Position=0, Mandatory=$False, HelpMessage="The server ")]
	[ValidateNotNullOrEmpty()]
	[String] $Server,

	[Parameter(Position=1, Mandatory=$False, HelpMessage="The database ")]
	[ValidateNotNullOrEmpty()]
	[String] $Database,

	[Parameter(Position=2, Mandatory=$False, HelpMessage="The filename for the output")]
	[ValidateNotNullOrEmpty()]
	[bool] $EmailRpt,

	[Parameter(Position=3, Mandatory=$False, HelpMessage="The filename for the output")]
	[ValidateNotNullOrEmpty()]
	[String] $CSVOutput
)

$SmtpServer = "srfs.vanderbilt.edu" #Enter FQDN of SMTP server
$SmtpFrom = "messaging.reports@vanderbilt.Edu" #Enter sender email address
$SmtpTo = "ECS@vanderbilt.edu" #Enter one or more recipient addresses in an array
#$SmtpTo = "mark.gossard@vanderbilt.edu"
#$SmtpCC = "ECS@vanderbilt.edu" #Enter one or more recipient addresses in an array
$SmtpSubject = "Vmail Database Report" #Enter subject of message
[int]$Log_Threshold = 2000
$script:bAlert=$false

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
{	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto

}

Function Get-MPSpace ([string]$Servername, $unit = "GB")
{
	$unit = "GB"
	$measure = "1$unit"

	Get-WmiObject -computername $Servername -query "select SystemName, Name, DriveType, FileSystem, FreeSpace, Capacity, Label from Win32_Volume where DriveType = 2 or DriveType = 3"|select SystemName, Name, @{name="SizeIn$unit";Expression={"{0:n2}" -f([int64]$_.Capacity/$measure)}}, @{name="FreeIn$unit";Expression={"{0:n2}" -f([int64]$_.freespace/$measure)}}, @{name="PercentFree";Expression={"{0:n2}" -f(([int64]$_.freespace / $_.Capacity) * 100)}}, Label
}

$Results = @()
$dt = Get-Date -format "yyyyMMdd_hhmm"

if (!$CSVOutout){
	$CSVOutput = ($MyInvocation.MyCommand.Name).Replace(".ps1","_$dt.csv")
}

if ($Database) {
	$DBs = get-mailboxdatabase $Database -status|? {$_.recovery -eq $false}|select name,server,DatabaseSize,edbfilepath,LogFolderPath,AvailableNewMailboxSpace,LastIncrementalBackup,LastFullBackup
} 

Elseif ($Server) {
	$DBs = get-mailboxdatabase -server $Server -status|? {$_.recovery -eq $false}|sort name|select name,server,DatabaseSize,edbfilepath,LogFolderPath,AvailableNewMailboxSpace,LastIncrementalBackup,LastFullBackup
}
else {
	# Assume you want all databases
	$DBs = get-mailboxdatabase -status|? {$_.recovery -eq $false}|sort name|select name,server,DatabaseSize,edbfilepath,LogFolderPath,AvailableNewMailboxSpace,LastIncrementalBackup,LastFullBackup
}

foreach ($DB in $Dbs) {
	write-host "Working on .... $($DB.name)"

	# Gather the Mailbox Counts 
	$MBX_Count = @(Get-Mailbox -Database $DB.Name).count 
	if($MBX_Count -eq $null) 
	{
		$MBX_Count = 0 
	} 


	$MBAvg = Get-MailboxStatistics -Database $DB.Name |
	%{$_.TotalItemSize.value.ToMb()} |
	Measure-Object -Average 



	# Gather the Log Counts
	$Log_Count = 0
	$driveLetter = $DB.LogFolderPath.DriveName.Trim(":") 
	$folders = $DB.LogFolderPath.PathName.Substring(2) 
	$path = "\\" + $DB.server + "\" + $driveLetter + "$" + $folders 

	#verify folder exists before trying to count logs in it 
	if(Test-Path $path -PathType Container) 
	{
		$Log_Count =(Get-ChildItem $path -filter "*.log").Count
		if ([int]$Log_Count -gt $Log_Threshold){
			$script:bAlert=$true
		} 

	} 

	$Obj = New-Object PSObject 
	$Obj | Add-Member NoteProperty -Name "Server" -value $DB.Server
	$Obj | Add-Member NoteProperty -Name "Database" -value $DB.Name
	$Obj | Add-Member NoteProperty -Name "Mailboxes" -value $MBX_Count
	$Obj | Add-Member NoteProperty -Name "Average Mailbox Size (MB)" -value ("{0:N2}" -f $MBAvg.Average)
	$Obj | Add-Member NoteProperty -Name "DBSize_GB" -value $DB.DatabaseSize.toGB()
	$Obj | Add-Member NoteProperty -Name "WhiteSpace (MB)" -value $DB.AvailableNewMailboxSpace.ToMB()
	$Obj | Add-Member NoteProperty -Name "LogCount" -value ("{0:N0}" -f $Log_Count)
	$Obj | Add-Member NoteProperty -Name "LogSize_GB" -value ("{0:N2}" -f $(($Log_Count * 1MB) /1GB))
	$Obj | Add-Member NoteProperty -Name "LogLabel" -value $(($DB.LogFolderPath).pathname.Split("\")[2])
	$Obj | Add-Member NoteProperty -Name "LastIncrementalBackup" -value $DB.LastIncrementalBackup
	$Obj | Add-Member NoteProperty -Name "LastFullBackup" -value $DB.LastFullBackup
	$Obj | Add-Member NoteProperty -Name "FullBackupAge" -value (new-timespan -start ($DB.LastFullBackup) -end (Get-date)).Days

	$Results += $Obj
}
$Results|ft -auto



$Results|sort Server,Database|export-csv .\$CSVOutput -notypeinformation
$DB_Info = $Results|Measure-Object DBSize_GB -ave -min -max
$MBX_Info = $Results|Measure-Object Mailboxes -ave -min -max

Write-Host "There are $($Results.count) databases."
Write-Host " Database Info: Average Size: $([math]::Round($DB_Info.Average,0))   Min Size: $($DB_Info.Minimum) Max Size:  $($DB_Info.Maximum)"
Write-Host " Mailbox Info:  Average Count: $([math]::Round($MBX_Info.Average,0)) Min Count:   $($MBX_Info.Minimum)  Max Count:  $($MBX_Info.Maximum)"

if ($EmailRpt -eq $true) {

 $head = '  
<style>  
BODY{font-family:Verdana;}  
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}  
TH{font-size:1.3em; border-width: 1px;padding: 2px;border-style: solid;border-color: black; text-align: Center}  
TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right}
TD.highlightyellow { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: yellow}
TD.highlightorange { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: orange}
TD.highlightred { border-width: 1px;padding: 2px;border-style: solid;border-color: black;text-align: right; background-color: red}
.center {text-align:center} 
</style>' 	

$title = "Vmail Database Report" 

$precontent = @"

  <H1>Vmail Database Report</H1>

 The following report was run on $(get-date) <br>
 <H2>SUMMARY</H2> 
"@

If(($Results.count) -eq 1) { 

	$precontent+=  "There is $($Results.count) database in this report." 

}
Else {
	$precontent+=  "There are $($Results.count) databases in this report." 
}

$precontent+= " <DL><DT> Database Info: </DT> <DD>Average Size: $([math]::Round($DB_Info.Average,0)) GB </DD> 
 <DD>Min Size: $($DB_Info.Minimum) GB </DD> 
<DD>Max Size:  $($DB_Info.Maximum) GB </DD>
</DL> 
 <DL><DT> Mailbox Info: </DT> <DD> Average Count: $([math]::Round($MBX_Info.Average,0)) </DD>
<DD>Min Count:   $($MBX_Info.Minimum)  </DD>
<DD>Max Count:  $($MBX_Info.Maximum) </DD> </DL>"

If ($script:bAlert -eq $true)
{
 $precontent+= "<br>The following databases need to be reviewed for high log counts:" 
 $precontent+=$Results|Where-Object {[int]$_.LogCount -gt $Log_Threshold}|select Server,Database,LogCount,LastIncrementalBackup,LastFullBackup |sort @{expression="LogCount";Descending=$true}|ConvertTo-Html -Head $head

}

$precontent+= " <H2>DETAIL</H2> "

	$DBReport = $Results|sort Database|select -property * |ConvertTo-Html -Head $head  -title $title 

If ($script:bAlert -eq $true)
{	$SmtpSubject += " (Attention Required)"
	$SmtpPriority = 'High'}
Else
{
	$SmtpPriority = 'Normal'}


	Send-MailMessage -From $SmtpFrom -To $SmtpTo -Subject $SmtpSubject -Body ($precontent,$DBReport|out-string) -BodyAsHTML -SMTPserver $SmtpServer -DeliveryNotificationOption onFailure -Priority $SmtpPriority
}
