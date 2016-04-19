#$DebugPreference = "SilentlyContinue"
$DebugPreference = "Continue"

#Begin customization-------------------------
$SmtpServer = "srfs.vanderbilt.edu" #Enter FQDN of SMTP server
$SmtpFrom = "Messaging.Reports@vanderbilt.edu" #Enter sender email address
$SmtpTo = "ECS@vanderbilt.edu" #Enter one or more recipient addresses in an array
#$SmtpTo = "e.zafar@vanderbilt.edu","taj.wolff@vanderbilt.edu"
#$SmtpCC = "ECS@vanderbilt.edu" #Enter one or more recipient addresses in an array
$SmtpSubject = "VUIT Team Member Script" #Enter subject of message
#End customization---------------------------



$Group = "MCIT Team Members"
$OneOffGroup = "MCIT_oneoff_membership"

import-module -name ActiveDirectory

Write-Debug "Getting users in specific departments"
$users = get-aduser -filter {(department -like "*108038") -or (department -like "*108039") -or (department -like "*108041") -or (department -like "*108042") -or (department -like "*108043") -or (department -like "*108044") -or (department -like "*108045") -or (department -like "*108046") -or (department -like "*108047") -or (department -like "*108048") -or (department -like "*108049") -or (department -like "*108051") -or (department -like "*108052") -or (department -like "*108053") -or (department -like "*108054") -or (department -like "*108055") -or (department -like "*108056") -or (department -like "*108057") -or (department -like "*108058") -or(department -like "*108059") -or (department -like "*108061") -or (department -like "*108062") -or (department -like "*108063") -or (department -like "*108064") -or (department -like "*108065") -or (department -like "*108066") -or (department -like "*108067") -or (department -like "*108068") -or (department -like "*108069") -or (department -like "*108080") -or (department -like "*108081") -or (department -like "*108082") -or (department -like "*108083") -or (department -like "*108084") -or (department -like "*108085") -or (department -like "*108086") -or (department -like "*108087") -or (department -like "*108088") -or (department -like "*108089") -or (department -like "*108090") -or (department -like "*108091") -or (department -like "*108092") -or (department -like "*108093") -or (department -like "*108094") -or (department -like "*108095") -or (department -like "*108139")}
If ($users){
	Write-Debug "Adding One off Users"
	#$users += get-adgroupmember "$OneOffGroup"
	Write-Debug "User total before OneOff in Users is $($Users.count)"
	$OneOffUsers = get-adgroupmember "$OneOffGroup"
	ForEach ($Oneoff in $OneOffUsers) 
	{
		$Matches = $False 
		$CurrentUser = get-aduser $OneOff
		ForEach ($user in $users)
		{
			If ($User -match $CurrentUser)
			{
				$Matches = $True
			}
		}
		If (!$Matches)
		{
			$Users += $OneOff
		}
	}
	$usercount = $users.count

	Write-Debug "User total in Users after OneOff is $Usercount"

	Write-Debug "Getting current Group Membership"
	$current = get-adgroupmember "$Group"
	$currentcount = $current.count
	Write-Debug "User total in current is $Currentcount"

	Write-Debug "Creating Remove Array"
	$RemoveArray = Compare-Object -ReferenceObject $Users -DifferenceObject $current -Property DistinguishedName -IncludeEqual -PassThru | Where-Object { $_.SideIndicator -eq "=>" } 


	Write-Debug "Remove Foreach Loop"
	If ($RemoveArray)
	{
		ForEach ($User in $RemoveArray)
		{
			$Account = $user.samaccountname
			Write-Debug "Removing $Account"
			get-adgroup $group | Remove-adgroupmember -members $Account -confirm:$false
		}
	}
	Write-Debug "Creating Add Array"
	$AddArray = Compare-Object -ReferenceObject $current -DifferenceObject $users -Property DistinguishedName -IncludeEqual -PassThru | Where-Object { $_.SideIndicator -eq "=>" } 



	Write-Debug "Add Foreach Loop"
	If ($AddArray)
	{
		ForEach ($User in $AddArray)
		{
			$Account = $user.samaccountname
			Write-Debug "Adding $Account"
			Add-adgroupmember $Group -members $Account -confirm:$false -ea Silentlycontinue
		}
	}
}
Else{
	$HtmlBody = "There was a problem with the script. Please check the group membership."
	$SmtpSubject += " (Attention Required)"
	$SmtpPriority = 'High'

	Send-MailMessage -From $SmtpFrom -To $SmtpTo  -Subject $SmtpSubject -Body $HtmlBody -SMTPserver $SmtpServer -DeliveryNotificationOption onFailure -Priority $SmtpPriority
}