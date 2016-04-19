# This Script will allow you to simulate a group having ownership of a Distribution group in Exchange 2010
# !!!! Change the value of $dn_storage to match the Customer Attribute you want to use in your organization !!!!
# 
#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################
#
# Written by Matthew Byrd
# Matbyrd@microsoft.com
# Last Updated 05.02.2012
#
# v2 Updated 10.04.2011
# Special Thanks to Dann Cox and Graham Thorpe for pointing out some issues,
# suggesting some changes and testing the updated script
#
# v3 Updated 05.02.2012
# Now properly handles distribution groups that have special characters and spaces in the name
# 


Param($DistributionGroup = $null,$GroupOwner = $null)

# Sets all users in the DL listed in $dn_storage as managing the DL they are listed in
Function SetUserAsOwners {
	Param ([string]$DistributionGrouptoSet)
	
	# Handle single DG vs Processing everything
	if ($DistributionGrouptoSet -eq ""){
	
		# Get a list of all groups that we need to manipulate
		$Groupstoset = Get-distributiongroup -resultsize unlimited -filter "($dn_storage -like 'CN*')"
		
	}
	# Set our Grouptoset to just the single group that was passed into the function
	else { $Groupstoset = Get-distributiongroup "$DistributionGrouptoSet" }
	
	# Process each group
	$Groupstoset | foreach {
	
		# Setting the array of users to null so that it doesn't keep adding to the array with each loop
		[array]$DNOfUserstoset = $null
		
		# Setting CheckedUserstoSet to Null to ensure we don't continue building the array over time
		[array]$CheckedUserstoset = $null
		
		Write-Host "Setting Members of" $_.($dn_storage) "as owners on" $_.identity
		
		# Get a list of the users that need to be listed as managers of the DL
		# This will preform this search recursively
		$Userstoset = Get-ADGroupMember $_.($dn_storage) -recursive
		
		# Convert the output from get-adgroupmember into an array of DNs
		$Userstoset | foreach { [array]$DNOfUserstoset = $DNOfUserstoset + [string]$_.distinguishedname }
		
		# Verify that each of the users in the array is a mailbox
		# This eliminates any nested groups / contact / or users and just leaves us with the mailboxes
		$DNOfUserstoset | foreach {
		
			If (Get-mailbox $_ -erroraction silentlycontinue){[array]$CheckedUserstoset = $CheckedUserstoset + $_ }
			else {}
		}
		
		# Throw any duplicates out of the $checkedUsersToSet
		$CheckedUserstoset = $CheckedUserstoset | Select-Object -Unique
		
		# Set that list of mailboxes as owners of the DL
		# Throw a warning if we didn't get any valid objects returned
		if ($CheckedUserstoset -eq $null){Write-warning "Didn't Find any valid objects in Owning Group"}
		else { Set-distributiongroup $_.identity -managedby $CheckedUserstoset -BypassSecurityGroupManagerCheck }
		
	}

}

# Setup a DL as "owner" of another DL
# This will place the DN of DistributionGroupOwner into the $dn_Storage file of the Distribtiongroup
Function SetDNofGroupOwner {
	Write-Host "Setting" $GroupOwner "as the owner of" $DistributionGroup

	# Build and Execute the command that we need to make this change
	$commandtorun = "Set-Distributiongroup -identity `'" + $DistributionGroup + "`' -" + $dn_storage + " `(get-adgroup `'" + $GroupOwner + "`'`)`.distinguishedname"
	Invoke-Expression $commandtorun
}

# Main Body
# ===============================

# Attribute to use for storing DN of group owner
# !!!! --- Change this to the correct attribute for your Organization --- !!!! #
$dn_storage = "CustomAttribute5"

#Check the OS Version
if ([system.environment]::OSversion.Version.Major -eq 6 -and [system.environment]::OSversion.Version.Minor -ge 1 -and (get-wmiobject Win32_OperatingSystem -comp .).Caption.Contains("R2")  ) {}
else {
	Write-Error "The Operating System requirements are not met, you must be running at least Windows 2008 R2"
	exit
}

# Check to see if the Exchange snapin is loaded, if not load it
if ((Get-PSSession | where {$_.configurationname -eq "Microsoft.Exchange"}) -eq $null) {
	Write-Host "Loading Exchange Shell"

	# Load up Exchange Powershell so we have the Exchange cmdlets
	# !!!! --- You will need to change this Path if your Exchange Install is not in the Default Location --- !!!! #
		. 'c:\Program Files\Microsoft\Exchange Server\v14\Bin\RemoteExchange.ps1'
		Connect-ExchangeServer -auto
}

# Import the AD management Module
Import-Module ActiveDirectory

# If no parameters passed process all Distribution groups
If (($GroupOwner -eq $null) -and ($DistributionGroup -eq $null)){ SetUserAsOwners }

# If we have a Distributiongroup but not an owner then just process that DL
elseif (($GroupOwner -eq $null) -and ($DistributionGroup -ne $null)){ SetUserAsOwners -DistributionGrouptoSet $DistributionGroup }

# If we have DL owner and don't have a DL then we need to generate an error
elseif (($DistributionGroup -eq $null) -and ($GroupOwner -ne $null)) {Write-Error "Incorrect Syntax"}

# If none of the above then we should have both DL and DLOwner so set the DL owner attribute
else { SetDNofGroupOwner }