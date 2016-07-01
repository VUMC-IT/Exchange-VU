<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.120
	 Created on:   	6/30/2016 4:54 PM
	 Created by: 	chanct  	 
	 Organization: 	 
	 Filename: add-distmanagers.ps1    	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
$groups = Get-DistributionGroup -filter { name -like "VUSN*" }

foreach ($group in $groups)
{
	$managers = $group.managedby
	$newmanagers = @("ds.vanderbilt.edu/users/grovet2",`
	                 "ds.vanderbilt.edu/users/mcnewr",`
	                 "ds.vanderbilt.edu/users/johns469",`
	                "ds.vanderbilt.edu/users/loerchse")
	$managers.Add($newmanagers)
	set-distributiongroup $group -managedby $managers -BypassSecurityGroupManagerCheck
}