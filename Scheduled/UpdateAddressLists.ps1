$WarningPreference = "SilentlyContinue"

if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue))
	{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010}

get-addresslist|Update-AddressList