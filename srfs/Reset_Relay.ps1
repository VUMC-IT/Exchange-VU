Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin

Set-ReceiveConnector -Identity "its-hcwnem18\SRFS-Connector" -RemoteIPRanges ("10.1.137.199","10.1.140.24","10.1.140.25","10.1.140.26","10.127.27.55","129.59.160.226","129.59.160.235")
