[cmdletBinding()]
	Param (
		[Parameter(Position = 0, Mandatory = $true, HelpMessage = "Type of Incident")]
		[ValidateSet("under", "over")]
		[String]$Incidenttype,
		[Parameter(Position = 1, Mandatory = $false, HelpMessage = "User data object")]
		[ValidateNotNullOrEmpty()]
		$userobj,
		[Parameter(Position = 2, Mandatory = $false, HelpMessage = "data for the ticket description")]
		$desclist,
		[Parameter(Position = 3, Mandatory = $false, HelpMessage = "data for the ticket resolution")]
		$resolvlist,
        [Parameter(Position = 4, Mandatory = $false, HelpMessage = "Users name")]
        $Name,
        [Parameter(Position = 5, Mandatory = $false, HelpMessage = "Vunet")]
        $VUnet,
        [Parameter(Position = 6, Mandatory = $false, HelpMessage = "Current Quota")]
        $CurrentQuota,
        [Parameter(Position = 7, Mandatory = $false, HelpMessage = "Quota")]
        $Quota
        
	)
	$uri = "https://pegasus.mc.vanderbilt.edu/services/incident.asmx"
	if ($Incidenttype -eq "under")
	{
		$IncidentTitle = "Quota Report Increase Incident - Under 5GB"
		$IncidentDescription = "The following users exceed 95% of their quota 
							$desclist"
		$IncidentResolution = "Resolution Applied: Server 
						   Resolution: The quotas for the following users were increased by 1GB
							$resolvlist"
		
		$xmlreq = "<soapenv:Envelope xmlns:soapenv=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:inc=`"http://itsm.vanderbilt.edu/services/Incident/`">
   			<soapenv:Header/>
   			<soapenv:Body>
      			<inc:CreateIncident>
         			<!--Optional:-->
         			<inc:apiKey>24D7002F9CEB6D4FE0535418980A12FB</inc:apiKey>
         			<!--Optional:-->
         			<inc:request>
            			<inc:AreaName>data</inc:AreaName>
            			<inc:AssignmentGroupName>VUIT Collaboration</inc:AssignmentGroupName>
						<inc:AssignedPersonID>chanct</inc:AssignedPersonID>
            			<inc:CategoryName>incident</inc:CategoryName>
						<inc:CCName>Solved</inc:CCName>
            			<inc:CIName>EXCH-QUOTA</inc:CIName>
						<inc:Description>$IncidentDescription</inc:Description>
            			<inc:ImpactID>8</inc:ImpactID>
            			<inc:Solution>$IncidentResolution</inc:Solution>
						<inc:SourceName>e-mail</inc:SourceName>
            			<inc:StatusName>Resolved</inc:StatusName>
						<inc:SubAreaName>storage limit exceeded</inc:SubAreaName>
            			<inc:Title>$IncidentTitle</inc:Title>
            			<inc:UrgencyID>4</inc:UrgencyID>
            			<inc:UserID>chanct</inc:UserID>
					</inc:request>
     			 </inc:CreateIncident>
   			</soapenv:Body>
			</soapenv:Envelope>"
	}
	if ($Incidenttype -eq "over")
	{   $IncidentTitle = "WARNING: Mailbox Cleanup Requested for $Name"
		$IncidentDescription = "The user $Name has exceeded 95% of their $CurrentQuota GB mailbox quota. We have proactively adjusted the quota from $CurrentQuota GB to $Quota GB . VUIT Collaboration Team requests the LAN Manager/Local Support Provider advise and assist the user with cleaning up their mailbox. A breakdown of folder usage statistics that will aid the cleanup process has been emailed to the user. VUIT Collaboration will be glad to answer any questions you may have."
		
		
		$xmlreq = "<soapenv:Envelope xmlns:soapenv=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:inc=`"http://itsm.vanderbilt.edu/services/Incident/`">
   			<soapenv:Header/>
   			<soapenv:Body>
      			<inc:CreateIncident>
         			<!--Optional:-->
         			<inc:apiKey>24D7002F9CEB6D4FE0535418980A12FB</inc:apiKey>
         			<!--Optional:-->
         			<inc:request>
            			<inc:AreaName>data</inc:AreaName>
            			<inc:AssignmentGroupName>Helpdesk</inc:AssignmentGroupName>
            			<inc:CategoryName>incident</inc:CategoryName>
            			<inc:CIName>DESKTOP</inc:CIName>
						<inc:Description>$IncidentDescription</inc:Description>
            			<inc:ImpactID>8</inc:ImpactID>
            			<inc:SourceName>e-mail</inc:SourceName>
            			<inc:SubAreaName>storage limit exceeded</inc:SubAreaName>
            			<inc:Title>$IncidentTitle</inc:Title>
            			<inc:UrgencyID>4</inc:UrgencyID>
            			<inc:UserID>$vunet</inc:UserID>
					</inc:request>
     			 </inc:CreateIncident>
   			</soapenv:Body>
			</soapenv:Envelope>"
	}
	
	$env:PegasusResult = Invoke-WebRequest -Uri $uri -Method Post -ContentType "text/xml" -Body $xmlreq
	
	