<#
-----------------------------------------------------------------------
Name: Create Server Object
Author:	Anthony Dunaway
Date: 07/19/18
Updated: 07/26/18
Description:
Creates a new server object for use in the CommVault script
-----------------------------------------------------------------------
#>

function New-Server {
	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$True)]
		[string]$name,
		[Parameter(Mandatory=$True)]
		[int]$critical,
		[Parameter(Mandatory=$True)]
		[int]$appdb,
		[Parameter(Mandatory=$True)]
		[string[]]$staff, 
		[Parameter(Mandatory=$True)]
		[string]$lead, 
		[Parameter(Mandatory=$True)]
		[string[]]$applications
	)
	
	New-Object psobject -property @{
		Name = $name
		Critical = $critical
		AppDB = $appdb
		Staff = $staff
		Lead = $lead
		Applications = $applications
	}
}

#Testing

# $staff = @("Ophelia.PandaQueen", "Melchior.PandaKing")
# $apps = @("panda cam", "panda time")

# $current_server = New-Server -name "happypanda" -critical 1 -appdb 0 -staff $staff -applications $apps

# $current_server.Name
# $current_server.Critical
# $current_server.AppDB
# $current_server.Staff
# $current_server.Applications
