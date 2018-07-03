<#
-----------------------------------------------------------------------
Name: Server Information Retrieval Script
Author:	Anthony Dunaway
Date: 06/15/18
Updated: 07/03/18
Description:
Helper script to grab our list of servers and their critical status.
If critical flag is passed returns a hashtable with the server name as the key 
and the critical status as the value. Otherwise returns a list of the servers
-----------------------------------------------------------------------
#>
Function Get-ServerList{
	[CmdletBinding()]
	param(
		[switch] $full,
		[switch] $critical,
		[string] $file_path
	)
    $all_servers_dict = @{}
	$all_servers = @()
	Write-Verbose "Opening Excel"
    $excel = New-Object -comobject Excel.Application -Verbose:$false
    $excel.DisplayAlerts = $False
    $server_db_name = "ServerList.xlsx"
	Write-Verbose "Opening the Server List"
    $server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
    $servers = $server_db.worksheets.item(1)
    $server_count = ($servers.UsedRange.Rows).count
	Write-Verbose "Building the server list"
	if($full -eq $true){
		For($line=2; $line -le $server_count - 1; $line++){
			$server_data = @()
			$current_server = $servers.Cells.Item($line,"A").Value().ToString().Trim()
			Write-Debug "Current server : $current_server"
			$server_is_critical = $servers.Cells.Item($line,"F").Value().ToString().Trim()
			Write-Debug "Critical Status : $server_is_critical"
			$staff = $servers.Cells.Item($line,"H").Value().ToString().Trim()
			Write-Debug "Staff are $staff"
			$applications = $servers.Cells.Item($line,"I").Value().ToString().Trim()
			Write-Debug "Applications include $applications"
			$server_data += $server_is_critical
			$server_data += $staff
			$server_data += $applications
			$all_servers_dict.add($current_server, $server_data)
		} 
	}
	elseif($critical -eq $true){
		For($line=2; $line -le $server_count - 1; $line++){
			$current_server = $servers.Cells.Item($line,"A").Value()
			$current_server = $current_server.ToString().Trim()
			Write-Debug "Current server : $current_server"
			$server_is_critical = $servers.Cells.Item($line,"F").Value()
			Write-Debug "Critical Status : $server_is_critical"
			$all_servers_dict.add($current_server, $server_is_critical)
		}
	}
	else{
		For($line=2; $line -le $server_count - 1; $line++){
				$current_server = $servers.Cells.Item($line,"A").Value()
				$current_server = $current_server.ToString().Trim()
				Write-Debug "Current server : $current_server"
				$all_servers += $current_server
		} 
	}
	Write-Verbose "Closing Excel and Cleaning Up"
    $excel.Workbooks.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($servers)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($server_db)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
	Write-Verbose "List Complete"
	if($critical -eq $true){
		Return $all_servers_dict
	}
	else{
		Return $all_servers
	}
}#Function

#test the functionality
# $test = Get-ServerList -verbose -full -file_path "I:\ISE\CommVault\CommVaultScript\"
# $test = Get-ServerList -verbose -critical -file_path "I:\ISE\CommVault\CommVaultScript\"
# $test = Get-ServerList -verbose -file_path "I:\ISE\CommVault\CommVaultScript\"
# foreach($server in $test.keys){
	# Write-Verbose $server
	# foreach($stat in $test[$server]){
		# Write-Verbose $stat
	# }
# }
