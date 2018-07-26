<#
-----------------------------------------------------------------------
Name: Server Information Retrieval Script
Author:	Anthony Dunaway
Date: 06/15/18
Updated: 07/26/18
Description:
Helper script to grab our list of servers.
Returns a list of Server objects
-----------------------------------------------------------------------
#>

$script_path = $PSScriptRoot.ToString()
."$script_path\Create_Server.ps1"

Function Get-ServerList{
	[CmdletBinding()]
	param(
		[string] $file_path
	)
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
	For($line=2; $line -le $server_count - 1; $line++){
		$current_server = $servers.Cells.Item($line,"A").Value().ToString().Trim()
		$critical = $servers.Cells.Item($line,"F").Value().ToString().Trim()
		$appdb = $servers.Cells.Item($line,"G").Value()
		$staff = $servers.Cells.Item($line,"H").Value().ToString().Trim().Split(",")
		$lead = $servers.Cells.Item($line,"I").Value().ToString().Trim()
		$apps = $servers.Cells.Item($line,"J").Value().ToString().Trim().Split(",")
		$all_servers += New-Server -name $current_server -critical $critical -appdb $appdb -staff $staff -lead $lead -applications $apps
	}
	Write-Verbose "Closing Excel and Cleaning Up"
    $excel.Workbooks.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($servers)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($server_db)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
	[GC]::Collect()
	Write-Verbose "List Complete"
	if(($critical -eq $True) -or ($full -eq $True)){
		Return $all_servers_dict
	}
	else{
		Return $all_servers
	}
}

#test the functionality
# $test = Get-ServerList -verbose -file_path "C:\Development\CommVaultScriptDev\"

# $find_ob = $test | Where-Object {$_.Name -eq "lrn_prod"}
# $find_ob
# foreach($server in $test){
	# Write-Host $server.Name
	# Write-Host $server.Critical
	# Write-Host $server.AppDB
	# Write-Host $server.Staff
	# Write-Host $server.Applications
# }
