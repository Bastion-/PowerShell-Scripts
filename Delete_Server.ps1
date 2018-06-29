<#
-----------------------------------------------------------------------
Name: CommVault Updater Delete Server
Author: Anthony Dunaway
Date: 06/26/18
Updated: 06/28/18
Description:
Deletes a server from ServerList.xlsx
-----------------------------------------------------------------------
#>
function Remove-DeleteServer{
	[CmdletBinding()]
	param(
		[string]$server, 
		[string]$file_path
	)
	Write-Verbose "Deleting Server $server"
	$excel = New-Object -comobject Excel.Application -Verbose:$false
	$excel.DisplayAlerts = $False
	$server_db_name = "ServerList.xlsx"
	$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
	$servers = $server_db.worksheets.item(1)
	$server_count = ($servers.UsedRange.Rows).count 
	$removed = 0
	for($row = 2; $row -le $server_count; $row++){
		if(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, 1).Value())){
			$current_server = $servers.Cells.Item($row, "A").Value().ToString()
			if($current_server -eq $server){
				$delete_range = $servers.Cells.Item($row, "A").EntireRow
				[void]$delete_range.Delete()
				$removed = 1
				Break
			}
		}
	}
	if($removed -eq 1){
		Write-Verbose "$server record removed"
	}
	else{
		Write-Verbose "Server $server was not found no need to delete"
	}
	$server_db.Save()
	$server_db.Close()
	$excel.Quit()
}

#Script Test Code
# $file_path = Get-Location
# Remove-DeleteServer -server "happypanda" -file_path $file_path -verbose