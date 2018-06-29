<#
-----------------------------------------------------------------------
Name: CommVault Updater Add Server
Author: Anthony Dunaway
Date: 06/26/18
Updated: 06/28/18
Description:
Adds a new server to ServerList.xlsx
-----------------------------------------------------------------------
#>
function Add-NewServer{
	[CmdletBinding()]
	param(
		[string]$new_server, 
		[int]$critical, 
		[int]$appdb, 
		[string]$file_path,
		[string]$application
	)
	Write-Verbose "Adding server $new_server to tracking"
	$excel = New-Object -comobject Excel.Application -Verbose:$false
	$excel.DisplayAlerts = $False
	$server_db_name = "ServerList.xlsx"
	$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
	$servers = $server_db.worksheets.item(1)
	$servers.Activate()
	$server_count = ($servers.UsedRange.Rows).count 
	$updated = 0
	for($row = 2; $row -le $server_count; $row++){	
		if(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, 1).Value())){
			$current_server = $servers.Cells.Item($row, "A").Value().ToString()
			if($current_server -eq $new_server){
				$servers.Cells.Item($row, "F").Value() = $critical
				$servers.Cells.Item($row, "G").Value() = $appdb
				$servers.Cells.Item($row, "I").Value() = $application
				$updated = 1
				Break
			}
		}
	}
	#server is new add it to the file and set its values
	if($updated -eq 0){
		$last_line = ($servers.UsedRange.Rows).count - 1
		$last_server_range = $servers.Range("A$last_line").EntireRow
		$last_server_range.Copy() | out-null
		$new_range = $servers.Range("A$server_count")
		$servers.Paste($new_range)
		$server_count = ($servers.UsedRange.Rows).count
		$servers.Cells.Item($server_count -1, "A").Value() = $new_server
		$servers.Cells.Item($server_count -1, "F").Value() = $critical
		$servers.Cells.Item($server_count -1, "G").Value() = $appdb
		$servers.Cells.Item($server_count -1, "H").Value() = ""
		$servers.Cells.Item($server_count -1, "I").Value() = $application
		$sort_range = $servers.UsedRange
		$sort_column = $servers.Range("A2")
		#keep it sorted so it's easy to find servers when looking at the file manually
		[void] $sort_range.Sort($sort_column,1,$null,$null,1,$null,1,1)
	}
	Write-Verbose "Server $new_server has been updated"
	$server_db.Save()
	$server_db.Close()
	$excel.Quit()
}

#Code to test script
# $file_path = Get-Location
# Add-NewServer -new_server "happypanda" -critical 1 -appdb 1 -file_path $file_path -application "panda cam" -verbose
# Add-NewServer -new_server "happypanda" -critical 1 -appdb 0 -file_path $file_path -verbose