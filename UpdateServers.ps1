<#
-----------------------------------------------------------------------
Name: CommVault Backup Report Script Server Updater
Author: Anthony Dunaway
Date: 03/15/18
Updated: 06/20/18
Description:
Adds a new server, deletes a server, or updates a current server. 
Servers can be entered individually or in bulk via a csv file. 
Users will need to know the server name, whether it is a critical server, 
if it is an application database, and the email address of the staff to be 
notified. If the staff are unknown the file will be added to tracking but 
not to email notifications.
-----------------------------------------------------------------------
#>

."I:\ISE\CommVault\CommVaultScript\Helper_Scripts\User_Input.ps1"

#----------------------------------------------------------------------
#Function to remove a server from the Server List
#----------------------------------------------------------------------
function Update-RemoveServerRecord([string]$server, [string]$file_path){
	$excel = New-Object -comobject Excel.Application
	$excel.DisplayAlerts = $False
	$server_db_name = "ServerList.xlsx"
	$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
	$servers = $server_db.worksheets.item(1)
	$servers.Activate()
	$server_count = ($servers.UsedRange.Rows).count 
	$removed = 0
	For($row = 2; $row -le $server_count; $row++){
		If(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, 1).Value())){
			$current_server = $servers.Cells.Item($row, "A").Value().ToString()
			If($current_server -eq $server){
				$delete_range = $servers.Cells.Item($row, "A").EntireRow
				[void]$delete_range.Delete()
				$removed = 1
				Break
			}
		}
	}
	If($removed -eq 1){
		Write-Host "$server record removed"
	}
	Else{
		Write-Host "Server $server was not found no need to delete"
	}
	$server_db.Save()
	$server_db.Close()
	$excel.Quit()
}
#----------------------------------------------------------------------
#Function to remove a server from a particular user
#----------------------------------------------------------------------
function Update-StaffMod([string[]]$staff, [string]$server, [string]$file_path){
	$excel = New-Object -comobject Excel.Application
	$excel.Visible = $False
	$excel.DisplayAlerts = $False
	$staff_server_list = $excel.Workbooks.Open("$file_path\StaffServerList.xlsx")
	$staff_list = $staff_server_list.worksheets.item(1)
	$staff_count = ($staff_list.UsedRange.Rows).Count
	$removed = 0
	#check every cell in every row for matches and delete them
	ForEach($person in $staff){
		For($row = 2; $row -le $staff_count; $row++){
			$staff_range= $staff_list.UsedRange.Cells
			$column_count = $staff_range.Columns.Count
			$current_staff = $staff_list.Cells.Item($row,"A").Value()
			If($current_staff -eq $person){
				For ($col = 2; $col -le $column_count; $col++){
					If(-Not [string]::IsNullOrEmpty($staff_list.Cells.Item($row, $col).Value())){
						If($staff_list.Cells.Item($row, $col).Value().ToString() -eq $server.ToString()){
							$staff_list.Cells.Item($row,$col).Clear()| out-null
							$removed = 1
							Write-Host "$person will no longer be notified regarding $server"
							Break
						}
					}
				}
			}
		}
	}
	#This removes any empty rows left behind after deleting a server
	If($removed -eq 1){
		$staff_server_list.Save()
		$staff_server_list.Close()
		$excel.Quit()
		Update-CleanupStaffServerList
	}
	Else{
		$staff_server_list.Close()
		$excel.Quit()
		Write-Host "$person was not tracking $server."
	}
}
#----------------------------------------------------------------------
#Function to remove a server comepletely from Staff Server List 
#----------------------------------------------------------------------
function Update-RemoveServerNotice([string]$server, [string]$file_path){
	$excel = New-Object -comobject Excel.Application
	$excel.DisplayAlerts = $False
	$staff_file = "StaffServerList.xlsx"
	$staff_server_list = $excel.Workbooks.Open("$file_path\$staff_file")
	$staff_list = $staff_server_list.worksheets.item(1)
	$staff_count = ($staff_list.UsedRange.Rows).count
	$removed = 0
	#check every cell in every row for matches and delete them
	For($row = 2; $row -le $staff_count; $row++){
		$staff_range= $staff_list.UsedRange.Cells
		$column_count = $staff_range.Columns.Count
		For ($col = 2; $col -le $column_count; $col++){
			If(-Not [string]::IsNullOrEmpty($staff_list.Cells.Item($row, $col).Value())){
				If($staff_list.Cells.Item($row, $col).Value().ToString() -eq $server.ToString()){
					$staff_list.Cells.Item($row,$col).Clear()| out-null
					$removed = 1
				}
			}
		}
	}
	If($removed -eq 1){
		Write-Host "Notifications will no longer be sent regarding $server"
		$staff_server_list.Save()
		$staff_server_list.Close()
		$excel.Quit()
		Update-CleanupStaffServerList
	}
	Else{
		$staff_server_list.Close()
		$excel.Quit()
		Write-Host "Server $server was not found."
	}
}
#----------------------------------------------------------------------
#Function to cleanup server staff_list after deleting a server
#----------------------------------------------------------------------
function Update-CleanupStaffServerList(){
	$excel = New-Object -comobject Excel.Application
	$excel.DisplayAlerts = $False
	$staff_server_list = $excel.Workbooks.Open("$file_path\StaffServerList.xlsx")
	$staff_list = $staff_server_list.worksheets.item(1)
	$staff_count = ($staff_list.UsedRange.Rows).count
	For($row = $staff_count; $row -ge 2; $row--){
		$empty = 1
		$staff_range= $staff_list.UsedRange.Cells
		$column_count = $staff_range.Columns.Count
		For ($col = 2; $col -le $column_count; $col++){
			If(-Not [string]::IsNullOrEmpty($staff_list.Cells.Item($row, $col).Value())){
				$empty = 0
			}
		}
		If($empty -eq 1){
			$delete_range = $staff_list.Cells.Item($row, "A").EntireRow
			[void]$delete_range.Delete()
		}
	}
	$staff_server_list.Save()
	$staff_server_list.Close()
	$excel.Quit()
}
#----------------------------------------------------------------------
#Function to add new server to the Server List
#----------------------------------------------------------------------
function Add-NewServerEntry([string]$new_server, [int]$critical, [int]$appdb, [string]$file_path, [string]$application){
	$excel = New-Object -comobject Excel.Application
	$excel.Visible = $False
	$excel.DisplayAlerts = $False
	$server_db_name = "ServerList.xlsx"
	$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
	$servers = $server_db.worksheets.item(1)
	$servers.Activate()
	$server_count = ($servers.UsedRange.Rows).count 
	$updated = 0
	For($row = 2; $row -le $server_count; $row++){	
		If(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, 1).Value())){
			$current_server = $servers.Cells.Item($row, "A").Value().ToString()
			If($current_server -eq $new_server){
				$servers.Cells.Item($row, "G").Value() = $critical
				$servers.Cells.Item($row, "H").Value() = $appdb
				$servers.Cells.Item($row, "I").Value() = $application
				$updated = 1
				Break
			}
		}
	}
	#server is new add it to the file and set its values
	If($updated -eq 0){
		$last_line = ($servers.UsedRange.Rows).count - 1
		$last_server_range = $servers.Range("A$last_line").EntireRow
		$last_server_range.Copy() | out-null
		$new_range = $servers.Range("A$server_count")
		$servers.Paste($new_range)
		$server_count = ($servers.UsedRange.Rows).count
		$servers.Cells.Item($server_count -1, "A").Value() = $new_server
		$servers.Cells.Item($server_count -1, "G").Value() = $critical
		$servers.Cells.Item($server_count -1, "H").Value() = $appdb
		$servers.Cells.Item($server_count -1, "I").Value() = $application
		$sort_range = $servers.UsedRange
		$sort_column = $servers.Range("A2")
		#keep it sorted so it's easy to find servers when looking at the file manually
		[void] $sort_range.Sort($sort_column,1,$null,$null,1,$null,1,1)
	}
	Write-Host "Server $new_server has been udpated"
	$server_db.Save()
	$server_db.Close()
	$excel.Quit()
}
#----------------------------------------------------------------------
#Function to update which staff to notify in Staff Server List
#----------------------------------------------------------------------	
function Update-StaffToNotify([string[]]$staff,[string]$new_server, [string]$file_path){
	$excel = New-Object -comobject Excel.Application
	$excel.DisplayAlerts = $False
	$staff_server_list = $excel.Workbooks.Open("$file_path\StaffServerList.xlsx")
	$staff_list = $staff_server_list.worksheets.item(1)
	$staff_count = ($staff_list.UsedRange.Rows).count
	foreach($person in $staff){
		$match = 0
		$count = 1
		$staff_row = 0
		$staff_found = 0
		#Make sure there are no blank rows at the end of the file. Blank rows throw ugly errors
		for($row = $staff_count; $row -ge 2; $row--){
			$staff_name = $staff_list.Cells.Item($row, "A").Value()
			if([string]::IsNullOrEmpty($staff_name)){
				$delete_range = $staff_list.Cells.Item($row, "A").EntireRow
				[void]$delete_range.Delete()
			}
		}
		#Check to see if the staff member already has a row
		for($row = 2; $row -le $staff_count; $row++){
			$staff_name = $staff_list.Cells.Item($row, "A").Value().ToString()
			if($staff_name -eq $person){
				$staff_row = $row
				$staff_found = 1
			}
			if($staff_found -eq 1){
				Break
			}
		}
		#Person is new, add them to the file with the server
		if($staff_found -eq 0){
			$new_staff_row = $staff_list.Cells.Item($staff_count + 1,1).EntireRow()
			$active_range = $new_staff_row.Activate()
			$active_range = $new_staff_row.insert($xlShiftDown)
			$staff_list.Cells.Item($staff_count + 1,1).Value() = $person
			$staff_list.Cells.Item($staff_count + 1,2).Value() = $new_server
			Write-Host "$person will now be notified about $new_server"
		}
		#Person exists check if server is already in person's list, if not add it
		else{
			$staff_range = $staff_list.UsedRange.Cells
			$column_count = $staff_range.Columns.Count
			for ($col = 2; $col -le $column_count; $col++){
				#Check if cell is empty, if not check and see if it is a match
				if(-Not [string]::IsNullOrEmpty($staff_list.Cells.Item($staff_row, $col).Value())){
					if($staff_list.Cells.Item($staff_row, $col).Value().ToString() -eq $new_server.ToString()){
						$match = 1
						Write-Host "User $person is already set to be notified about $new_server"
					}
					if($match -eq 1){
						Break
					}
					else{
						$count++
					}
				}
			}
			#cell is empty add the server to that person's list
			if(($match -ne 1) ){
				$staff_list.Cells.Item($staff_row,$count+1).Value() = $new_server
				Write-Host "$person will now be notified about $new_server"
			}
		}
	}
	$staff_server_list.Save()
	$staff_server_list.Close()
	$excel.Quit()
}
#----------------------------------------------------------------------
#Get Server Information: Name, critical, AppDB, staff, and entry type - manually or through CSV
#----------------------------------------------------------------------
Write-Host "CommVault Script Server Updater"
$file_path = Get-Location
$more_updates = 1
$application = "Unknown"
$csv = Get-UserInput -Question "Update with the Staff Server List CSV?    :"
If($csv -eq 0){
	While($more_updates -eq 1){
		$staff = @()
		$delete = Get-UserInput -Question "Are you deleting an entry?                :"
		If($delete -eq 1){
			$remove_server = Read-Host "what is the name of the server to delete? :"
			Update-RemoveServerRecord $remove_server $file_path
			Update-RemoveServerNotice $remove_server $file_path
		}
		Else{
			$staffmod = Get-UserInput -Question "Are you making a change to notifications? :"
			$new_server = Read-Host "What is the name of the server?           :"
			If([string]::IsNullOrEmpty($new_server.Trim())){
				Write-Host "No server entered"
				Continue
			}
			Else{
				$new_server = $new_server.ToUpper()
			}
			If($staffmod -eq 0){
				$critical = Get-UserInput -Question "Is this a critical server? 		  :" 
				$appdb = Get-UserInput -Question "Is this an application database?  	  :"
				$application = Read-Host "What application uses this server?        :"
			}
			Else{
				$add_remove = Get-UserInput -Question "Are you adding a person?                  :"
			}
			$new_staff = Read-Host "Email address of staff                    :"
			$new_staff = $new_staff.split("@")
			$staff += $new_staff[0]
			$decision = 1
			While($decision -eq 1){
				$response = Read-Host "Does anyone else need to be updated?      :"
				$response = $response.ToLower()
				$response_split = $response.substring(0,1)
				If(($response_split -ne 'y') -and ($response_split -ne 'n')){
					While(($response_split -ne 'y') -and ($response_split -ne 'n')){
						$response = Read-Host 'Just a simple yes or no will do nicely thank you'
						$response = $response.ToLower()
						$response_split = $response.substring(0,1)
					}
				}
				If($response_split -eq 'y'){
					$new_staff = Read-Host "Email address of staff                    :"
					$new_staff = $new_staff.split("@")
					$staff += $new_staff[0]
				}
				Else{
					$decision = 0
				}
			}
			If($staffmod -eq 0){
				#Insert new server into script DB and update staff_list file for manual entries
				Add-NewServerEntry $new_server $critical $appdb $file_path $application
				Update-StaffToNotify $staff $new_server $file_path
			}
			Else{
				If($add_remove -eq 1){
					Update-StaffToNotify $staff $new_server $file_path
				}
				Else{
					Update-StaffMod $staff $new_server $file_path
				}
			}
		}
		$more_updates = Get-UserInput -Question "Would you like to make anymore updates?   :"
	}
}
Else{
	#update the script with a CSV
	ForEach($line in Get-Content (Join-Path $file_path "ServerUpdates.txt")| select-object -skip 7){
		$staff = @()
		$line = $line.Split(",")
		If([string]::IsNullOrEmpty($line)){
			Continue
		}
		$entry_type = $line[0].Trim().ToLower()
		$new_server = $line[1]
		#check if a server was provided
		If([string]::IsNullOrEmpty($new_server.Trim())){
			Write-Host "No server was provided"
			Continue
		}
		Else{
			$new_server = $line[1].Trim().ToUpper()
		}
		#delete a server
		If($entry_type -eq 'delete'){
			Update-RemoveServerRecord $new_server $file_path
			Update-RemoveServerNotice $new_server $file_path
			Continue
		}
		#make staff changes to notifications
		ElseIf($entry_type -eq 'staffmod'){
			$add_remove = $line[2].Trim()
			If(($add_remove -ne 1) -and ($add_remove -ne 0)){
				Write-Host "No safe default can be assumed skipping entry"
				Continue
			}
			For($i = 3; $i -le $line.Count -1; $i++){
				$new_staff = $line[$i].split("@")
				$staff += $new_staff[0].Trim()
			}
			If($staff.Count -gt 0){
				#Add staff to notifications
				If($add_remove -eq 1){
					Update-StaffToNotify $staff $new_server $file_path
				}
				#remove staff from notifications
				Else{
					Update-StaffMod $staff $new_server $file_path 
				}
			}
			Else{
				Write-Host "No staff were listed in the entry"
			}
			Continue
		}
		#entry type is neither staffmod nor delete default to new
		Else{
			$critical = $line[2].Trim()
			If(($critical -ne 1) -and ($critical -ne 0)){
				Write-Host "Critical was not 1 or 0. Defaulting to 1 critical server"
				$critical = 1
			}
			$appdb = $line[3].Trim()
			If(($appdb -ne 1) -and ($appdb -ne 0)){
				Write-Host "Application Database was not 1 or 0. Defaulting to 0 don't notify DRM"
				$appdb = 0
			}
			If(![string]::IsNullOrEmpty($line[4])){
				$application = $line[4]
			}
			else{
				$application = "Unknown"
			}
			For($i = 5; $i -le $line.Count -1; $i++){
				$new_staff = $line[$i].split("@")
				$staff += $new_staff[0].Trim()
			}
			Add-NewServerEntry $new_server $critical $appdb $file_path $application
			If($staff.Count -gt 0){
				Update-StaffToNotify $staff $new_server $file_path
			}
			Else{
				Write-Host "No staff were listed in the file"
			}
		}
	}
}
#----------------------------------------------------------------------
#Cleaning up Excel so there are not a bunch of excel tasks left running
#----------------------------------------------------------------------
[GC]::Collect()