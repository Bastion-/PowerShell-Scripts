<#
-----------------------------------------------------------------------
Name: CommVault Backup Report Script Server Updater
Author: Anthony Dunaway
Date: 03/15/18
Updated: 06/28/18
Description:
Adds a new server, deletes a server, or updates a current server. 
Servers can be entered individually or in bulk via a csv file. 
Users will need to know the server name, whether it is a critical server, 
if it is an application database, and the email address of the staff to be 
notified. if the staff are unknown the file will be added to tracking but 
not to email notifications.
-----------------------------------------------------------------------
#>

."C:\Development\CommVaultScript\Helper_Scripts\Add_Staff.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\Delete_Server.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\Delete_Staff.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\New_Server.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\User_Input.ps1"

#-----------------------------------------------------------------------------------------------------------------------------
#Get Server Information: Name, critical, AppDB, staff, and entry type - manually or through CSV
#-----------------------------------------------------------------------------------------------------------------------------
Write-Host "CommVault Script Server Updater"
$file_path = Get-Location
$more_updates = 1
$application = "Unknown"
$csv = Get-UserInput -Question "Update with the Staff Server List CSV?    :"
if($csv -eq 0){
	While($more_updates -eq 1){
		$staff = @()
		$delete = Get-UserInput -Question "Are you deleting an entry?                :"
		if($delete -eq 1){
			$remove_server = Read-Host "what is the name of the server to delete? :"
			Remove-DeleteServer -server $remove_server -file_path $file_path
		}
		else{
			$staffmod = Get-UserInput -Question "Are you making a change to notifications? :"
			if($staffmod -eq 1){
				$add_remove = Get-UserInput -Question "Are you adding a person?                  :"
			}
			$new_server = Read-Host "What is the name of the server?           :"
			if([string]::IsNullOrEmpty($new_server.Trim())){
				Write-Host "No server entered"
				Continue
			}
			else{
				$new_server = $new_server.ToUpper()
			}
			if($staffmod -eq 0){
				$critical = Get-UserInput -Question "Is this a critical server? 		  :" 
				$appdb = Get-UserInput -Question "Is this an application database?  	  :"
				$application = Read-Host "What application uses this server?        :"
				Add-NewServer -new_server $new_server -critical $critical -appdb $appdb -file_path $file_path -application $application
				$add_person = Get-UserInput -Question "Are you adding a person?                  :"
				if($add_person -eq 1){
					$staffmod = 1
					$add_remove = 1
				}
			}
			if($staffmod -eq 1){
				if($add_remove -eq 1){
					$new_staff = Read-Host "Email address of staff                    :"
				}
				else{
					$new_staff = Read-Host "Email address of staff to remove          :"
				}
				$new_staff = $new_staff.split("@")
				$staff += $new_staff[0]
				$decision = 1
				While($decision -eq 1){
					$response = Get-UserInput -Question "Does anyone else need to be updated?      :"
					if($response -eq 1){
						$new_staff = Read-Host "Email address of staff                    :"
						$new_staff = $new_staff.split("@")
						$staff += $new_staff[0]
					}
					else{
						$decision = 0
					}
				}
				if($add_remove -eq 1){
					Update-NotifyStaff -staff $staff -server $new_server -file_path $file_path
				}
				else{
					Remove-DeleteStaff -staff $staff -server $new_server -file_path $file_path
				}
			}
		}
		$more_updates = Get-UserInput -Question "Would you like to make anymore updates?   :"
	}
}
else{
	#update the script with a CSV
	forEach($line in Get-Content (Join-Path $file_path "ServerUpdates.txt")| select-object -skip 7){
		$staff = @()
		$line = $line.Split(",")
		if([string]::IsNullOrEmpty($line)){
			Continue
		}
		$entry_type = $line[0].Trim().ToLower()
		$new_server = $line[1]
		#check if a server was provided
		if([string]::IsNullOrEmpty($new_server.Trim())){
			Write-Host "No server was provided"
			Continue
		}
		else{
			$new_server = $line[1].Trim().ToUpper()
		}
		#delete a server
		if($entry_type -eq 'delete'){
			Remove-DeleteServer -server $new_server -file_path $file_path
			Continue
		}
		#make staff changes to notifications
		elseif($entry_type -eq 'staffmod'){
			$add_remove = $line[2].Trim()
			if(($add_remove -ne 1) -and ($add_remove -ne 0)){
				Write-Host "No safe default can be assumed skipping entry"
				Continue
			}
			for($i = 3; $i -le $line.Count -1; $i++){
				$new_staff = $line[$i].split("@")
				$staff += $new_staff[0].Trim()
			}
			if($staff.Count -gt 0){
				#Add staff to notifications
				if($add_remove -eq 1){
					Update-NotifyStaff -staff $staff -server $new_server -file_path $file_path
				}
				#remove staff from notifications
				else{
					Remove-DeleteStaff -staff $staff -server $new_server -file_path $file_path
				}
			}
			else{
				Write-Host "No staff were listed in the entry"
			}
			Continue
		}
		#entry type is neither staffmod nor delete default to new
		else{
			$critical = $line[2].Trim()
			if(($critical -ne 1) -and ($critical -ne 0)){
				Write-Host "Critical was not 1 or 0. Defaulting to 1 critical server"
				$critical = 1
			}
			$appdb = $line[3].Trim()
			if(($appdb -ne 1) -and ($appdb -ne 0)){
				Write-Host "Application Database was not 1 or 0. Defaulting to 0 don't notify DRM"
				$appdb = 0
			}
			if(![string]::IsNullOrEmpty($line[4])){
				$application = $line[4]
			}
			else{
				$application = "Unknown"
			}
			for($i = 5; $i -le $line.Count -1; $i++){
				$new_staff = $line[$i].split("@")
				$staff += $new_staff[0].Trim()
			}
			Add-NewServer -new_server $new_server -critical $critical -appdb $appdb -file_path $file_path -application $application
			if($staff.Count -gt 0){
				Update-NotifyStaff -staff $staff -server $new_server -file_path $file_path
			}
			else{
				Write-Host "No staff were listed in the file"
			}
		}
	}
}
#-----------------------------------------------------------------------------------------------------------------------------
#Cleaning up Excel so there are not a bunch of excel tasks left running
#-----------------------------------------------------------------------------------------------------------------------------
[GC]::Collect()
