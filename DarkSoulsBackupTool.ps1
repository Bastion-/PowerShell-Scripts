<#
-----------------------------------------------------------------------
Name:  	 Dark Souls 3 Save Backup Utility
Author:	 Anthony Dunaway
Date:    07/01/18
Description:
Creates save backups of either DS3 or DS2.  
-----------------------------------------------------------------------
#>

function Get-UserInput{
	param([string]$Question)
	[string]$response = Read-Host $Question 
	$response = $response.ToLower()
	$response_split = $response.substring(0,1)
	If(($response_split -ne 'y') -and ($response_split -ne 'n')){
		While(($response_split -ne 'y') -and ($response_split -ne 'n')){
			$user_response = Read-Host 'Just a simple yes or no will do nicely thank you'
			$user_response = $user_response.ToLower()
			$response_split = $user_response.substring(0,1)
		}
	}
	If($response_split -eq 'y'){
		$decision = 1
	}
	Else{
		$decision = 0
	}
	return [int]$decision
}

$ds2 = Get-UserInput -Question "Default is DS3. Do you want to backup DS2 saves instead?"

if($ds2 -eq 0){
	#Backup DS3 saves
	#Full path to Dark Souls 3 saves
	$save_path = "C:\Users\Anthony\AppData\Roaming\DarkSoulsIII\01100001051dfe45\"
	#Full path to where you would like the backups saved
	$backup_path = "C:\Users\Anthony\Documents\DarkSouls3Backups\"
	#Name of the Dark Souls 3 save file - should be constant
	$file_name = "DS30000.sl2"
}
else{
	#Backup DS2 saves
	$save_path = "C:\Users\Anthony\AppData\Roaming\DarkSoulsII\01100001051dfe45\"
	$backup_path = "C:\Users\Anthony\Documents\DarkSouls2Backups\"
	$file_name = "DS2SOFS0000.sl2"
}
#Number of backup files you would like
$concurrent_backups = Read-Host "How many concurrent backups would you like? "
$minutes = Read-Host "How many minutes between backups? "
$minutes = [int]$minutes * 60 
$current_backups = @()
Remove-Item "$backup_path$file_name*" | Where { ! $_.PSIsContainer}
While(1){
	$date = Get-Date -UFormat "-%A-%H.%M"
	Copy-Item "$save_path$file_name" -Destination "$backup_path$file_name$date"
	$current_backups += ("$backup_path$file_name$date")
	
	If($current_backups.Count -gt [int]$concurrent_backups){
		Remove-Item $current_backups[0]
		$current_backups = $current_backups[1..($current_backups.Length-1)]
	}
	Write-Host "Backup Created"
	#Number of seconds to wait between each backup - 600 seconds is 10 minutes
	Start-Sleep -s $minutes
}