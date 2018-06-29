<#
-----------------------------------------------------------------------
Name: CommVault Updater Add Staff
Author: Anthony Dunaway
Date: 06/26/18
Updated: 06/28/18
Description:
Adds new staff members to notifications in ServerList.xlsx
-----------------------------------------------------------------------
#>
function Update-NotifyStaff{
	[CmdletBinding()]
	param(
		[string[]]$staff, 
		[string]$server, 
		[string]$file_path
	)
	Write-Verbose "Updating Staff Notifications"
	$excel = New-Object -comobject Excel.Application -Verbose:$false
	$excel.Visible = $False
	$excel.DisplayAlerts = $False
	$server_db = $excel.Workbooks.Open("$file_path\ServerList.xlsx")
	$servers = $server_db.worksheets.item(1)
	$server_count = ($servers.UsedRange.Rows).Count
	$updated_staff = ""
	$new_staff_list = @()
	$update_row = -999
	for($row = 2; $row -le $server_count; $row++){
		if(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, "A").Value())){
			$current_server = $servers.Cells.Item($row, "A").Value().ToString()
			if($current_server -eq $server){
				$update_row = $row
				if(-Not [string]::IsNullOrWhitespace($servers.Cells.Item($row, "H").Value())){
					$staff_list = $servers.Cells.Item($row, "H").Value().ToString().Split(",")
					$servers.Cells.Item($row, "H").Value() = ""
					$new_staff_list = $staff_list
					foreach($person in $staff){
						foreach($human in $staff_list){
							$found = 0
							if($person -eq $human){
								$found = 1
								Write-Verbose "$person is already being notified about $server "
								Break
							}
						}
						if($found -eq 0){
							$new_staff_list += $person
							Write-Verbose "$person is now being notified about $server "
						}
					}
				}
				else{
					foreach($person in $staff){
						$new_staff_list += $person
						Write-Verbose "$person is now being notified about $server "
					}
				}
			}
		}
	}
	if($update_row -ge 0){
		foreach($person in $new_staff_list){
			$updated_staff += "$person,"
		}
		$updated_staff = $updated_staff.Substring(0,$updated_staff.length - 1).Trim()
		Write-Verbose "The staff being notified about $server are $updated_staff"
		$servers.Cells.Item($update_row, "H").Value() = $updated_staff
		$server_db.save()
	}
	else{
		Write-Verbose "$server not found"
	}
	$server_db.close()
}

#Code to test the script
# $file_path = Get-Location
# $staff1 = @("ellory.pandalover", "coraline.jones", "melchior.pandaking")
# $staff2 = @("melchior.pandaking","holly.pandaqueen")
# $staff3 = @("ophelia.pandaprincess")
# $staff4 = @("melchior.pandaking")

# Update-NotifyStaff -staff $staff1 -server "happypanda" -file_path $file_path -verbose
# Update-NotifyStaff -staff $staff2 -server "happypanda" -file_path $file_path -verbose
# Update-NotifyStaff -staff $staff1 -server "happypanda" -file_path $file_path -verbose
# Update-NotifyStaff -staff $staff3 -server "happypanda" -file_path $file_path -verbose
# Update-NotifyStaff -staff $staff4 -server "happypanda" -file_path $file_path -verbose