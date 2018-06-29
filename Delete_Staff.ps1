<#
-----------------------------------------------------------------------
Name: CommVault Updater Delete Staff
Author: Anthony Dunaway
Date: 06/26/18
Updated: 06/28/18
Description:
Removes staff members from notifications in ServerList.xlsx
-----------------------------------------------------------------------
#>
function Remove-DeleteStaff{
	[CmdletBinding()]
	param(
		[string[]]$staff, 
		[string]$server, 
		[string]$file_path
	)
	Write-Verbose "Removing $staff from notifications"
	$excel = New-Object -comobject Excel.Application -Verbose:$false
	$excel.DisplayAlerts = $False
	$server_db = $excel.Workbooks.Open("$file_path\ServerList.xlsx")
	$servers = $server_db.worksheets.item(1)
	$server_count = ($servers.UsedRange.Rows).Count
	$update_row = -999
	foreach($person in $staff){
		$updated_staff = ""
		Write-Verbose "Removing $person"
		$removed = 0
		for($row = 2; $row -le $server_count; $row++){
			if(-Not [string]::IsNullOrEmpty($servers.Cells.Item($row, 1).Value())){
				$current_server = $servers.Cells.Item($row, "A").Value().ToString()
				if($current_server -eq $server){
					$update_row = $row
					if(-Not[string]::IsNullOrWhitespace($servers.Cells.Item($row, "H").Value())){
						$staff_list = $servers.Cells.Item($row, "H").Value().ToString().Split(",")
						foreach($human in $staff_list){
							Write-Verbose "Human is $human"
							if($human -ne $person){
								Write-Verbose "No Match"
								$updated_staff += "$human,"
							}
							else{
								$removed = 1
							}
						}
						if($removed -eq 1){
							$servers.Cells.Item($row, "H").Value() = ""
							if($updated_staff.Length -gt 0){
								$updated_staff = $updated_staff.Substring(0, $updated_staff.Length - 1).Trim()
							}
							$servers.Cells.Item($update_row, "H").Value() = $updated_staff
							Write-Verbose "$person was removed from $server"
							Break
						}
					}
				}
			}
		}
	}		
	#This removes any empty rows left behind after deleting a server
	$server_db.Save()
	$server_db.Close()
	$excel.Quit()
	if($update_row -ge 0){
		if($removed -eq 1){
			Write-Verbose "$staff removed"
		}
		else{
			Write-Verbose "$staff not found"
		}
	}
	else{
		Write-Verbose "$server not found"
	}
}

#Testing code
# $file_path = Get-Location
# $staff1 = @("ellory.pandalover", "coraline.jones", "melchior.pandaking")
# $staff2 = @("melchior.pandaking","holly.pandaqueen")
# $staff3 = @("ophelia.pandaprincess")
# $staff4 = @("melchior.pandaking")

# Remove-DeleteStaff -staff $staff1 -server "happypanda" -file_path $file_path -verbose
# Remove-DeleteStaff -staff $staff2 -server "happypanda" -file_path $file_path -verbose
# Remove-DeleteStaff -staff $staff3 -server "happypanda" -file_path $file_path -verbose
# Remove-DeleteStaff -staff $staff4 -server "happypanda" -file_path $file_path -verbose