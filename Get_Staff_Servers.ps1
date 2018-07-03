<#
-----------------------------------------------------------------------
Name: Get Server Owner List
Author:	Anthony Dunaway
Date: 06/15/18
Updated: 06/29/18
Description:
Creates a CSV of each team member and their servers
-----------------------------------------------------------------------
#>
[CmdletBinding()]
param(
	[string] $export_path = "none"
)

$file_path = "I:\ISE\CommVault\CommVaultScript"

Write-Verbose "Opening Excel"
$excel = New-Object -comobject Excel.Application -Verbose:$false
$excel.DisplayAlerts = $False
Write-Verbose "Opening the server list"
$server_list = $excel.Workbooks.Open("$file_path\ServerList.xlsx" )
$worksheet = $server_list.worksheets.item(1)
$staff_list = @{}
$num_rows = ($worksheet.UsedRange.Rows).Count
#Get a list of all the staff
for($row = 2; $row -le $num_rows; $row++){
	if(-Not[string]::IsNullOrWhitespace($worksheet.Cells.Item($row, 'H').Value())){
		$current_staff = $worksheet.Cells.Item($row, 'H').Value().ToString().Trim()
	}
	$all_staff += "$current_staff,"
}
$all_staff = $all_staff.Split(",")
$all_staff = $all_staff | Select-Object -Unique
#find out which servers belong to each staff
foreach($person in $all_staff){
	Write-Verbose "Finding the servers for $person "
	$servers = @()
	for($row = 2; $row -le $num_rows; $row++){
		if(-Not[string]::IsNullOrWhitespace($worksheet.Cells.Item($row, 'H').Value())){
			$current_staff = $worksheet.Cells.Item($row, 'H').Value().ToString().Trim().Split(",")
			if($current_staff.Contains($person)){
				$servers += $worksheet.Cells.Item($row, 'A').Value().ToString().Trim()
			}
		}
	}
	$servers = $servers -join ","
	$staff_list.add($person, $servers)
}
Write-Verbose "List Complete"
Write-Verbose "Closing Excel and Cleaning Up"
#close Excel
$server_list.Close()
$excel.Quit()
#Export hash table to CSV at the specified directory
if($export_path -ne "none"){
	Write-Verbose "Exporting list to CSV at $export_path"
	$staff_list.GetEnumerator() | Select-Object -Property Key,Value | Export-Csv -NoTypeInformation -path "$export_path\staff_servers.csv"
}
#no directory provided export to location of serverlist.xlsx
else{
	Write-Verbose "Exporting list to CSV at $file_path"
	$staff_list.GetEnumerator() | Select-Object -Property Key,Value | Export-Csv -NoTypeInformation -path "$file_path\staff_servers.csv"
}

