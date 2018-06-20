<#
-----------------------------------------------------------------------
Name: Get Server Owner List
Author:	Anthony Dunaway
Date: 06/15/18
Updated: 06/18/18
Description:
Helper script to get a list of servers and their staff. 
Returns a hash table with the staff member as the key and 
a list of their servers as the value.
-----------------------------------------------------------------------
#>
Function Get-StaffServers{
	[CmdletBinding()]
	param()
	Write-Verbose "Opening Excel"
	$file_path = Get-Location
	$excel = New-Object -comobject Excel.Application -Verbose:$false
    $excel.DisplayAlerts = $False
	Write-Verbose "Opening the list of staff"
	$staff_server_list = $excel.Workbooks.Open("$file_path\StaffServerList.xlsx" )
	$staff = $staff_server_list.worksheets.item(1)
	$staff_list = @{}
	$staff_count = ($staff.UsedRange.Rows).count
	Write-Verbose "Number of staff Found : $staff_count"
	For($row = 2; $row -le $staff_count; $row++){
		$current_owner = $staff.Cells.Item($row, "A").Value().ToString()
		Write-Debug "Current owner : $current_owner"
		$current_servers = @()
		$staff_range = $staff.UsedRange.Cells
		$column_count = $staff_range.Columns.Count
		For ($col = 2; $col -le $column_count; $col++){
			If(-Not [string]::IsNullOrEmpty($staff.Cells.Item($row, $col).Value())){
				$current_servers += $staff.Cells.Item($row, $col).Value()
				Write-Debug "Current server is $staff.Cells.Item($row, $col).Value()"
			}
		}
		Write-Debug "Adding $current_owner to the list"
		$staff_list.add($current_owner, $current_servers)
	}
	Write-Verbose "Closing Excel and Cleaning Up"
	$excel.Workbooks.Close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($staff)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($staff_server_list)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
	Write-Verbose "List Complete"
	Return $staff_list
}