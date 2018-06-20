<#
-----------------------------------------------------------------------
Name:  Oregon Death Record Blob Grabber
Author:	Anthony Dunaway
Date:    05/09/2018
Updated: 06/08/2018
Description:
Takes a screencapture of the current death record. 
-----------------------------------------------------------------------
#>
$location = Get-Location
."$location\Mouse_Click.ps1"
."$location\Auto_Screenshot.ps1"

Write-Host "Blob Grabber"
$number_of_records = Read-Host "How many records would you like to process?	"
$number_of_records = [int]$number_of_records
$offset = Read-Host "What offset would you like to use?              "
$offset = [int]$offset
$file_name = "record"
$save_path = "$location\captures\"
$bounds = [Drawing.Rectangle]::FromLTRB(1988, 137, 2458, 435)
[Clicker]::LeftClickAtPoint(2120,100)
$start_time = Get-Date
for($record = 1; $record -le $number_of_records; $record ++){
	$number_offset = $offset + $record
	screenshot $bounds "$save_path$number_offset.jpg"
	[Clicker]::LeftClickAtPoint(1950,250)
	Start-Sleep -m 75
}
$end_time = Get-Date
$seconds = ($end_time - $start_time).TotalSeconds
Write-Host "Total time taken for $number_of_records records was $seconds seconds"
Read-Host "Press Enter to Exit"
