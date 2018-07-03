<#
-----------------------------------------------------------------------
Name: Heads Up Script Morning Update Parser
Author:	Anthony Dunaway
Date: 06/22/18
Updated: 07/02/18
Description:
Searches the RFC's in the morning update emails
-----------------------------------------------------------------------
#>

."C:\Development\Heads_Up\S3_Scraper.ps1"

$web_file_path = Get-Location
$processing_file = "process.txt"
$web_text = "web.txt"

function Search-MorningUpdate{
	[cmdletbinding()]
	param(
		[string] $file_path
	)
	Write-Verbose "Checking the Morning Update for RFC's"
	$date = ""
	$rfc = ""
	$status = ""
	$result = @()
	$weekdays = @("Mon", "Tue", "Wed", "Thu", "Fri","Sat", "Sun")
	$date_line = 0
	$file_reader = [System.IO.File]::OpenText($file_path)
	
	for(){
		$file_line = $file_reader.ReadLine()
		if ($file_line -eq $null) { 
			break 
		}
		$split_line = $file_line.Split(" ")
		foreach($day in $weekdays){
			if($file_line.Substring(0,3) -eq $day){
				$date = $file_line
				Write-Verbose "The date is $date"
				$date_line = 1
				Break
			}
		}
		if($date_line -eq 1){
			$date_line = 0
			Continue
		}
		elseif ($split_line[0] -ne "*"){
			Continue
		}
		else{
			foreach($word in $split_line){
				if($word -like "RFC*"){
					$rfc = $word.Substring(3)
					$rfc_end = "--------------------------------------------------END OF RFC $rfc--------------------------------------------------"
					Write-Verbose "Found the RFC $rfc"
					Get-S3Data -rfc $rfc
				}
			}
			if($split_line[$split_line.Count - 1].Substring(0,1) -eq "("){
				$status = $split_line[$split_line.Count - 1]
				Write-Verbose "Status of RFC is $status"
				Add-Content "$web_file_path\$web_text" $status
				Add-Content "$web_file_path\$web_text" $rfc_end
			}
			elseif($split_line[$split_line.Count - 3].Substring(0,3) -eq "New"){
				$status = $split_line[$split_line.Count -4] + " " + $split_line[$split_line.Count - 1]
				Write-Verbose "Status of RFC is $status"
				Add-Content "$web_file_path\$web_text" $status
				Add-Content "$web_file_path\$web_text" $rfc_end
			}
		}
	}
	$file_reader.Close()
}