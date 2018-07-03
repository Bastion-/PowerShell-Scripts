<#
-----------------------------------------------------------------------
Name: Heads Up Script CAB Agenda Parser
Author:	Anthony Dunaway
Date: 06/22/18
Updated: 07/03/18
Description:
Parses CAB agenda emails looking for RFC's relevant to our team
-----------------------------------------------------------------------
#>

."C:\Development\Heads_Up\S3_Scraper.ps1"

$web_file_path = Get-Location
$processing_file = "process.txt"
$web_text = "web.txt"

function Search-CabAgenda{
	[cmdletbinding()]
	param(
		[string]$file_path
	)
	Write-Verbose "Checking the CAB agenda for relevant RFC's"
	$rfc = ""
	$content = @()
	$file_reader = [System.IO.File]::OpenText($file_path)
	$file_line = "start"
	while($file_line -ne $null){
		$file_line = $file_reader.ReadLine()
		$content += $file_line
	}
	for($i = 0; $i -le $content.count - 2; $i++){
		if($content[$i][0] -eq "4"){
			$rfc = $content[$i]
			$rfc_end = "--------------------------------------------------END OF RFC $rfc--------------------------------------------------"
			for($j = $i + 1; $j -le $content.count - 2; $j++){
				if($content[$j][0] -ne "4"){
					if($content[$j] -like "*place*"){
						Write-Verbose "$rfc affects place"
						Get-S3Data -rfc $rfc
						Add-Content "$web_file_path\$web_text" $rfc_end
						$i = $j - 1
						Break
					}
				}
				else{
					$i = $j - 1
					Break
				}
			}
		}
	}
}

