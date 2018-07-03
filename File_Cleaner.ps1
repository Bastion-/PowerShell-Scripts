<#
-----------------------------------------------------------------------
Name: File Content Cleaner
Author:	Anthony Dunaway
Date: 07/02/18
Updated: 07/02/18
Description:
Replaces the values in a file using a CSV.
-----------------------------------------------------------------------
#>
function Optimize-CleanFile {
	[cmdletbinding()]
	param(
		[string] $file,
		[string] $csv
	)
	$replacement_list = Import-Csv $csv
	Get-ChildItem $file | Out-Null
	$content = Get-content $file;
	foreach ($value in $replacement_list){
		$content = $content.Replace($value.OldValue, $value.NewValue)
	}
	Set-content $file -Value $content
	(gc $file) | ? {$_.trim() -ne "" } | set-content $file
}