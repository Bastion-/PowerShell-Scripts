<#
-----------------------------------------------------------------------
Name: CommVault Helper: Move Columns
Author: Anthony Dunaway
Date: 07/10/18
Updated: 07/10/18
Description:
Moves the from column to the to column
-----------------------------------------------------------------------
#>
function Move-Columns(){
	[cmdletbinding()]
	param(
	$worksheet,
	[string]$from,
	[string]$to
	)	

	$shift_right = -4161
	[void]$worksheet.cells.entireColumn.Autofit()
	[void]$worksheet.cells.entireRow.Autofit()
	$critical_column = $worksheet.range($from + "1:" + $from + "1").EntireColumn
	$critical_column.Copy() | out-null
	$new_column = $worksheet.Range("$to"+"1").EntireColumn
	$new_column.Insert($shift_right) | out-null
	[void]$critical_column.Delete()
}