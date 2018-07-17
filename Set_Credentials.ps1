<#
-----------------------------------------------------------------------
Name: Set Credentials
Author:	Anthony Dunaway
Date: 07/16/18
Updated: 07/17/18
Description:
Create a PSCredential Object from secure file
-----------------------------------------------------------------------
#>

."C:\Development\Helper_Scripts\Select_Path.ps1"

function Set-Credentials(){
	[cmdletbinding()]
	param(
		$file_name = "password.txt"
	)
	
	$file_path = Select-Path
	read-host "Password" -assecurestring | convertfrom-securestring | out-file "$file_path\$file_name"
}