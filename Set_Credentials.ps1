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

function Select-Path(){
	
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	
    $path_name = New-Object System.Windows.Forms.FolderBrowserDialog
    $path_name.Description = "Select a folder"
    $path_name.rootfolder = "MyComputer"

    if($path_name.ShowDialog() -eq "OK")
    {
        $path += $path_name.SelectedPath
    }
    return $path
	
}

function Set-Credentials(){
	[cmdletbinding()]
	param(
		$file_name = "password.txt"
	)
	
	$file_path = Select-Path
	read-host "Password" -assecurestring | convertfrom-securestring | out-file "$file_path\$file_name"
}
