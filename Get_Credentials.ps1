<#
-----------------------------------------------------------------------
Name: Get Credentials
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

function Get-MyCredentials(){
	[cmdletbinding()]
	param(
		[string]$file_path = "C:\Development\Helper_Scripts\Password_Files",
		[string]$user_name = "username",
		[string]$file_name = "password.txt",
		[switch] $manual
	)
	# Ask user to pick a folder using window popup
	if($manual){
		$file_path = Select-Path
	}
	
	$password = get-content "$file_path\$file_name" | convertto-securestring
	$credentials = new-object System.Management.Automation.PSCredential($user_name, $password)
	Return $credentials
	
}
