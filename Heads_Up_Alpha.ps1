<#
-----------------------------------------------------------------------
Name: Heads Up Script
Author:	Anthony Dunaway
Date: 06/21/18
Updated: 07/03/18
Description:
Monitors various information sources to try and get ahead of
possible issues affecting public health systems. Emails SME's 
regarding events that could affect availability of their applications
and systems.
-----------------------------------------------------------------------
#>
#----------------------------------------------------------------------------------------------------------------------
#Script Imports
#----------------------------------------------------------------------------------------------------------------------
."C:\Development\Heads_Up\Cab_Agenda.ps1"
."C:\Development\Heads_Up\File_Cleaner.ps1"
."C:\Development\Heads_Up\Hide_Window.ps1"
."C:\Development\Heads_Up\Morning_Updates.ps1"
."C:\Development\Heads_Up\S3_Scraper.ps1"
."C:\Development\Heads_Up\Subject_Parser.ps1"

#----------------------------------------------------------------------------------------------------------------------
#Hide Window
#----------------------------------------------------------------------------------------------------------------------
# $consolePtr = [Console.Window]::GetConsoleWindow()
# [Console.Window]::ShowWindow($consolePtr, 0)

#----------------------------------------------------------------------------------------------------------------------
#Get the list of servers 
#----------------------------------------------------------------------------------------------------------------------
# $server_list = Get-ServerList -Verbose:$false

#----------------------------------------------------------------------------------------------------------------------
#Script Variables
#----------------------------------------------------------------------------------------------------------------------
$file_path = Get-Location
$processing_file = "process.txt"
$csv = "Rfc_cleaner.csv"
$web_file = "web.txt"

#----------------------------------------------------------------------------------------------------------------------
#Cleanup Old Files
#----------------------------------------------------------------------------------------------------------------------
if(Test-Path $file_path\$processing_file){
	Remove-Item $file_path\$processing_file
}
if(Test-Path $file_path\$web_file){
	Remove-Item $file_path\$web_file
}

#----------------------------------------------------------------------------------------------------------------------
#Open Outlook 
#----------------------------------------------------------------------------------------------------------------------
$outlook = new-object -comobject "Outlook.Application"
$mapi = $outlook.getnamespace("mapi")
$desktop_path = [Environment]::GetFolderPath("Desktop")
$inbox = $mapi.GetDefaultFolder(6)

# while(1){
	#Number of seconds between each check
	# sleep -s 1
#----------------------------------------------------------------------------------------------------------------------
#Loop through change management emails
#----------------------------------------------------------------------------------------------------------------------
	$change_notifications = $inbox.Folders | where-object {$_.name -eq "Change Management"}
	foreach($notice in $change_notifications.Items){
		if($notice.UnRead -eq $true){
			$subject = $notice.Subject
			$type = Get-MessageType -subject $subject
			if($type -like "RFC*"){
				$rfc = $type.Substring(3)
				Get-S3Data -rfc $rfc
			}
			elseif($type -eq "morning"){
				$content= $notice.body
				Add-Content $file_path\$processing_file $content
				(Get-Content $file_path\$processing_file) | ? {$_.trim() -ne "" } | Set-Content $file_path\$processing_file
				Search-MorningUpdate -file_path $file_path\$processing_file -verbose
			}
			elseif($type -eq "CAB"){
				$content= $notice.body
				Add-Content $file_path\$processing_file $content
				(Get-Content $file_path\$processing_file) | ? {$_.trim() -ne "" } | Set-Content $file_path\$processing_file
				Search-CABAgenda -file_path $file_path\$processing_file
			}
		}
	}
	Optimize-CleanFile -file $file_path\$web_file -csv $file_path\$csv
	(Get-Content $file_path\$web_file) | ? {$_.trim() -ne "" } | Set-Content $file_path\$web_file
#}
Read-Host -prompt "Exit"


















