<#
-----------------------------------------------------------------------
Name: CommVault Backup Report Script
Author: Anthony Dunaway
Date: 01/16/18
Updated: 07/19/18
Description:
This script gets all of the CommVault backup reports from the user's outlook inbox,
copies our servers onto a new worksheet, and records missing and failed backups to their own
respective worksheets. Once that is done it saves the file to sharepoint and deletes the
temp files from the desktop. After that it will check if any servers failed and email the responsible parties.
It marks each email as read after it is processed.
This script was written for Powershell 5.0. Currently it does not operate correctly on
Powershell 2.0 which is the Windows 7 default. The script is designed to run silently in the background.
for best results schedule the script to run daily via the Windows task scheduler.
-----------------------------------------------------------------------
#>
param(
	[switch] $debug,
	[switch] $verbose
)
#------------------------------------------------------------------------------------------------------------------------------
# Imports
#------------------------------------------------------------------------------------------------------------------------------
$file_path = $PSScriptRoot.ToString()
."$file_path\Helper_Scripts\Get_Server_Objects.ps1"
."$file_path\Helper_Scripts\Hide_Window.ps1"
."$file_path\Helper_Scripts\Swap_Columns.ps1"
."$file_path\Helper_Scripts\User_Input.ps1"

#------------------------------------------------------------------------------------------------------------------------------
# Debug menu options
#------------------------------------------------------------------------------------------------------------------------------
if($debug -eq $true){
	Write-Host "RUNNING IN DEBUG MODE"
	#Decide to mark emails as read or not
	$mark_read = Get-UserInput -Question 'Mark emails as read?            :'
	#Enables mass failure protection. Failure notices will not be sent if there are more than 10 failures
	$mass_failure_check = Get-UserInput -Question 'Enable mass failure protection? :'
	#Decide whether to save to sharepoint
	$save_to_sharepoint = Get-UserInput -Question 'Save this file to sharepoint?   :'
	#Decide whether to send failure notices to staff
	$send_reports = Get-UserInput -Question 'Inform staff of failed backups? :'
	#Decide whether to display Excel windows
	$show_excel = Get-UserInput -Question 'Display Excel Windows?          :'
	#Prints status updates to the console
	$talk_to_me = Get-UserInput -Question 'Run script in verbose?          :'
	#Prevent the script from updating the database. Useful if you want to setup the DB a specific way to test behavior
	$update_db = Get-UserInput -Question 'Update the server DB?           :'
}
else{
	$mark_read = 1
	$mass_failure_check = 1
	$save_to_sharepoint = 1
	$send_reports = 1
	$show_excel = 0
	$talk_to_me = 0
	$update_db = 1
}

$start_time = Get-Date

if(($talk_to_me -eq 1) -or ($verbose -eq $true)){
	$VerbosePreference = "Continue"
}
else{
	#Not writing any messages to console, hide window and run in the background
	[Console.Window]::ShowWindow($consolePtr, 0)
}
Write-Verbose "CREATING THE COMMVAULT BACKUP REPORT"
#------------------------------------------------------------------------------------------------------------------------------
#Get the list of servers from (ServerList.xlsx)
#------------------------------------------------------------------------------------------------------------------------------
Write-Verbose "GETTING THE LIST OF SERVERS"
$server_list = Get-ServerList -file_path $file_path -Verbose:$false

#------------------------------------------------------------------------------------------------------------------------------
#Preparing the Excel ComObject
#------------------------------------------------------------------------------------------------------------------------------
$excel = New-Object -comobject Excel.Application -Verbose:$false
if($show_excel -eq 1){
	$excel.Visible = $True
}

#Get the server DB
$server_db_name = "ServerList.xlsx"
$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
$server_stats = $server_db.worksheets.item(1)
#------------------------------------------------------------------------------------------------------------------------------
# Get the CommVault reports from Outlook and save a temp to the desktop
#------------------------------------------------------------------------------------------------------------------------------
Write-Verbose "GETTING REPORTS FROM OUTLOOK"
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
$AllFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
$outlook = new-object -comobject outlook.application -Verbose:$False
$namespace = $outlook.GetNameSpace("MAPI")
$mailbox = $namespace.getDefaultFolder($AllFolders::olFolderInBox)
$messages = @()
$cv_reports = @()
#Put all messages from all folders into a list to check for CommVault reports
$messages += $namespace.getDefaultFolder(6).Items
foreach($folder in $mailbox.Folders){
	$messages += $folder.Items
}
foreach($item in $messages){
	if(($item.SenderName -match 'SDC_CommVault') -and ($item.UnRead -eq $True)){
		$cv_reports += $item
	}
} 
#sort messages so the oldest reports get processed first
$cv_reports = $cv_reports | Sort-Object -Property SentOn

if($cv_reports.Length -eq 0){
	Write-Verbose "NO NEW REPORTS WERE FOUND"
	Exit
}

foreach($message in $cv_reports) {
	$file = $message.Attachments.Item(1).filename
	$split_name = $file.ToString().split("_")
	if($split_name[0] -eq "BackupJobSummaryReport"){
		$savename = "CommVaultBackupReport-"
		#Get the date out of the attachment name, format it, and append it to the report name
		$savename = $savename + $split_name[4] + $split_name[2] + $split_name[3]
		$message.Attachments.Item(1).saveasfile((Join-Path $file_path "$savename.xls"))
#------------------------------------------------------------------------------------------------------------------------------
#Open the CommVault report in Excel, setup report workbook
#------------------------------------------------------------------------------------------------------------------------------
		Write-Verbose "OPENING THE FILE IN EXCEL"
		$excel.DisplayAlerts = $False
		$xlFixedformat = [Microsoft.Office.Interop.Excel.XlFileformat]::xlWorkbookDefault
		$workbook = $excel.Workbooks.Open("$file_path\$savename.xls")

		#Add worksheets for matching, failed, and missing servers and name them
		$workbook.Sheets.Add() | out-null
		$workbook.Sheets.Add() | out-null
		$workbook.Sheets.Add() | out-null
		$missing = $workbook.worksheets.item(1)
		$match = $workbook.worksheets.item(2)
		$failed = $workbook.worksheets.item(3)
		$original = $workbook.worksheets.item(4)
		$missing.name = 'MissingServers'
		$match.name = 'Match'
		$failed.name = 'Failed'
		
		#Find the report header.
		$location = $original.Cells.Find("Host Name")
		$header_begin = $location.Row
		$header_end = $header_begin +1

		#Find the first virtual machine header and status column
		$status_column = "D"
		$location = $original.Cells.Find("Machine Name")
		$virtual_row = $location.Row
		$status_column = $original.Range("A$($virtual_row):A$($virtual_row)").Find("Status").Column

		#The column values are not static. This finds which columns have the information the script needs
		$header = "A$($header_begin):A$($header_begin)"
		$active = $original.Range($header).Find("Running").Column
		$completed = $original.Range($header).Find("Completed").Column
		$cwe = $original.Range($header).Find("Completed with errors").Column
		$cww = $original.Range($header).Find("Completed with warnings").Column
		$delayed = $original.Range($header).Find("Delayed").Column
		$killed = $original.Range($header).Find("Killed").Column
		$server_column = $original.Range($header).Find("Client").Column
		$unsuccessful = $original.Range($header).Find("Unsuccessful").Column
		
		#ServerList Columns
		$server_ref = "A"
		$status_ref = "B"
		$failed_ref = "C"
		$backup_ref = "D"
		$ignore_ref = "E"
		$critical_ref = "F"
		$appdb_ref = "G"
		$staff_ref = "H"
		$app_ref = "I"

		#Copy the header and paste it into the first row of match and failed sheets add missing sheet header
		$header = "A$($header_begin):A$($header_end)"
		$header_range = $original.Range($header).EntireRow
		$header_range.Copy() | out-null
		$match.Paste()
		$failed.Paste()
		$missing.Cells(1,1) = "List of Servers Not Found In Backup List"

		#Create the Critical column for match and failed sheets
		$critical_column = "S"
		$match.Cells(1,$critical_column) = "Critical"
		$failed.Cells(1,$critical_column) = "Critical"
		$match.Cells.Item(1,$critical_column).Interior.ColorIndex = 15
		$match.Cells.Item(2,$critical_column).Interior.ColorIndex = 15
		$failed.Cells.Item(1,$critical_column).Interior.ColorIndex = 15
		$failed.Cells.Item(2,$critical_column).Interior.ColorIndex = 15
#------------------------------------------------------------------------------------------------------------------------------
#Copy our servers onto the match worksheet
#------------------------------------------------------------------------------------------------------------------------------
		Write-Verbose "FINDING OUR SERVERS AND COPYING THEM"
		#Saves any servers which were not found
		$missing_servers = @()
		
		#Hash table of backup status values
		$status_list = [ordered]@{$completed.ToString() = "Completed"; $active.ToString() = "Active"; $delayed.ToString() = "Delayed"; 
			$cww.ToString() = "CWW"; $cwe.ToString() = "CWE"; $unsuccessful.ToString() = "Failed"; $killed.ToString() = "Killed"}
		
		#Hash table to convert virtual machine status values to standardized status values
		$virtual_stats = @{"Failed" = "Failed"; "Killed" = "Killed"; "Completed with errors" = "CWE"; "Completed" = "Completed"; 
			"In Progress" = "Active"; "Waiting" = "Delayed"; "Completed with warnings" = "CWW" }
		
		#Hash table of Excel background color values
		$color_values = @{Active = 42; Completed = 35; CWE = 40; CWW = 8; Delayed = 39; Failed = 18; Killed = 38; Unknown = 18}
		
		foreach($server in $server_list){
			$server_location = $original.Cells.Find($server.Name)
			if($server_location){
				$row = $server_location.Row
				$critical = $server.Critical
				if($row -lt $virtual_row){
					foreach($status in $status_list.keys){
						if($original.Cells.Item($row, [int]$status).Value() -gt 0){
							$backup_status = $status_list[$status]
							Break
						}
						else{
							$backup_status = "Unknown"
						}
					}
				}
				else{
					$status = $original.Cells.Item($row,$status_column).Value()
					$backup_status = $virtual_stats[$status]
				}
				#found a match. Copy the row and paste to match worksheet
				$color = $color_values[$backup_status]
				$original_range = $original.Range("A$($row):A$($row)").EntireRow
				$original_range.Copy() | out-null
				$match_row = ($match.UsedRange.Rows).Count + 1
				$match_range = $match.Range("A$($match_row):A$($match_row)")
				$match.Paste($match_range)
				$match.Cells($match_row, 1).EntireRow.Interior.ColorIndex = $color
				if($critical -eq 1){
					$color_value = 6
					$crit_status = "Y"
				}
				else{
					$color_value = 2
					$crit_status = "N"
				}
				$match.Cells.Item($match_row, $critical_column).Value() = $crit_status
				$match.Cells.Item($match_row, $critical_column).Interior.ColorIndex = $color_value
#------------------------------------------------------------------------------------------------------------------------------
#Update the server status database (ServerList.xlsx)
#------------------------------------------------------------------------------------------------------------------------------
				if($update_db -eq 1){
					$stats_location = $server_stats.Cells.Find($server.Name)
					$stats_row = $stats_location.Row
					$server_stats.Cells.Item($stats_row, $status_ref).Value() = $backup_status
					if($backup_status -eq "Completed"){
						$server_stats.Cells.Item($stats_row, $failed_ref).Value() = 0
						$server_stats.Cells.Item($stats_row, $backup_ref).Value() = 0
						$server_stats.Cells.Item($stats_row, $ignore_ref).Value() = 0
					}
					#A value of 1 in column C means it needs to be reported. if the server is not critical and
					#just had errors it increments the number of backup attempts by 1.
					elseif(($backup_status -eq "Failed") -or ($backup_status -eq "Killed")){
						$server_stats.Cells.Item($stats_row, $failed_ref).Value() = 1
						$server_stats.Cells.Item($stats_row, $backup_ref).Value() = $server_stats.Cells.Item($stats_row, $backup_ref).Value() + 1
					}
					elseif($critical -eq 1){
						$server_stats.Cells.Item($stats_row, $failed_ref).Value() = 1
					}
					else{
						$server_stats.Cells.Item($stats_row, $backup_ref).Value() = $server_stats.Cells.Item($stats_row, $backup_ref).Value() + 1
						if($server_stats.Cells.Item($stats_row, $backup_ref).Value() -ge 3){
							$server_stats.Cells.Item($stats_row, $failed_ref).Value() = 1
						}
					}
				}
#------------------------------------------------------------------------------------------------------------------------------
#Copy failed server backups to the failed worksheet
#------------------------------------------------------------------------------------------------------------------------------
				if($backup_status -ne "Completed"){
					$failed_row = ($failed.UsedRange.Rows).count + 1
					$failed_range = $failed.Range("A$($failed_row):A$($failed_row)")
					$failed.Paste($failed_range)
					$failed.Cells($failed_row, 1).EntireRow.Interior.ColorIndex = $color
					$failed.Cells.Item($failed_row,$critical_column).Value() = $crit_status
					$failed.Cells.Item($failed_row,$critical_column).Interior.ColorIndex = $color_value
				}
			}
			else{
				$missing_servers += $server.Name
			}
		}	
#------------------------------------------------------------------------------------------------------------------------------
#Move critical column from S to A
#------------------------------------------------------------------------------------------------------------------------------
		Move-Columns -worksheet $match -from "S" -to "A"
		Move-Columns -worksheet $failed -from "S" -to "A"
#------------------------------------------------------------------------------------------------------------------------------
#Check for any missing servers and save them to missing worksheet
#------------------------------------------------------------------------------------------------------------------------------
		Write-Verbose "CHECKING if ANY SERVERS ARE MISSING"
		if($missing_servers.Count -gt 0){
			foreach($server in $missing_servers){
				$missing_row = ($missing.UsedRange.Rows).count + 1
				$missing.Cells.Item($missing_row,"A").Value() = $server
				if($update_db -eq 1){
					$missing_location = $server_stats.Cells.Find($server)
					if($missing_location){
						$server_stats.Cells.Item($missing_location.Row, $status_ref).Value() = "Missing"
					}
				}
			}
		}
#------------------------------------------------------------------------------------------------------------------------------
#Save the file to the Sharepoint site, open it for review, delete the temp file
#------------------------------------------------------------------------------------------------------------------------------
		if($save_to_sharepoint -eq 1){
			Write-Verbose "SAVING THE FILE TO SHAREPOINT"
			$tempname = $savename
			$savename = $savename + ".xlsx"
			$workbook.SaveAs((Join-Path $file_path $savename), $xlFixedformat)
			$workbook.Close()
			$file_to_save = get-childitem(Join-Path $file_path $savename)
			$sharepoint_url = "url"
			$webclient = New-Object System.Net.WebClient -Verbose:$false
			$webclient.UseDefaultCredentials = $True
			$webclient.UploadFile($sharepoint_url + "/" + $file_to_save.Name, "PUT", $file_to_save.FullName)
			$webclient.Dispose()
			if($show_excel -eq 1){
				$excel.Visible = $True
				$workbook = $excel.Workbooks.Open("$sharepoint_url/$savename")
			}
			Write-Verbose "CLEANING UP TEMP FILES"
			Remove-Item "$file_path\$savename"
			Remove-Item "$file_path\$tempname.xls"
		}
		else{
			$tempname = $savename
			$savename = $savename + ".xlsx"
			$workbook.SaveAs((Join-Path $file_path $savename), $xlFixedformat)
			$workbook.Close()
			Remove-Item $file_path"\"$tempname.xls
		}
		$server_db.Save()
		if($mark_read -eq 1){
			$message.UnRead = $False
		}
#------------------------------------------------------------------------------------------------------------------------------
#Report generated, notify report creator via email
#------------------------------------------------------------------------------------------------------------------------------
		Write-Verbose "GENERATION OF REPORT $savename COMPLETE"
		$mail = $outlook.CreateItem(0)
		$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
		$anthony = "email"
		$boris = "email"
		$john = "email"
		$mail.To = $me
		if($send_reports -eq 1){
			$mail.Cc = "$anthony; $boris; $john"
		}
		$mail.Subject = "CommVault Report Generated"
		$mail.Body = "$savename  was generated successfully."
		$mail.Send()
	}
}
#------------------------------------------------------------------------------------------------------------------------------
#Check if any servers need to be reported
#------------------------------------------------------------------------------------------------------------------------------
Write-Verbose "CHECKING if ANY SERVERS NEED TO BE REPORTED"
#Check the server status database and create list of servers with failed backups
$failed_servers = @{}
$server_count = ($server_stats.UsedRange.Rows).count
for($line=2; $line -le $server_count; $line++){
	$server_name = $server_stats.Cells.Item($line,$server_ref).Value()
	$server_status = $server_stats.Cells.Item($line, $status_ref).Value()
	$failed_backup = $server_stats.Cells.Item($line, $failed_ref).Value()
	$ignore = $server_stats.Cells.Item($line, $ignore_ref).Value()
	if($ignore -eq 0){
		if($server_status -eq "Missing"){
			$failed_servers.add($server_name,$server_status)
		}
		elseif($failed_backup -eq 1){
			$failed_servers.add($server_name,$server_status)
		}
	}
}
$server_db.Close()
$excel.Quit()
#------------------------------------------------------------------------------------------------------------------------------
#Lookup who needs to know about failures
#------------------------------------------------------------------------------------------------------------------------------
if($failed_servers.Count -gt 0){
	if($mass_failure_check -eq 1){
		if($failed_servers.Count -gt 10){
			#More than 10 servers had bad backups. Don't send notices. Manually verify that the backups failed.
			$send_reports = 0
			$mail = $outlook.CreateItem(0)
			$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
			$mail.To = $me
			$mail.Subject = "Mass Server Backup Failure"
			$mail.Body = "More than 10 servers had backup failures. Verify that the report was generated correctly"
			$mail.Send()
		}
	}
	Write-Verbose "LOOKING UP WHO NEEDS TO KNOW ABOUT THE FAILURES"
	#Get the stats of each failed server
	foreach($bad_server in $failed_servers.Keys){
		$server = $server_list| Where-Object {$_.Name -eq $bad_server}
		$is_critical = $server.Critical
		$backup_status = $failed_servers[$bad_server]
		$notify_drm = $server.AppDB
		$people_to_notify = $server.Staff
		$applications = $server.Applications
		if($is_critical -eq 1){
			$type = "CRITICAL"
		}
		else{
			$type = "NON-CRITICAL"
		}
#------------------------------------------------------------------------------------------------------------------------------
#Create the email notifications and send them out
#------------------------------------------------------------------------------------------------------------------------------
		foreach($person in $people_to_notify){
			$person = $person + "suffix"
			#This gets the current user's email address. I found the code here:
			#https://stackoverflow.com/questions/8666627/how-to-obtain-email-of-the-logged-in-user-in-powershell
			$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
			$team_lead = "email"
			$drm = "email"
			Write-Verbose "if notifications are enabled they will be sent to $person"
			if($send_reports -eq 1){
				Write-Verbose "REPORTING FAILURES TO THE STAFF RESPONSIBLE"
				$mail = $outlook.CreateItem(0)
				$mail.To = $person
				$mail.Cc = $team_lead
				if($notify_drm -eq 1){
					$mail.Cc = "$team_lead; $drm"
				}
				if($backup_status -eq "Missing"){
					$mail.Subject = "$type Server $bad_server Was Missed"
					$mail.Body = "The $type server $bad_server was not included in the CommVault backup. Please investigate why the 
server was missed. $bad_server is used for $applications
This message was generated by a script. if you believe you have received this message in error contact $me."
					$mail.Send()
				}
				elseif(($backup_status -eq "Active") -and ($is_critical -eq 1)){
					$mail.Subject = "CRITICAL Server $bad_server Backup is Active"
					$mail.Body = "The backup for CRITICAL server $bad_server is still Active. Please monitor this backup for completion.
$bad_server is used for $applications
This message was generated by a script. if you believe you have received this message in error contact $me."
					$mail.Send()
				}
				else{
					$mail.Subject = "$type Server $bad_server Failed Backup"
					$mail.Body = "The backup for $type server $bad_server had status $backup_status. It has failed or had 
errors for 3 days or more. Please investigate. $bad_server is used for $applications
This message was generated by a script. if you believe you have received this message in error contact $me."
					$mail.Send()
				}
			}
		}
		if($send_reports -eq 1){
			Write-Verbose "REPORTING TO SELF"
			$mail = $outlook.CreateItem(0)
			$mail.To = $me
			$mail.Subject = "Server Backup Failure Notification Report"
			$mail.Body = "The following people: $people_to_notify  were informed about the backup failure of: $bad_server"
			$mail.Send()
		}
	}
}
$end_time = Get-Date
$seconds = ($end_time - $start_time).TotalSeconds
Write-Host "Total time taken to run the report was $seconds seconds"
