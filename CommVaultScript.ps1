<#
-----------------------------------------------------------------------
Name: CommVault Backup Report Script
Author: Anthony Dunaway
Date: 01/16/18
Updated: 06/27/18
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
."C:\Development\CommVaultScript\Helper_Scripts\Get_Servers.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\Hide_Window.ps1"
."C:\Development\CommVaultScript\Helper_Scripts\User_Input.ps1"

#------------------------------------------------------------------------------------------------------------------------------
# Debug menu options
#------------------------------------------------------------------------------------------------------------------------------
if($debug -eq $true){
	Write-Verbose "RUNNING IN DEBUG MODE"
	#Allows you to run the script as much as you like on the same email without the need to manually mark the message as unread
	$mark_read = Get-UserInput -Question 'Mark emails as read?            :'
	#Enables mass failure protection. Notices will not be sent if there are more than 10 failures
	$mass_failure_check = Get-UserInput -Question 'Enable mass failure protection? :'
	#When testing you might not want to save files to sharepoint over and over again
	$save_to_sharepoint = Get-UserInput -Question 'Save this file to sharepoint?   :'
	#Stops script from emailing staff. Useful to prevent false alerts when testing
	$send_reports = Get-UserInput -Question 'Inform staff of failed backups? :'
	#Display Excel windows
	$show_excel = Get-UserInput -Question 'Display Excel Windows?          :'
	#Prints status updates to the console
	$talk_to_me = Get-UserInput -Question 'Run script in verbose?          :'
	#Prevent the script from updating the database. Useful if you want to setup the DB a specific way to test behavior
	$update_db = Get-UserInput -Question 'Update the server DB?           :'

}
else{
	$mark_read = 1
	$mass_failure_check = 1
	$save_to_sharepoint = 0
	$send_reports = 0
	$show_excel = 0
	$talk_to_me = 1
	$update_db = 1
}

if(($talk_to_me -eq 1) -or ($verbose -eq $true)){
	$VerbosePreference = "Continue"
}
else{
	#Not writing any messages to console, hide window and run in the background
	[Console.Window]::ShowWindow($consolePtr, 0)
}

Write-Verbose "CREATING THE COMMVAULT BACKUP REPORT"
$file_path = "C:\Development\CommVaultScript\"
$server_db_name = "ServerList.xlsx"

#------------------------------------------------------------------------------------------------------------------------------
#Get the list of servers from (ServerList.xlsx)
#------------------------------------------------------------------------------------------------------------------------------
Write-Verbose "GETTING THE LIST OF SERVERS"
$server_list = Get-ServerList -critical -Verbose:$false

#--------------------------------------------------------------------
# Get the CommVault reports from Outlook and save a temp to the desktop
#--------------------------------------------------------------------
Write-Verbose "GETTING REPORTS FROM OUTLOOK"
$outlook = new-object -comobject Outlook.Application -Verbose:$false
$mapi = $outlook.getnamespace("mapi")
#Check the Inbox. Change 6 to another value if you want to check a different folder.
$inbox = $mapi.GetDefaultFolder(6)
$commvault_reports = $inbox.Folders | where-object {$_.name -eq "CommVault Reports"}
$cv_reports = @()
#I move my CommVault reports to another folder
if($commvault_reports){
	foreach($item in $commvault_reports.Items){
		if(($item.SenderName -match 'sender') -and ($item.UnRead -eq $True)){
			$cv_reports += $item
		}
	}
}
#custom folder does not exist use inbox
else{
	foreach($item in $inbox.Items){
		if(($item.SenderName -match 'sender') -and ($item.UnRead -eq $True)){
			$cv_reports += $item
		}
	}
}
$cv_reports = $cv_reports | Sort-Object -Property SentOn
if($cv_reports.Length -eq 0){
	Write-Verbose "NO NEW REPORTS WERE FOUND"
	Exit
}

foreach($message in $cv_reports) {
	$savename = "CommVaultBackupReport-"
	$file = $message.Attachments.Item(1).filename
	$split_name = $file.ToString().split("_")
	#This gets the date out of the attachment name, formats it, and appends it to the report name
	$savename = $savename + $split_name[4] + $split_name[2] + $split_name[3]
	$message.Attachments.Item(1).saveasfile((Join-Path $file_path "$savename.xls"))

#------------------------------------------------------------------------------------------------------------------------------
#Open the CommVault report in Excel, setup report workbook
#------------------------------------------------------------------------------------------------------------------------------
	Write-Verbose "OPENING THE FILE IN EXCEL"
	$excel = New-Object -comobject Excel.Application -Verbose:$false
	if($show_excel -eq 1){
		$excel.Visible = $True
	}
	$excel.DisplayAlerts = $False
	$xlFixedformat = [Microsoft.Office.Interop.Excel.XlFileformat]::xlWorkbookDefault
	$workbook = $excel.Workbooks.Open("$file_path\$savename.xls")

	#Get the server DB
	$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
	$server_stats = $server_db.worksheets.item(1)

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
	$header_begin = 0
	$header_end = 1
	for($row=1; $row -le 100; $row++){
		$value = $original.Cells.Item($row,"A").Value()
		if($value -eq "client"){
			$header_begin = $row
			$header_end = $row + 1
			Break
		}
	}

	#Find the last row for the needed search range.
	$last_row = 0
	for($row=500; $row -le 800; $row++){
		$value = $original.Cells.Item($row,"A").Value()
		if($value -eq "client"){
			$last_row = $row
			Break
		}
	}

	#The column values are not static. This finds which columns have the information the script needs
	$active = " "
	$completed = " "
	$cwe = " "
	$cww = " "
	$delayed = " "
	$killed = " "
	$server_column = " "
	$unsuccessful = " "
	
	for($column = 1; $column -le 26; $column++ ){
		if($original.Cells.Item($header_begin, $column).Value() -eq "Client" ){
			$server_column = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Completed" ){
			$completed = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Completed with errors" ){
			$cwe = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Completed with warnings" ){
			$cww = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Killed" ){
			$killed = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Unsuccessful" ){
			$unsuccessful= $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Running" ){
			$active = $column
		}
		elseif($original.Cells.Item($header_begin, $column).Value() -eq "Delayed" ){
			$delayed = $column
		}
	}
	
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
	#Saves which of our servers were reported
	$reported_servers = @()

	#Current row counters
	$failed_counter = 2
	$missing_counter = 2
	$match_counter = 2
	
	#Hash table of backup status values
	$status_list = [ordered]@{$completed.ToString() = "Completed"; $active.ToString() = "Active"; $delayed.ToString() = "Delayed"; 
	$cww.ToString() = "CWW"; $cwe.ToString() = "CWE"; $unsuccessful.ToString() = "Failed"; $killed.ToString() = "Killed"}
	
	#Hash table of Excel background color values
	$color_values = @{Active = 42; Completed = 35; CWE = 40; CWW = 8; Delayed = 39; Failed = 18; Killed = 38; Unknown = 18}
	
	for($row=$header_end + 2; $row -le $last_row; $row++){
		$report_server = $original.Cells.Item($row, $server_column).Value()
		$critical = 0
		$found = 0
		#Check if current row is one of our servers
		foreach($server in $server_list.keys){
			if($report_server -eq $server){
				$reported_servers += $report_server
				$critical = $server_list[$server]
				$found = 1
				$backup_status = " "
				#Get the backup status
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
		}
		if($found -eq 1){
			#found a match. Copy the row and paste to match worksheet
			$color = $color_values[$backup_status]
			$reference = "A$($row):A$($row)"
			$original_range = $original.range($reference).EntireRow
			$original_range.Copy() | out-null
			$match_counter = $match_counter + 1
			$match_range = $match.Range("A$($match_counter):A$($match_counter)")
			$match.Paste($match_range)
			$match.Cells($match_counter, 1).EntireRow.Interior.ColorIndex = $color
			if($critical -eq 1){
				$color_value = 6
				$crit_status = "Y"
			}
			else{
				$color_value = 2
				$crit_status = "N"
			}
			$match.Cells.Item($match_counter, $critical_column).Value() = $crit_status
			$match.Cells.Item($match_counter, $critical_column).Interior.ColorIndex = $color_value
#------------------------------------------------------------------------------------------------------------------------------
#Update the server status database (ServerList.xlsx)
#------------------------------------------------------------------------------------------------------------------------------
			$server_count = ($server_stats.UsedRange.Rows).count
			if($update_db -eq 1){
				for($line=1; $line -le $server_count - 1; $line++){
					$current_server = $server_stats.Cells.Item($line,"A").Value()
					if(($report_server -eq $current_server) -and ($server_stats.Cells.Item($line, "E").Value() -ne 1)){
						$server_stats.Cells.Item($line, "B").Value() = $backup_status
						if($backup_status -eq "Completed"){
							$server_stats.Cells.Item($line, "C").Value() = 0
							$server_stats.Cells.Item($line,"D").Value() = 0
							$server_stats.Cells.Item($line,"E").Value() = 0
						}
						#A value of 1 in column C means it needs to be reported. if the server is not critical and
						#just had errors it increments the number of backup attempts by 1.
						elseif(($backup_status -eq "Failed") -or ($backup_status -eq "Killed")){
							$server_stats.Cells.Item($line, "C").Value() = 1
							$server_stats.Cells.Item($line, "D").Value() = $server_stats.Cells.Item($line, "D").Value() + 1
						}
						elseif($critical -eq 1){
							$server_stats.Cells.Item($line, "C").Value() = 1
							$server_stats.Cells.Item($line, "D").Value() = $server_stats.Cells.Item($line, "D").Value() + 1
						}
						else{
							$server_stats.Cells.Item($line, "D").Value() = $server_stats.Cells.Item($line, "D").Value() + 1
							if($server_stats.Cells.Item($line, "D").Value() -ge 3){
							}
						}
						Break
					}
				}
			}
#------------------------------------------------------------------------------------------------------------------------------
#Copy failed server backups to the failed worksheet
#------------------------------------------------------------------------------------------------------------------------------
			if($backup_status -ne "Completed"){
				$failed_counter = $failed_counter + 1
				$failed_row = $failed.UsedRange.rows.count + 1
				$failed_range = $failed.Range("A$($failed_counter):A$($failed_counter)")
				$failed.Paste($failed_range)
				$failed.Cells($failed_counter, 1).EntireRow.Interior.ColorIndex = $color
				$failed.Cells.Item($failed_counter,$critical_column).Value() = $crit_status
				$failed.Cells.Item($failed_counter,$critical_column).Interior.ColorIndex = $color_value
			}
		}
	}

#------------------------------------------------------------------------------------------------------------------------------
#Move critical columns to column A
#------------------------------------------------------------------------------------------------------------------------------
	$shift_right = -4161
	$critical_reference = "S1:S1"

	#Match worksheet
	[void]$match.cells.entireColumn.Autofit()
	[void]$match.cells.entireRow.Autofit()
	$match_critical = $match.range($critical_reference).EntireColumn
	$match_critical.Copy() | out-null
	$new_column = $match.Range("A1").EntireColumn
	$new_column.Insert($shift_right) | out-null
	[void]$match_critical.Delete()

	#Failed worksheet
	[void]$failed.cells.entireColumn.Autofit()
	[void]$failed.cells.entireRow.Autofit()
	$failed_critical = $failed.range($critical_reference).EntireColumn
	$failed_critical.Copy() | out-null
	$new_column = $failed.Range("A1").EntireColumn
	$new_column.Insert($shift_right) | out-null
	[void]$failed_critical.Delete()

#------------------------------------------------------------------------------------------------------------------------------
#Check for any missing servers and save them to missing worksheet
#------------------------------------------------------------------------------------------------------------------------------
	Write-Verbose "CHECKING if ANY SERVERS ARE MISSING"
	foreach($server in $server_list.keys){
		$not_missing = 0
		foreach($report in $reported_servers){
			if($server -eq $report){
				$not_missing = 1
				Break
			}
		}
		if($not_missing -eq 0){
			$missing.Cells.Item($missing_counter,"A").Value() = $server
			$missing_counter = $missing_counter + 1
			if($update_db -eq 1){
				for($line=2; $line -le $server_count; $line++){
					$server_name = $server_stats.Cells.Item($line,"A").Value()
					if($server -eq $server_name){
						$server_stats.Cells.Item($line,"B").Value() = "Missing"
						Break
					}
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
		#uncomment the two lines below this if you want the script to open the file from sharepoint for inspection
		#$excel.Visible = $True
		#$workbook = $excel.Workbooks.Open("$sharepoint_url/$savename")
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
	$server_db.Close()
	if($mark_read -eq 1){
		$message.UnRead = $False
	}

#------------------------------------------------------------------------------------------------------------------------------
#Report generated notify report creator via email
#------------------------------------------------------------------------------------------------------------------------------
	# Write-Verbose "GENERATION OF REPORT $savename COMPLETE"
	$mail = $outlook.CreateItem(0)
	$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
	$anthony = "email"
	$boris = "email"
	$john = "email"
	$mail.To = $me
	#$mail.Cc = "$anthony; $boris; $john"
	$mail.Subject = "CommVault Report Generated"
	$mail.Body = "$savename  was generated successfully."
	$mail.Send()
}

#------------------------------------------------------------------------------------------------------------------------------
#Check if any servers need to be reported
#------------------------------------------------------------------------------------------------------------------------------
Write-Verbose "CHECKING if ANY SERVERS NEED TO BE REPORTED"
#Check the server status database
$server_db = $excel.Workbooks.Open("$file_path\$server_db_name")
$server_stats = $server_db.worksheets.item(1)
#Create list of servers with failed backups
$failed_servers = @()
$server_count = ($server_stats.UsedRange.Rows).count
for($line=2; $line -le $server_count; $line++){
	$server_name = $server_stats.Cells.Item($line,"A").Value()
	$server_status = $server_stats.Cells.Item($line,"B").Value()
	$failed_backup = $server_stats.Cells.Item($line,"C").Value()
	$backup_attempts = $server_stats.Cells.Item($line,"D").Value()
	$ignore = $server_stats.Cells.Item($line,"E").Value()
	if(($ignore -eq 0) -and ($server_status -eq "Missing")){
		$failed_servers += $server_name
	}
	elseif(($ignore -eq 0) -and ($failed_backup -eq 1)){
		$failed_servers += $server_name
	}
	#If there have been errors for 3 or more days send a notice
	elseif(($ignore -eq 0) -and ($backup_attempts -ge 3)){
		$failed_servers += $server_name
	}
}

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
	foreach($bad_server in $failed_servers){
		for($line = 2; $line -le $server_count; $line++){
			$server_name = $server_stats.Cells.Item($line,"A").Value()
			if($bad_server -eq $server_name){
				$is_critical = $server_list[$bad_server]
				$backup_status = $server_stats.Cells.Item($line,"B").Value()
				$notify_drm = $server_stats.Cells.Item($line,"G").Value()
				$people_to_notify = $server_stats.Cells.Item($line,"H").Value().ToString().Split(",")
				$three_days = 0
				if(($server_stats.Cells.Item($line,"C").Value() -eq 0) -and ($server_stats.Cells.Item($line,"D").Value() -ge 3)){
					$three_days = 1
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
					$team_lead2 = "email"
					$drm = "email"
					Write-Verbose "if notifications are enabled they will be sent to $person"
					if($send_reports -eq 1){
						Write-Verbose "REPORTING FAILURES TO THE STAFF RESPONSIBLE"
						#Create the email notification and send it out.
						$mail = $outlook.CreateItem(0)
						$mail.To = $person
						$mail.Cc = $team_lead
						if($notify_drm -eq 1){
							$mail.Cc = "$team_lead; $drm"
						}
						if(($backup_status -eq "Missing")-and ($is_critical -eq 1)){
							$mail.Subject = "CRITICAL Server $bad_server Was Missed"
							$mail.Body = "The CRITICAL server $bad_server was not included in the CommVault backup and requires action.
This message was generated by a script. if you believe you have received this message in error contact $me."
							$mail.Send()
						}
						elseif(($backup_status -eq "Missing")-and ($is_critical -eq 0)){
							$mail.Subject = "NON-CRITICAL Server $bad_server Was Missed"
							$mail.Body = "The NON-CRITICAL server $bad_server was not included in the CommVault backup.
This message was generated by a script. if you believe you have received this message in error contact $me."
							$mail.Send()
						}
						elseif(($backup_status -eq "Active") -and ($is_critical -eq 1)){
							$mail.Subject = "CRITICAL Server $bad_server Backup Not Complete"
							$mail.Body = "The backup for CRITICAL server $bad_server is still Active. Please monitor this backup for completion.
This message was generated by a script. if you believe you have received this message in error contact $me."
							$mail.Send()
						}
						elseif($is_critical -eq 1){
							$mail.Subject = "CRITICAL Server $bad_server Failed Backup"
							$mail.Body = "The backup for CRITICAL server $bad_server was incomplete with status $backup_status and requires action.
This message was generated by a script. if you believe you have received this message in error contact $me."
							$mail.Send()
						}
						elseif($three_days -eq 0){
							$mail.Subject = "NON-CRITICAL Server $bad_server Failed Backup"
							$mail.Body = "The backup for NON-CRITICAL server $bad_server was incomplete with status $backup_status.
This message was generated by a script. if you believe you have received this message in error contact $me."
							$mail.Send()
						}
						else{
							$mail.Subject = "NON-CRITICAL Server $bad_server Failed Backup"
							$mail.Body = "The backup for NON-CRITICAL server $bad_server has had errors or reported as active or delayed for 3 days or more.
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
	}
}
$server_db.Save()
$server_db.Close()
$excel.Quit()
#------------------------------------------------------------------------------------------------------------------------------
#Cleaning up so there are not a bunch of Excel tasks left running.
#------------------------------------------------------------------------------------------------------------------------------
[GC]::Collect()
