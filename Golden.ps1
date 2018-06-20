<#
-----------------------------------------------------------------------
Name: What's Up Gold Server Monitoring Script
Author:	Anthony Dunaway
Date: 05/21/18
Updated: 06/20/18
Description:
This is a testing script for a future What's Up Gold monitoring script.
Currently it monitors the What's Up Gold Outlook folder for any unread
emails. It checks if the email regards one of our servers. If it does it 
logs the information about the server to a master document. Once a day
it emails the report from the previous 24 hours to the user who is running 
the script. This report includes all of the relevant server information
as well as the total number of servers reported that we care about. 
-----------------------------------------------------------------------
#>
#----------------------------------------------------------------------------------------------------------------------
#Script Imports
#----------------------------------------------------------------------------------------------------------------------
."I:\ISE\CommVault\CommVaultScript\Helper_Scripts\Get_Servers.ps1"
."I:\ISE\CommVault\CommVaultScript\Helper_Scripts\Get_Staff_Servers.ps1"

#----------------------------------------------------------------------------------------------------------------------
#Hide Window
#----------------------------------------------------------------------------------------------------------------------
Add-Type -Name Window -Namespace Console -MemberDefinition '
	[DllImport("Kernel32.dll")]
	public static extern IntPtr GetConsoleWindow();

	[DllImport("user32.dll")]
	public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
	'
# $consolePtr = [Console.Window]::GetConsoleWindow()
# [Console.Window]::ShowWindow($consolePtr, 0)

#----------------------------------------------------------------------------------------------------------------------
#Get the list of servers and the list of staff and servers
#----------------------------------------------------------------------------------------------------------------------
$server_list = Get-ServerInformation -Verbose:$false
#$staff_servers = Get-StaffServers -Verbose:$false

#----------------------------------------------------------------------------------------------------------------------
#Setup some variables
#----------------------------------------------------------------------------------------------------------------------
$file_path = Get-Location
#temp file
$gold_digging = "gold_digging.txt"
#What's Up Gold report file
$gold_report = "Gold_Report.txt"
#Flag for if a report has been sent today
$sent_today = 0
#number of our servers found during a 24 hour period
$nuggets = 0
#The hour from a 24 hour clock that Gold Reports get sent out. 
$send_report_time = 8
#Blank space for report spacing
$space = " "

#----------------------------------------------------------------------------------------------------------------------
#Open Outlook 
#----------------------------------------------------------------------------------------------------------------------
$outlook = new-object -comobject "Outlook.Application"
$mapi = $outlook.getnamespace("mapi")
$desktop_path = [Environment]::GetFolderPath("Desktop")
$inbox = $mapi.GetDefaultFolder(6)

While(1){
	#Number of seconds between each check
	sleep -s 1
#----------------------------------------------------------------------------------------------------------------------
#Loop through unread What's Up Gold emails
#----------------------------------------------------------------------------------------------------------------------
	$gold_reports = $inbox.Folders | where-object {$_.name -eq "What's Up Gold"}
	ForEach($report in $gold_reports.Items){
		$subject = $report.Subject
		$sent = $report.ReceivedTime
		$type = ""
		$name = ""
		$details = ""
		$server_found = 0
		if($report.UnRead -eq $True){
			$content = $report.body
			#Add email content to text file so it can be processed line by line
			Add-Content "$file_path\$gold_digging" $content
			(gc "$file_path\$gold_digging") | ? {$_.trim() -ne "" } | set-content "$file_path\$gold_digging"
			$content_reader = [System.IO.File]::OpenText("$file_path\$gold_digging")
			$report_type = $subject.Split(":")
			$report_type = $report_type.Trim()
			if($report_type[1] -eq "WhatsUp Gold Alert Center"){
				Add-Content "$file_path\$gold_report" $space
				Add-Content "$file_path\$gold_report" "Disk Utilization Report"
				Add-Content "$file_path\$gold_report" $sent
				Add-Content "$file_path\$gold_report" $subject
			}
			for(){
				$content_lines = $content_reader.ReadLine()
				if ($content_lines -eq $null) { 
					break 
				}
#----------------------------------------------------------------------------------------------------------------------
#Check what type of WhatsUp Gold report this is
#----------------------------------------------------------------------------------------------------------------------
				if($report_type[1] -eq "WhatsUp Gold Alert Center"){
					#report is a disk utilization report
					$content_line = $content_lines.Split(" ")
					$name = $content_line[0]
					$name = $name.Trim()
					foreach($server in $server_list){
						if($server -eq $name){
							$server_found = 1
							Add-Content "$file_path\$gold_report" $content_lines
							$nuggets++
						}
					}
					Continue
				}
				else{
					#report is not a disk utilization report
					$content_line = $content_lines.Split(" ")
					if($content_line[0] -eq "Type:"){
						$type = $content_lines
					}
					if($content_line[0] -eq "Name:"){
						$name = $content_line[1].Split(".")
						$name = $name[0]
					}
					if($content_line[0] -eq "-"){
						$details += $content_lines
						$details += $space
					}
				}
				if(-Not [string]::IsNullOrEmpty($name)){
					foreach($server in $server_list){
						if($server -eq $name){
							Add-Content "$file_path\$gold_report" $sent
							Add-Content "$file_path\$gold_report" $subject
							Add-Content "$file_path\$gold_report" $type
							Add-Content "$file_path\$gold_report" $name
							Add-Content "$file_path\$gold_report" $details
							Add-Content "$file_path\$gold_report" $space
							$nuggets++
						}
					}
				}
			}
			if(($report_type[1] -eq "WhatsUp Gold Alert Center") -and ($server_found -eq 0)){
					Add-Content "$file_path\$gold_report" "None of our servers were reported"
				}
			$content_reader.Close()
			Remove-Item "$file_path\$gold_digging"
			$report.UnRead = $False
		}
	}
#----------------------------------------------------------------------------------------------------------------------
# Check if it is time to email out the report and if so send it out
#----------------------------------------------------------------------------------------------------------------------
	$report_time = (Get-Date -Uformat %H).ToString()
	$report_time = [int]$report_time
	if(($send_report_time - $report_time -eq 0) -and ($sent_today -eq 0)){
		$time_stamp = Get-Date -Format g 
		$sent_today = 1
		if($nuggets -eq 0){
			$report_footer = "None of our servers were reported by What's Up Gold in the last 24 hours. $time_stamp"
			Add-Content "$file_path\$gold_report" $report_footer
		}
		elseif($nuggets -eq 1){
			$report_footer = "Our servers were reported once by What's Up Gold in the last 24 hours. $time_stamp"
			Add-Content "$file_path\$gold_report" $report_footer
		}
		else{
			$report_footer = "Our servers reported $nuggets times by What's Up Gold in the last 24 hours. $time_stamp"
			Add-Content "$file_path\$gold_report" $report_footer
		}
		$mail = $outlook.CreateItem(0)
		$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
		$team_lead = "email"
		$mail.To = $me
		$mail.Cc = $team_lead
		$mail.Subject = "What's Up Gold Daily Report"
		$mail.Body = "Attached is the report"
		$mail.Attachments.Add("$file_path\$gold_report") | out-null
		$mail.Send()
		$nuggets = 0
	}
	elseif($send_report_time - $report_time -ne 0){
		$sent_today = 0
	}
	else{
		continue
	}
}

