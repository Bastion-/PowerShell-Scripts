<#
-----------------------------------------------------------------------
Name: What's Up Gold Server Monitoring Script
Author:	Anthony Dunaway
Date: 05/21/18
Updated: 07/31/18
Description:
Records any of our servers that are reported by What's Up Gold and 
logs them in a log file called Gold_Report.
-----------------------------------------------------------------------
#>
#----------------------------------------------------------------------------------------------------------------------
#Script Imports
#----------------------------------------------------------------------------------------------------------------------
."C:\Development\Helper_Scripts\Get_Server_Objects.ps1"
."C:\Development\Helper_Scripts\Hide_Window.ps1"

#----------------------------------------------------------------------------------------------------------------------
#Get the list of servers and the list of staff and servers
#----------------------------------------------------------------------------------------------------------------------
$server_list = Get-ServerList -file_path "C:\Development\CommVaultScriptDev"

#----------------------------------------------------------------------------------------------------------------------
#Setup some variables
#----------------------------------------------------------------------------------------------------------------------
#Hide window
[Console.Window]::ShowWindow($consolePtr, 0)
#Script Path
$file_path = $PSScriptRoot.ToString()
#temp file
$gold_digging = "gold_digging.txt"
#What's Up Gold report file
$gold_report = "Gold_Report"
#Flag for if a report has been sent today
$sent_today = 0
#number of our servers found during a 24 hour period
$nuggets = 0

#----------------------------------------------------------------------------------------------------------------------
#Open Outlook 
#----------------------------------------------------------------------------------------------------------------------
$outlook = new-object -comobject "Outlook.Application"
$mapi = $outlook.getnamespace("mapi")
$desktop_path = [Environment]::GetFolderPath("Desktop")
$inbox = $mapi.GetDefaultFolder(6)

Start-Log -path $file_path -name $gold_report
#----------------------------------------------------------------------------------------------------------------------
#Loop through unread What's Up Gold emails
#----------------------------------------------------------------------------------------------------------------------
$gold_reports = $inbox.Folders | where-object {$_.name -eq "What's Up Gold"}
ForEach($report in $gold_reports.Items){
	$subject = $report.Subject
	$subject = $subject.Substring(3, $subject.Length - 3)
	$sent = $report.ReceivedTime
	$name = ""
	$details = ""
	if($report.UnRead -eq $True){
		$content = $report.body
		#Add email content to text file so it can be processed line by line
		Add-Content "$file_path\$gold_digging" $content
		(gc "$file_path\$gold_digging") | ? {$_.trim() -ne "" } | set-content "$file_path\$gold_digging"
		$content_reader = [System.IO.File]::OpenText("$file_path\$gold_digging")
		$report_type = $subject.Split(":")
		$report_type = $report_type.Trim()
		$report_type = $report_type[0].ToString()
		for(){
			$content_lines = $content_reader.ReadLine()
			if ($content_lines -eq $null) { 
				break 
			}
#----------------------------------------------------------------------------------------------------------------------
#Check what type of WhatsUp Gold report this is
#----------------------------------------------------------------------------------------------------------------------
			if($report_type -eq "WhatsUp Gold Alert Center"){
				#report is a disk utilization report
				$content_line = $content_lines.Split(" ")
				$name = $content_line[0]
				$name = $name.Trim()
				foreach($server in $server_list){
					if($server.Name -eq $name){
						$info = "Disk Utilization Report: Date $sent, Server: $content_lines "
						Add-LogEntry -path $file_path -name $gold_report -warning -content $info
						$nuggets++
					}
				}
				Continue
			}
			else{
				#report is not a disk utilization report
				$content_line = $content_lines.Split(" ")
				if($content_line[0] -eq "Name:"){
					$name = $content_line[1].Split(".")
					$name = $name[0]
				}
				if($content_line[0] -eq "-"){
					$details += $content_lines
				}
			}
			if(-Not [string]::IsNullOrEmpty($name)){
				foreach($server in $server_list){
					if($server.Name -eq $name){
						$info = "$subject, " + "Date $sent, " + " Details: $details"
						Add-LogEntry -path $file_path -name $gold_report -warning -content $info
						$nuggets++
					}
				}
			}
		}
		$content_reader.Close()
		Remove-Item "$file_path\$gold_digging"
		$report.UnRead = $False
	}
}
#----------------------------------------------------------------------------------------------------------------------
# Check if it is time to email out the report and if so send it out
#----------------------------------------------------------------------------------------------------------------------
$time_stamp = Get-Date -Format g 
if($nuggets -eq 0){
	$report_footer = "None of our servers were reported by What's Up Gold in the last 24 hours. $time_stamp"
	Add-LogEntry -path $file_path -name $gold_report -info -content $report_footer
}
elseif($nuggets -eq 1){
	$report_footer = "Our servers were reported once by What's Up Gold in the last 24 hours. $time_stamp"
	Add-LogEntry -path $file_path -name $gold_report -info -content $report_footer
}
else{
	$report_footer = "Our servers were reported $nuggets times by What's Up Gold in the last 24 hours. $time_stamp"
	Add-LogEntry -path $file_path -name $gold_report -info -content $report_footer
}
Stop-Log -path $file_path -name $gold_report
$mail = $outlook.CreateItem(0)
$me = ([adsi]"LDAP://$(whoami /fqdn)").mail.ToString()
$team_lead = "kenneth.b.hill@dhsoha.state.or.us"
$mail.To = $me
$mail.Cc = $team_lead
$mail.Subject = "What's Up Gold Daily Report"
$mail.Body = "Attached is the report"
$mail.Attachments.Add("$file_path\$gold_report"+".txt") | out-null
$mail.Send()
$nuggets = 0


