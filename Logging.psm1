<#
-----------------------------------------------------------------------
Name: Logging Module
Author:	Anthony Dunaway
Date: 07/20/18
Updated: 07/26/18
Description:
Logging Module
-----------------------------------------------------------------------
#>

#Returns a PSCredential object. Requires a text file with an encrypted password. PRIVATE
function Get-MyCredentials(){
	[cmdletbinding()]
	param(
		[string]$filepath = "C:\Development\Password_Files",
		[string]$username = "OR0237713",
		[string]$filename = "password.txt"
	)
	
	$password = get-content "$filepath\$filename" | convertto-securestring
	$credentials = new-object System.Management.Automation.PSCredential($username, $password)
	Return $credentials
}

#Returns the name of the calling script. PRIVATE
function Get-CallingScriptName(){
	[cmdletbinding()]
	param()
	$calling_script_path = Get-Item (Get-PSCallStack)[(Get-PSCallStack).length-2].ScriptName
	$script = $calling_script_path.ToString().Split("\")
	$script = $script[$script.Length - 1]
	$script = $script.Split(".")
	$script = $script[0] + "_log.txt"
	Return $script	
}

#Create a new log file. PRIVATE
function New-LogFile(){
	[cmdletbinding()]
	param(
		[string] $path =  [Environment]::GetFolderPath("MyDocuments")
	)
	$script = Get-CallingScriptName
	if(!(Test-Path -Path "$path\$script")){
		New-Item -Path $path -Name $script | Out-Null 
	}
}

#Start the log with informational header
function Start-Log(){
<#
.SYNOPSIS
	Start a log file
.DESCRIPTION
	Check if log file exists. If not, create one and initialize it with the name of the calling script and date.
	The file name will be the name of the calling script _log.txt
.PARAMETER path
	Location where you would like the log file saved. Default is MyDocuments.
.EXAMPLE
	Start-Log -path "C:\Development\"
#>
	[cmdletbinding()]
	param(
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$path
	)
	$script = Get-CallingScriptName
	if($path){
		New-LogFile -path $path
	}
	else{
		New-LogFile 
		$path = [Environment]::GetFolderPath("MyDocuments")
	}
	$date = Get-Date -UFormat "%A %H:%M"
	$content = "----------------------------------------------------------------------------------------
Date: $date
Beginning of $script 
----------------------------------------------------------------------------------------

"
	Add-Content -Path "$path\$script" -Value $content
}

#Stop the log with informational footer
function Stop-Log(){
<#
.SYNOPSIS
	Adds a footer to the log file with the date and name of the calling script.
.PARAMETER path
	Location of the log file you would like to add a stop footer to. Defualt is MyDocuments.
.EXAMPLE
	Stop-Log -path "C:\Development\"
#>
	[cmdletbinding()]
	param(
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$path = [Environment]::GetFolderPath("MyDocuments")
	)	
	$script = Get-CallingScriptName	
	$date = Get-Date -UFormat "%A %H:%M"
	$content = "
	
----------------------------------------------------------------------------------------
Date: $date
End of $script
----------------------------------------------------------------------------------------
"	
	Add-Content -Path "$path\$script" -Value $content
}

function Add-LogEntry(){
<#
.SYNOPSIS
	Adds an entry to the log file.
.PARAMETER info
	Switch, if used prefixes INFORMATION: to the entry.
.PARAMETER error
	Switch, if used prefixes ERROR: to the entry.
.PARAMETER warning
	Switch, if used prefixes WARNING: to the entry.
.PARAMETER debugging
	Switch, if used prefixes DEBUG: to the entry.
.PARAMETER path
	Location of the log file to add an entry to
.PARAMETER content
	The text to be written to the log file.
.EXAMPLE
	Add-LogEntry -path "C:\Development\" -info -content "This will be written to the log file"
#>
	[cmdletbinding()]
	param(
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[switch]$info,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[switch]$error,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[switch]$warning,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[switch]$debugging,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$path = [Environment]::GetFolderPath("MyDocuments"),
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$content
	)	
	$script = Get-CallingScriptName	
	if($info){
		$prefix = "INFORMATION: "
	}
	elseif($error){
		$prefix = "ERROR: "
	}
	elseif($warning){
		$prefix = "WARNING: "
	}
	elseif($debugging){
		$prefix = "DEBUG: "
	}
	else{
		$prefix = ""
	}
	$content = $prefix + $content	
	Add-Content -Path "$path\$script" -Value $content
	Write-Debug "Content"
}

function Send-Log(){
<#
.SYNOPSIS
	Sends the log file via email.
.PARAMETER from
	Email address of sender
.PARAMETER to
	Email address of recipient
.PARAMETER attach
	Path to the log file to be sent as an attachment
.PARAMETER credentials
	Array containing strings with credential info. 0 = user name, 1 = path to password file, 2 = password file name
.PARAMETER gmail
	Switch, if used email is sent via Gmail rather than Outlook
.EXAMPLE
	$credentials = @("anthony.dunaway", "C:\Development\Password_Files", "gmail.txt")
	Send-Log -from "user@domain.com" -to "user@domain.com" -credentials $credentials -attach "C:\Development\script_log.txt" -gmail
.NOTES
	Requires Get_Credentials.ps1 function Get-MyCredentials. Outlook requires firewall access to port 587
	Credential parameter required to send the log. 
#>
	[cmdletbinding()]
	param(
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$from,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$to,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string]$attach,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[string[]]$credentials = @("anthony.dunaway", "C:\Development\Password_Files", "gmail.txt"),
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		$pscred = 0,
		[parameter(ValueFromPipelineByPropertyName,ValueFromPipeline)]
		[switch]$gmail
	)
	
	if($pscred -ne 0){
		$credential = $pscred
	}
	else{
		$credential = Get-MyCredentials -username $credentials[0] -filepath $credentials[1] -filename $credentials[2]
	}
	
	if($gmail){
		$server = "smtp.gmail.com"
	}
	else{
		$server = "smtp_mail.outlook.com"
	}
	
	$script = Get-CallingScriptName
	
	$message = @{
		SmtpServer = $server
		Credential = $credential
		Port = "587"
		UseSsl = $true
		From = $from
		To = $to
		Subject = "$script Log File Report"
		Body = "Attached is the log file for $script"
		Attachments = $attach
	}
	Send-MailMessage @message
	
}
