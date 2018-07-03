<#
-----------------------------------------------------------------------
Name: Heads Up Script Mail Parser
Author:	Anthony Dunaway
Date: 06/22/18
Updated: 06/22/18
Description:
Determines what kind of notification a message is. 
-----------------------------------------------------------------------
#>
function Get-MessageType{
	[cmdletbinding()]
	param(
		[parameter(mandatory=$true) ]
		[string] $subject
	)
	
	Write-Verbose "Subject passed was $subject"
	$split_subject = $subject.Split(" ")
	foreach ($word in $split_subject){
		$word = $word.Trim()
		if($word -like "rfc*"){
			Write-Verbose "Found a single RFC"
			Return $word
		}
	}
	if(($split_subject[1] -like "ETS") -and ($split_subject[2] -like "Morning")){
		Write-Verbose "Found a morning update"
		Return "morning"
	}
	elseif($split_subject[1] -like "CAB"){
		Write-Verbose "Found a CAB agenda"
		Return "CAB"
	}
	else{
		Return "unknown"
	}
}