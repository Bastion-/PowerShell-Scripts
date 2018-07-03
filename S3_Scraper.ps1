<#
-----------------------------------------------------------------------
Name: Heads Up Script Web Scraper
Author:	Anthony Dunaway
Date: 06/22/18
Updated: 06/22/18
Description:
Grabs information from the S3 website
-----------------------------------------------------------------------
#>

#----------------------------------------------------------------------------------------------------------------------
#Scrape the S3 RFC Page
#----------------------------------------------------------------------------------------------------------------------
function Get-S3Data{
	[cmdletbinding()]
	param(
		[string] $username = "username",
		[string] $password = "password",
		[parameter(mandatory=$true) ]
		[string] $rfc
	)
	Write-Verbose "Scraping S3 site for RFC $rfc info"
	$file_path = Get-Location
	$web_text = "web.txt"
	$url = "url"
	$url = $url + $rfc
	$ie = New-Object -com internetexplorer.application
	$ie.visible = $false
	$ie.navigate($url)
	while($ie.Busy -eq $true){
			Start-Sleep -s 1
	}
	if(![String]::IsNullOrEmpty($ie.document.getElementById("edit-name"))){
		Write-Verbose "Logging in to S3"
		$ie.document.getElementById("edit-name").Value = $username
		$ie.document.getElementById("edit-pass").Value = $password
		$submit = $ie.document.getElementById("edit-submit")
		$submit.click()
	}
	while($ie.Busy -eq $true){
			Start-Sleep -s 1
	}
	Write-Verbose "Grabbing the inner text"
	$web_content = $ie.document.getElementById("page").innerText
	Add-Content "$file_path\$web_text" $web_content
	#Return $web_content
	$ie.Quit()
	[GC]::collect()
}