<#
-----------------------------------------------------------------------
Name: 	Oregon Death Record Index Extractor
Author:	Anthony Dunaway
Date:     	05/21/2018
Updated: 	06/08/2018
Description:
Uses OCR to extract records from Oregon Death Index Software. The script takes screenshots
of the last name, first name, middle name, death date, and county columns. It then uses imagemagick to
perform preprocessing on the the images to prepare them for Tesseract OCR processing. Tesseract extracts
the text then outputs it to a text file. This file is appended to a master file, one for each column, then the script
advances the records by 40 to get the next set. 
-----------------------------------------------------------------------
#>
$location = Get-Location
."$location\Mouse_Click.ps1"
."$location\Auto_Screenshot.ps1"
."$location\User_Input.ps1"

#------------------------------------------------------------------------------------------------------------------------
# Common Tessract and Imagemagick commands, saved here for easy access
#------------------------------------------------------------------------------------------------------------------------
#- -c load_system_dawg=false
#- -c load_freq_dawg=false
#- -c classify_enable_learning=0
#- -c page_separator=' '
#- -c tessedit_char_whitelist=abcdefghijklmnopqrstuvwzy()
# -colorspace gray -type grayscale -contrast -background white -flatten +matte
# -c classify_enable_learning=0 -c load_freq_dawg=false -c load_system_dawg=false
# -morphology erode square:1 -colorspace gray -type grayscale
#------------------------------------------------------------------------------------------------------------------------
# Setup Tesseract command line arguments for each column
#------------------------------------------------------------------------------------------------------------------------
$tesseract = "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
$last_arguments = "--psm 6 --oem 1 -c classify_enable_learning=0 -c load_freq_dawg=false -c load_system_dawg=false $location\master_files\last.tif $location\master_files\last quiet"
$first_arguments = "--psm 6 --oem 1 -c classify_enable_learning=0 -c load_freq_dawg=false -c load_system_dawg=false $location\master_files\first.tif $location\master_files\first quiet"
$middle_arguments = "--psm 6 --oem 1 $location\master_files\middle.tif $location\master_files\middle quiet"
$date_arguments = "--psm 6 --oem 1 -c classify_enable_learning=0 -c load_freq_dawg=false -c load_system_dawg=false $location\master_files\date.tif $location\master_files\date quiet"

#------------------------------------------------------------------------------------------------------------------------
# Setup Imagemagick command line arguments for each column
#------------------------------------------------------------------------------------------------------------------------
$magick = "C:\Program Files\ImageMagick-7.0.7-Q16\magick.exe"
$last_spells = "convert -units PixelsPerInch $location\master_files\last.png -density 300 -resize 300% -colorspace gray -type grayscale -blur 1x1 -normalize -level 30% -contrast -sharpen 0x3.0 $location\master_files\last.tif"
$first_spells = "convert -units PixelsPerInch $location\master_files\first.png -density 300 -resize 200% -colorspace gray -type grayscale -blur 1x1 -normalize -level 30% -contrast -sharpen 0x3.0 $location\master_files\first.tif"
$middle_spells = "convert -units PixelsPerInch $location\master_files\middle.png -density 300 -resize 400% -contrast -colorspace gray -type grayscale -fuzz 30% -blur 1x1 -sharpen 0x5.0 -normalize -level 30% $location\master_files\middle.tif"
$date_spells = "convert -units PixelsPerInch $location\master_files\date.png -density 300 -resize 300% -colorspace gray -type grayscale -blur 1x1 -normalize -level 30% -contrast $location\master_files\date.tif"

#------------------------------------------------------------------------------------------------------------------------
# Setup script variables
#------------------------------------------------------------------------------------------------------------------------
Write-Host "ODI OCR Extractor"
$number_of_pages = Read-Host "How many pages would you like to process ?  "
$number_of_pages = [int]$number_of_pages
$run_middle = Get-UserInput -Question 'Do you want to capture middle names?        '
$run_first = Get-UserInput -Question 'Do you want to capture first names?         '
$num_columns = $run_first + $run_middle + 2
$first = "first"
$last = "last"
$middle = "middle"
$date = "date"
$master_first = "master_first.txt"
$master_last = "master_last.txt"
$master_middle = "master_middle.txt"
$master_date = "master_date.txt"
$save_path = "$location\master_files\"
$lname_bounds  = [Drawing.Rectangle]::FromLTRB(1921, 135, 2030, 935)
$fname_bounds  = [Drawing.Rectangle]::FromLTRB(2035, 135, 2145, 935)
$mname_bounds  = [Drawing.Rectangle]::FromLTRB(2149, 135, 2233, 935)
$date_bounds  = [Drawing.Rectangle]::FromLTRB(2238, 135, 2500, 935)
$start_time = Get-Date
#make sure ODI is activated
[Clicker]::LeftClickAtPoint(1950,100)
for($record = 1; $record -le $number_of_pages; $record ++){
	#These were used for testing
	# Start-Sleep -m 200
	# [Clicker]::LeftClickAtPoint(2360,950)
#------------------------------------------------------------------------------------------------------------------------
# Take screenshots of each column
#------------------------------------------------------------------------------------------------------------------------
	screenshot $lname_bounds "$save_path$last.png"
	if($run_first -eq 1){
		screenshot $fname_bounds "$save_path$first.png"
	}
	if($run_middle -eq 1){
		screenshot $mname_bounds "$save_path$middle.png"
	}
	screenshot $date_bounds "$save_path$date.png"

#------------------------------------------------------------------------------------------------------------------------
# Use Imagemagick to do preprocessing on the images to improve Tesseract accuracy
#------------------------------------------------------------------------------------------------------------------------
	Start-Process $magick -ArgumentList $last_spells -NoNewWindow -Wait
	if($run_first -eq 1){
		Start-Process $magick -ArgumentList $first_spells -NoNewWindow -Wait
	}
	if($run_middle -eq 1){
		Start-Process $magick -ArgumentList $middle_spells -NoNewWindow -Wait
	}
	Start-Process $magick -ArgumentList $date_spells -NoNewWindow -Wait
#------------------------------------------------------------------------------------------------------------------------
# Extract the text from the images with Tesseract
#------------------------------------------------------------------------------------------------------------------------
	Start-Process $tesseract -ArgumentList $last_arguments -NoNewWindow -Wait
	if($run_first -eq 1){
		Start-Process $tesseract -ArgumentList $first_arguments -NoNewWindow -Wait
	}
	if($run_middle -eq 1){
		Start-Process $tesseract -ArgumentList $middle_arguments -NoNewWindow -Wait
	}
	Start-Process $tesseract -ArgumentList $date_arguments -NoNewWindow -Wait

#------------------------------------------------------------------------------------------------------------------------
# Setup file lists to match user input
#------------------------------------------------------------------------------------------------------------------------	
	if(($run_first -eq 1) -and ($run_middle -eq 1)){
		$temp_files = ("$save_path$last.txt", "$save_path$first.txt", "$save_path$middle.txt", "$save_path$date.txt")
		$master_files = ("$save_path$master_last", "$save_path$master_first", "$save_path$master_middle", "$save_path$master_date")
		$files = ("$save_path$last.txt", "$save_path$first.txt", "$save_path$middle.txt", "$save_path$date.txt",
			"$save_path$last.png", "$save_path$first.png", "$save_path$middle.png", "$save_path$date.png",
			"$save_path$last.tif", "$save_path$first.tif", "$save_path$middle.tif", "$save_path$date.tif")
	}
	elseif(($run_first -eq 1) -and ($run_middle -eq 0)){
		$temp_files = ("$save_path$last.txt", "$save_path$first.txt", "$save_path$date.txt")
		$master_files = ("$save_path$master_last", "$save_path$master_first", "$save_path$master_date")
		$files = ("$save_path$last.txt", "$save_path$first.txt", "$save_path$date.txt",
			"$save_path$last.png", "$save_path$first.png", "$save_path$date.png",
			"$save_path$last.tif", "$save_path$first.tif", "$save_path$date.tif")
	}
	elseif(($run_first -eq 0) -and ($run_middle -eq 1)){
		$temp_files = ("$save_path$last.txt", "$save_path$middle.txt", "$save_path$date.txt")
		$master_files = ("$save_path$master_last", "$save_path$master_middle", "$save_path$master_date")
		$files = ("$save_path$last.txt", "$save_path$middle.txt", "$save_path$date.txt",
			"$save_path$last.png", "$save_path$middle.png", "$save_path$date.png",
			"$save_path$last.tif", "$save_path$middle.tif", "$save_path$date.tif")
	}
	else{
		$temp_files = ("$save_path$last.txt", "$save_path$date.txt")
		$master_files = ("$save_path$master_last", "$save_path$master_date")
		$files = ("$save_path$last.txt", "$save_path$date.txt",
			"$save_path$last.png", "$save_path$date.png",
			"$save_path$last.tif", "$save_path$date.tif")
	}

#------------------------------------------------------------------------------------------------------------------------
# Append Tesseract output to a master file for each column
#------------------------------------------------------------------------------------------------------------------------	
	for($i = 0; $i -lt $num_columns; $i++){
		$record_file = Get-Content $temp_files[$i]
		$file_stats = Get-Content  $temp_files[$i] | Measure-Object -Line
		$num_lines = $file_stats.Lines - 1
		Add-Content $master_files[$i] $record_file
		if($num_lines -lt 40){
			Add-Content $master_files[$i] "MISSING ENTRIES"
		}
		else{
			Add-Content $master_files[$i] "THERE WERE $num_lines Records"
		}
	}
#------------------------------------------------------------------------------------------------------------------------
# Delete temp files
#------------------------------------------------------------------------------------------------------------------------	
	foreach($file in $files){
		Remove-Item $file
	}
#------------------------------------------------------------------------------------------------------------------------
# Advance 40 records to get to the next page
#------------------------------------------------------------------------------------------------------------------------	
	if(($number_of_pages - $record)  -ge 1){
		for($clicks = 0; $clicks -lt 40; $clicks ++){
			[Clicker]::LeftClickAtPoint(3830, 1040)
			Start-Sleep -m 250
			[Clicker]::LeftClickAtPoint(1950,100)

		}
	}
	Write-Host "Processed Page #$record"
}
$end_time = Get-Date
$seconds = ($end_time - $start_time).TotalSeconds
Write-Host "Total time taken for $number_of_pages pages was $seconds seconds"
Read-Host -prompt 'press enter'
