<#
-----------------------------------------------------------------------
Name: 	Crazy Pills Prank Script
Author:	Anthony Dunaway
Date:     	04/16/2018
Description:
Minimizes and maximizes everything on the user's screen between 1 and 3 times while mocking them.
It activates randomly at an interval between 5 and 10 minutes. Just cut and paste it into a PowerShell 
window and your unsuspecting victim will feel like they are taking crazy pills.
-----------------------------------------------------------------------
#>
#This is setup in an if statement so that it does not autorun on older versions of powershell. 
$run = 1
If($run -eq 1){
	Add-Type -Name Window -Namespace Console -MemberDefinition '
	[DllImport("Kernel32.dll")]
	public static extern IntPtr GetConsoleWindow();

	[DllImport("user32.dll")]
	public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
	'
	$consolePtr = [Console.Window]::GetConsoleWindow()
	[Console.Window]::ShowWindow($consolePtr, 0)
	$debug = 0
	Add-Type -AssemblyName System.speech 
	[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
	$shell = New-Object -ComObject "Shell.Application"
	While(1){
		$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
		#$speak.SelectVoice('Microsoft Zira Desktop')
		$trigger = Get-Random -Maximum 600 -Minimum 300
		$num_cycles = Get-Random -Maximum 3 -Minimum 1
		If($debug -eq 1){
			$trigger = 5
			$num_cycles = 2
		}
		Start-Sleep -s $trigger
		For($cycles = 1; $cycles -le $num_cycles; $cycles++){
			$shell.minimizeall()
			$speak.Speak('Oh No Oh No Oh No')
			Start-Sleep -s 3
			$shell.undominimizeall()
			$speak.Speak('Thats better')
		}
		$speak.Dispose()
	}
}
