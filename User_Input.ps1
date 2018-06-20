function Get-UserInput{
	param([string]$Question)
	[string]$response = Read-Host $Question 
	$response = $response.ToLower()
	$response_split = $response.substring(0,1)
	If(($response_split -ne 'y') -and ($response_split -ne 'n')){
		While(($response_split -ne 'y') -and ($response_split -ne 'n')){
			$user_response = Read-Host 'Just a simple yes or no will do nicely thank you'
			$user_response = $user_response.ToLower()
			$response_split = $user_response.substring(0,1)
		}
	}
	If($response_split -eq 'y'){
		$decision = 1
	}
	Else{
		$decision = 0
	}
	return [int]$decision
}