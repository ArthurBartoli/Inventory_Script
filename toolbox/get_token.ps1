## Just in case

Import-Module MicrosoftPowerBIMgmt
Login-PowerBI

$headers = Get-PowerBIAccessToken
$token = $headers.Authorization

Write-Output $token