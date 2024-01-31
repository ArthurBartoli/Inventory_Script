Import-Module MicrosoftPowerBIMgmt
Login-PowerBI

$headers = Get-PowerBIAccessToken
$token = $headers.Authorization

echo $token