# By the date of submission, the Exchange Online hasn't deployed OAuth for POP3 in PROD. So this example may fail to authenticate. 

# Set-Location D:\dev\work\Exchange\PopImap
Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData.psm1
$userData = Get-UserData

# 3rd party module required to fetch AccessToken from Azure.
Import-Module MSAL.PS
$scopes =  @("https://outlook.office365.com/POP.AccessAsUser.All")
$msalToken = Get-MsalToken -ClientId $userData.clientId -TenantId $userData.tenantId -Scopes $scopes -Interactive

Import-Module .\src\PopImap\PopImap.psd1

$server = "outlook.office365.com"
$port = 995
$pop3 = Get-Pop3Client -Server $server -Port $port

$pop3.Connect()
$success = $pop3.O365Authenticate($msalToken.AccessToken, $userData.user)
if ($success)
{
  $pop3.ExecuteCommand("LIST")
}
$pop3.Close()
