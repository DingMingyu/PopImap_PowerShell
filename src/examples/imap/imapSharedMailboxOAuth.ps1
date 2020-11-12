# Set-Location D:\dev\work\Exchange\PopImap_PowerShell
Set-Location (Join-Path -Path $PSScriptRoot -ChildPath "..\..\..\")

# MyUserData.psm1 is not included in the project source control. Its content is similar to that of .\Data\UserData.psm1.
Import-Module .\Data\MyUserData.psm1
$userData = Get-UserData

# 3rd party module required to fetch AccessToken from Azure.
Import-Module MSAL.PS
$scopes =  @("https://outlook.office365.com/IMAP.AccessAsUser.All")
$msalToken = Get-MsalToken -ClientId $userData.clientId -TenantId $userData.tenantId -Scopes $scopes -Interactive

Import-Module .\src\PopImap\PopImap.psd1

$server = "outlook.office365.com"
$port = 993
$imap = Get-ImapClient -Server $server -Port $port

$imap.Connect()
$success = $imap.O365AuthenticateSharedMailbox($msalToken.AccessToken, $userData.user, $userData.sharedMailbox)
if ($success)
{
  $success = $imap.ExecuteCommand("list `"INBOX/`" *")
}
$imap.Close()
